import streamlit as st
import pandas as pd
import numpy as np
import math
import re
import io
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 画面のデザイン設定
st.set_page_config(page_title="製造計画自動スケジュールシステム", page_icon="🚜", layout="wide")

st.title("🚜 製造計画全自動スケジュールシステム (Excel完全対応版)")
st.markdown("### エクセルファイル（.xlsx）をそのまま置くだけで、日次・号機別のスケジュールを自動生成します")

st.sidebar.markdown("## 1. ファイルのアップロード")
file_zai = st.sidebar.file_uploader("① 在庫推移リスト (Excel形式: .xlsx)", type=["xlsx"])
file_gekkan = st.sidebar.file_uploader("② 本社 月間製造計画書 (Excel形式: .xlsx)", type=["xlsx"])
file_bom = st.sidebar.file_uploader("③ [任意] 新しいBOM構成表マスタ (ExcelまたはCSV)", type=["xlsx", "csv"])

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ プラント固定ルール")
st.sidebar.info(
    "・稼働時間: 415分/日\n"
    "・ロス率: 10% (投入量の90%を製品化)\n"
    "・最小バッチ: 5m3\n"
    "・基本バッチ: 10m3 (端数時は1バッチ追加)\n"
    "・同配合大→小連続製造: 切り替え5分に短縮"
)

# 複数シートから目的のシートを自動検知して読み込む関数
def load_excel_sheet_smart(file, keywords):
    xl = pd.ExcelFile(file)
    target_sheet = xl.sheet_names[0]
    for sheet in xl.sheet_names:
        if any(kw in sheet for kw in keywords):
            target_sheet = sheet
            break
    return pd.read_excel(xl, sheet_name=target_sheet, header=None)

# 2. 構成表マスタ（BOM）の保持・永続化ロジック
df_bom = None
if file_bom is not None:
    if file_bom.name.endswith('.csv'):
        try:
            df_bom = pd.read_csv(file_bom, encoding='utf-8')
        except UnicodeDecodeError:
            file_bom.seek(0)
            df_bom = pd.read_csv(file_bom, encoding='cp932')
    else:
        df_bom = load_excel_sheet_smart(file_bom, ["マスタ", "BOM", "BomMaster"])
    st.session_state['bom_data'] = df_bom
    st.sidebar.success("新しいマスタを一時読込しました")
elif os.path.exists("bom_master.xlsx"):
    df_bom = pd.read_excel("bom_master.xlsx")
elif os.path.exists("bom_master.csv"):
    try:
        df_bom = pd.read_csv("bom_master.csv", encoding='utf-8')
    except UnicodeDecodeError:
        df_bom = pd.read_csv("bom_master.csv", encoding='cp932')
elif 'bom_data' in st.session_state:
    df_bom = st.session_state['bom_data']

if df_bom is not None:
    st.sidebar.success("🟢 構成表マスタ: 読込済み (入力不要)")
else:
    st.sidebar.warning("⚠️ 構成表マスタが未登録です。GitHubに配置するか、ファイルを選択してください。")

if st.sidebar.button("🚀 製造計画スケジュールを生成する"):
    if not file_zai or not file_gekkan:
        st.error("エラー: ①在庫推移リスト と ②月間製造計画書 のエクセルファイルをアップロードしてください。")
    elif df_bom is None:
        st.error("エラー: 構成表マスタがシステム内に見つかりません。")
    else:
        with st.spinner("現在、エクセルから位置を自動検知し、スケジュールを演算中です..."):
            try:
                # 1. 在庫推移リストの読み込みと自動検知
                df_zai_raw = load_excel_sheet_smart(file_zai, ["在庫推移リスト", "在庫推移"])
                
                header_idx = None
                for i in range(len(df_zai_raw)):
                    row_vals = [str(v).strip() for v in df_zai_raw.iloc[i].values]
                    if any(kw in row_vals for kw in ['品目コード', '品目ｺｰﾄﾞ', '商品コード', '商品CD']):
                        header_idx = i
                        break
                
                if header_idx is None:
                    st.error("エラー: ①のファイル内に『品目コード』という見出し列が見つかりません。シート構成を確認してください。")
                    st.stop()

                raw_headers = [str(h).strip() for h in df_zai_raw.iloc[header_idx].values]
                
                standard_headers = []
                for h in raw_headers:
                    if h in ['品目コード', '品目ｺｰﾄﾞ', '商品コード', '商品CD']: standard_headers.append('品目コード')
                    elif h in ['品目名', '商品名']: standard_headers.append('品目名')
                    elif h in ['安全在庫数', '安全在庫']: standard_headers.append('安全在庫数')
                    elif h in ['種類', '区分']: standard_headers.append('種類')
                    else: standard_headers.append(h)

                df_zai_fixed = df_zai_raw.iloc[header_idx+1:].copy()
                df_zai_fixed.columns = standard_headers

                df_zai_fixed['品目コード'] = df_zai_fixed['品目コード'].ffill().astype(str).str.strip()
                df_zai_fixed['品目名'] = df_zai_fixed['品目名'].ffill().astype(str).str.strip()
                df_zai_fixed['安全在庫数'] = df_zai_fixed['安全在庫数'].ffill()

                df_zai_in_zai = df_zai_fixed[df_zai_fixed['種類'] == '在'].copy()
                df_zai_in_zai['安全在庫数'] = pd.to_numeric(df_zai_in_zai['安全在庫数'], errors='coerce')
                date_cols = [c for c in df_zai_in_zai.columns if '(日)' in str(c)]
                base_date = date_cols[0]
                df_zai_in_zai['現在の在庫'] = pd.to_numeric(df_zai_in_zai[base_date], errors='coerce')
                df_zai_in_zai['安全割れ不足数'] = df_zai_in_zai['安全在庫数'] - df_zai_in_zai['現在の在庫']
                df_zai_in_zai['安全割れ不足数'] = df_zai_in_zai['安全割れ不足数'].apply(lambda x: max(0, x))

                # 2. 月間製造計画書の読み込み
                df_monthly_raw = load_excel_sheet_smart(file_gekkan, ["本社 月間製造計画書", "月間製造計画書", "月間計画"])
                
                item_row_idx = None
                for i in range(min(15, len(df_monthly_raw))):
                    row_vals = [str(v).strip() for v in df_monthly_raw.iloc[i].values]
                    if any(kw in row_vals for kw in ['商品CD', '商品コード', '品目コード', '品目ｺｰﾄﾞ']):
                        item_row_idx = i
                        break
                if item_row_idx is None: item_row_idx = 1

                target_month_col = None
                for r in range(item_row_idx + 1):
                    row_vals = [str(v).strip() for v in df_monthly_raw.iloc[r].values]
                    for c_idx, val in enumerate(row_vals):
                        if '6月度' in val or '6月' in val:
                            target_month_col = c_idx
                            break
                    if target_month_col is not None: break
                
                plan_col_idx, actual_col_idx = None, None
                if target_month_col is not None:
                    for c in range(target_month_col, min(target_month_col + 8, len(df_monthly_raw.columns))):
                        col_val = str(df_monthly_raw.iloc[item_row_idx, c]).strip()
                        if '製造予定' in col_val: plan_col_idx = c
                        elif '製造実績' in col_val: actual_col_idx = c
                
                if plan_col_idx is None: plan_col_idx = 46
                if actual_col_idx is None: actual_col_idx = 47

                code_col_idx, name_col_idx = 0, 1
                for c_idx in range(min(5, len(df_monthly_raw.columns))):
                    val = str(df_monthly_raw.iloc[item_row_idx, c_idx]).strip()
                    if any(kw in val for kw in ['商品CD', '品目コード', 'コード']): code_col_idx = c_idx
                    elif any(kw in val for kw in ['商品名', '品目名', '名']): name_col_idx = c_idx

                df_m = df_monthly_raw.iloc[item_row_idx+1:].copy()
                df_m_clean = pd.DataFrame({
                    '品目コード': df_m.iloc[:, code_col_idx].astype(str).str.strip(),
                    '品目名_計画書': df_m.iloc[:, name_col_idx].astype(str).str.strip(),
                    '6月_製造予定': pd.to_numeric(df_m.iloc[:, plan_col_idx], errors='coerce').fillna(0),
                    '6月_製造実績': pd.to_numeric(df_m.iloc[:, actual_col_idx], errors='coerce').fillna(0)
                })
                df_m_clean['6月_計画残数'] = df_m_clean['6月_製造予定'] - df_m_clean['6月_製造実績']
                df_m_clean['6月_計画残数'] = df_m_clean['6月_計画残数'].apply(lambda x: max(0, x))
                df_m_distinct = df_m_clean[df_m_clean['品目コード'].notna() & (df_m_clean['品目コード'] != 'nan') & (df_m_clean['品目コード'] != '')].drop_duplicates(subset=['品目コード'])

                # 3. アウター合流（大きい方を自動採用）
                all_codes = set(df_zai_in_zai['品目コード']).union(set(df_m_distinct[df_m_distinct['6月_計画残数'] > 0]['品目コード']))
                
                master_list = []
                for code in all_codes:
                    if code in ['合計', 'nan', '商品CD', '品目コード', 'None']: continue
                    zai_row = df_zai_in_zai[df_zai_in_zai['品目コード'] == code]
                    plan_row = df_m_distinct[df_m_distinct['品目コード'] == code]
                    
                    name = zai_row['品目名'].iloc[0] if not zai_row.empty else (plan_row['品目名_計画書'].iloc[0] if not plan_row.empty else "不明")
                    safety_gap = zai_row['安全割れ不足数'].iloc[0] if not zai_row.empty else 0.0
                    current_stock = zai_row['現在の在庫'].iloc[0] if not zai_row.empty else np.nan
                    safety_stock = zai_row['安全在庫数'].iloc[0] if not zai_row.empty else np.nan
                    plan_gap = plan_row['6月_計画残数'].iloc[0] if not plan_row.empty else 0.0
                    
                    master_list.append({
                        '品目コード': code, '品目名': name, '現在の在庫': current_stock,
                        '安全在庫数': safety_stock, '安全割れ不足数': safety_gap, '今月の計画残数': plan_gap
                    })

                df_master_combined = pd.DataFrame(master_list)
                df_master_combined['採用ベース数量'] = df_master_combined[['安全割れ不足数', '今月の計画残数']].max(axis=1)
                df_master_combined = df_master_combined[df_master_combined['採用ベース数量'] > 0].copy()

                # 4. 容量Lの抽出
                def get_volume(name):
                    match = re.search(r'(\d+)\s*[LLｌｌＬＬ]', str(name))
                    if match: return int(match.group(1))
                    if '特大袋' in str(name): return 55
                    return 0

                df_master_combined['容量_L'] = df_master_combined['品目名'].apply(get_volume)
                df_master_combined['ベース必要容量_L'] = df_master_combined['採用ベース数量'] * df_master_combined['容量_L']

                # 5. マスタ列の自動認識
                df_bom.columns = [str(c).strip() for c in df_bom.columns]
                parent_col, child_col = None, None
                for c in df_bom.columns:
                    if c in ['商品CODE', '商品コード', '親品目コード', '親品目']: parent_col = c
                    elif c in ['配合CODE', '配合コード', '子品目コード', '子品目']: child_col = c
                
                if parent_col is None: parent_col = df_bom.columns[2] if len(df_bom.columns) > 2 else df_bom.columns[0]
                if child_col is None: child_col = df_bom.columns[0]
                
                def extract_content_code(item_code):
                    sub_bom = df_bom[df_bom[parent_col].astype(str).str.strip() == item_code]
                    if sub_bom.empty: return item_code
                    bh_items = sub_bom[sub_bom[child_col].astype(str).str.startswith('BH')]
                    if not bh_items.empty: return bh_items[child_col].iloc[0]
                    return sub_bom[child_col].iloc[0]

                df_master_combined['中身設計コード'] = df_master_combined['品目コード'].apply(extract_content_code)

                # 6. バッチ計算 (投入量に対して90%製品化)
                grouped = df_master_combined.groupby('中身設計コード').agg({'ベース必要容量_L': 'sum'}).reset_index()
                grouped['純計算_m3_ロス込'] = (grouped['ベース必要容量_L'] / 0.9) / 1000

                def apply_final_batch_rule(m3):
                    if m3 <= 0: return 0.0
                    if m3 <= 5.0: return 5.0
                    if m3 <= 10.0: return 10.0
                    return float(math.ceil(m3 / 10.0) * 10.0)

                grouped['製造決定_m3'] = grouped['純計算_m3_ロス込'].apply(apply_final_batch_rule)

                # 7. 最終指示袋数の確定 (製品化容量 = 決定バッチ × 0.9)
                df_final = df_master_combined.merge(grouped[['中身設計コード', '製造決定_m3']], on='中身設計コード', how='left')
                total_volume_by_recipe = df_final.groupby('中身設計コード')['ベース必要容量_L'].transform('sum')
                df_final['分配比率'] = df_final['ベース必要容量_L'] / total_volume_by_recipe
                df_final['分配比率'] = df_final['分配比率'].fillna(1.0)

                df_final['製品化容量_L'] = (df_final['製造決定_m3'] * 1000 * 0.9) * df_final['分配比率']
                df_final['計画製造袋数'] = (df_final['製品化容量_L'] / df_final['容量_L']).round().astype(int)

                # 8. 製造ライン判定
                def get_tf_flag(name):
                    return '腐葉土' in str(name) or '堆肥' in str(name) or '特大袋' in str(name)
                df_final['堆肥・腐葉土フラグ'] = df_final['品目名'].apply(get_tf_flag)

                def determine_line_advanced(row):
                    code = row['品目コード']
                    name = row['品目名']
                    vol = row['容量_L']
                    is_tf = row['堆肥・腐葉土フラグ']
                    if code == 'H0620030' or '再生材' in name or 'もう一土元気' in name: return '3号機'
                    if is_tf: return '3号機'
                    if vol <= 12: return '5号機'
                    if 14 <= vol <= 20: return '2号機'
                    if vol >= 25: return '6号機'
                    return '要確認'

                df_final['製造ライン'] = df_final.apply(determine_line_advanced, axis=1)

                # 9. スピードと所要時間の計算
                def calc_duration_mins(row):
                    line = row['製造ライン']
                    bags = row['計画製造袋数']
                    vol = row['容量_L']
                    if bags <= 0: return 0.0
                    if line == '2号機': speed = 500
                    elif line == '3号機': speed = 300
                    elif line == '5号機':
                        if vol == 14: speed = 1000
                        elif vol == 25: speed = 700
                        else: speed = 750
                    elif line == '6号機': speed = 300
                    else: speed = 400
                    return (bags / speed) * 60

                df_final['製造所要時間_分'] = df_final.apply(calc_duration_mins, axis=1)

                # 緊急度ソート
                def calc_urgency(row):
                    current = row['現在の在庫']
                    if pd.isna(current): return 500
                    if current <= 0: return current - 10000
                    return current - row['安全在庫数']

                df_final['緊急度'] = df_final.apply(calc_urgency, axis=1)
                group_urgency = df_final.groupby('中身設計コード')['緊急度'].min().to_dict()
                df_final['グループ緊急度'] = df_final['中身設計コード'].map(group_urgency)

                df_final_sorted = df_final[df_final['計画製造袋数'] > 0].sort_values(
                    by=['製造ライン', 'グループ緊急度', '中身設計コード', '容量_L'], ascending=[True, True, True, False]
                ).copy()

                # 10. 日次ハメ込みパズル
                full_schedule = []
                for line, group in df_final_sorted.groupby('製造ライン'):
                    current_day = 1
                    current_time_spent = 0.0
                    prev_recipe = None
                    prev_vol = None
                    
                    for idx, row in group.iterrows():
                        bags_left = row['計画製造袋数']
                        vol = row['容量_L']
                        recipe = row['中身設計コード']
                        
                        while bags_left > 0:
                            switch_time = 0.0
                            if current_time_spent > 0.0 and prev_recipe is not None:
                                if prev_recipe == recipe:
                                    if prev_vol is not None and prev_vol > vol: switch_time = 5.0
                                    else: switch_time = 10.0
                                else: switch_time = 10.0
                            
                            available_time = 415.0 - current_time_spent - switch_time
                            if available_time <= 5.0:
                                current_day += 1
                                current_time_spent = 0.0
                                prev_recipe = None
                                prev_vol = None
                                continue
                            
                            if line == '2号機': speed_per_min = 500 / 60
                            elif line == '3号機': speed_per_min = 300 / 60
                            elif line == '5号機':
                                if vol == 14: speed_per_min = 1000 / 60
                                elif vol == 25: speed_per_min = 700 / 60
                                else: speed_per_min = 750 / 60
                            elif line == '6号機': speed_per_min = 300 / 60
                            else: speed_per_min = 400 / 60
                            
                            max_bags_today = available_time * speed_per_min
                            
                            if bags_left <= max_bags_today:
                                bags_to_make = bags_left
                                job_duration = bags_to_make / speed_per_min
                                current_time_spent += switch_time + job_duration
                                
                                full_schedule.append({
                                    '稼働日': f"{current_day}日目", '製造ライン': line, '配合コード': recipe,
                                    '品目コード': row['品目コード'], '品目名': row['品目名'], '指示数量(袋)': int(bags_to_make),
                                    '製造時間(分)': round(job_duration, 1), '切り替え(分)': round(switch_time, 1),
                                    '合計拘束時間(分)': round(switch_time + job_duration, 1), '備考': '全量完了'
                                })
                                bags_left = 0
                                prev_recipe = recipe
                                prev_vol = vol
                            else:
                                bags_to_make = math.floor(max_bags_today)
                                if bags_to_make <= 0:
                                    current_day += 1
                                    current_time_spent = 0.0
                                    prev_recipe = None
                                    prev_vol = None
                                    continue
                                job_duration = bags_to_make / speed_per_min
                                full_schedule.append({
                                    '稼働日': f"{current_day}日目", '製造ライン': line, '配合コード': recipe,
                                    '品目コード': row['品目コード'], '品目名': row['品目名'], '指示数量(袋)': int(bags_to_make),
                                    '製造時間(分)': round(job_duration, 1), '切り替え(分)': round(switch_time, 1),
                                    '合計拘束時間(分)': round(switch_time + job_duration, 1), '備考': '翌日へ分割継続'
                                })
                                bags_left -= bags_to_make
                                current_day += 1
                                current_time_spent = 0.0
                                prev_recipe = None
                                prev_vol = None

                # 11. エクセル作成・出力
                wb = Workbook()
                wb.remove(wb.active)

                ws_summary = wb.create_sheet(title="製造品目・バッチ集計")
                ws_summary.views.sheetView[0].showGridLines = True
                ws_summary.append(["品目コード", "品目名", "製造ライン", "配合レシピ", "現在の在庫", "安全在庫数", "安全割れ不足数", "今月の計画残数", "決定製造m3", "最終製造総数(袋)"])
                for idx, row in df_final_sorted.iterrows():
                    ws_summary.append([
                        row['品目コード'], row['品目名'], row['製造ライン'], row['中身設計コード'],
                        row['現在の在庫'], row['安全在庫数'], row['安全割れ不足数'], row['今月の計画残数'],
                        row['製造決定_m3'], row['計画製造袋数']
                    ])

                ws_daily = wb.create_sheet(title="日別・号機別製造計画")
                ws_daily.views.sheetView[0].showGridLines = True
                ws_daily.append(["稼働日", "製造ライン", "配合コード", "品目コード", "品目名", "指示数量(袋)", "製造時間(分)", "切り替え(分)", "合計拘束時間(分)", "備考"])
                for job in full_schedule:
                    ws_daily.append([
                        job['稼働日'], job['製造ライン'], job['配合コード'], job['品目コード'], job['品目名'],
                        job['指示数量(袋)'], job['製造時間(分)'], job['切り替え(分)'], job['合計拘束時間(分)'], job['備考']
                    ])

                navy_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                zebra_fill = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
                white_font = Font(name="Meiryo UI", size=11, bold=True, color="FFFFFF")
                regular_font = Font(name="Meiryo UI", size=10)
                thin_side = Side(border_style="thin", color="D9D9D9")
                border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

                for ws in [ws_summary, ws_daily]:
                    ws.row_dimensions[1].height = 26
                    for cell in ws[1]:
                        cell.fill = navy_fill
                        cell.font = white_font
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    for row_idx in range(2, ws.max_row + 1):
                        ws.row_dimensions[row_idx].height = 20
                        is_zebra = (row_idx % 2 == 0)
                        for cell in ws[row_idx]:
                            cell.font = regular_font
                            cell.border = border_all
                            if is_zebra: cell.fill = zebra_fill
                            if isinstance(cell.value, (int, float)):
                                cell.number_format = "#,##0" if isinstance(cell.value, int) else "#,##0.0"
                                cell.alignment = Alignment(horizontal="right", vertical="center")
                            else:
                                cell.alignment = Alignment(horizontal="left", vertical="center")
                    for col in ws.columns:
                        max_len = 0
                        col_letter = get_column_letter(col[0].column)
                        for cell in col:
                            val_str = str(cell.value or '')
                            cell_len = sum([2 if ord(c) > 128 else 1 for c in val_str])
                            if cell_len > max_len: max_len = cell_len
                        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
                    ws.freeze_panes = "A2"

                excel_data = io.BytesIO()
                wb.save(excel_data)
                excel_data.seek(0)

                st.success("🎉 製造計画の作成が完了しました！以下のボタンから最新エクセルを保存してください。")
                st.download_button(
                    label="📊 製造指示スケジュール表(.xlsx)をダウンロード",
                    data=excel_data,
                    file_name="【確定完成版】当月日次製造指示スケジュール表.xlsx",
                    mime="application/vnd.openpyxlformats-officedocument.spreadsheetml.sheet"
                )

                st.markdown("### 🔍 明日（1日目）の予定プレビュー")
                df_preview = pd.DataFrame(full_schedule)
                st.dataframe(df_preview[df_preview['稼働日'] == '1日目'][['製造ライン', '品目名', '指示数量(袋)', '合計拘束時間(分)']], use_container_width=True)

            except Exception as e:
                st.error(f"エラーが発生しました。エクセルの形式（シート名や列位置）が正しいか確認してください。詳細: {str(e)}")
