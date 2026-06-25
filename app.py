import streamlit as st
import pandas as pd
import math
import re
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 画面のデザイン設定
st.set_page_config(page_title="製造計画自動スケジュールシステム", page_icon="🚜", layout="wide")

st.title("🚜 製造計画全自動スケジュールシステム")
st.markdown("### GEN ERPデータと月間計画書を置くだけで、日次の号機別スケジュールを自動生成します")

st.sidebar.markdown("## 1. ファイルのアップロード")
file_zai = st.sidebar.file_uploader("① 在庫推移リスト (CSV形式)", type=["csv"])
file_gekkan = st.sidebar.file_uploader("② 本社 月間製造計画書 (CSV形式)", type=["csv"])
file_bom = st.sidebar.file_uploader("③ BOM構成表マスタ (CSV形式)", type=["csv"])

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ プラント固定ルール")
st.sidebar.info(
    "・稼働時間: 415分/日\n"
    "・ロス率: 10% (投入量の90%を製品化)\n"
    "・最小バッチ: 5m3\n"
    "・基本バッチ: 10m3 (超えたら10m3追加)\n"
    "・大→小連続製造: 切り替え5分に短縮"
)

if st.sidebar.button("🚀 製造計画スケジュールを生成する"):
    if not (file_zai and file_gekkan and file_bom):
        st.error("エラー: 3つのCSVファイルをすべてアップロードしてください。")
    else:
        with st.spinner("現在、高度な配分パズルとカレンダーハメ込みを演算中です..."):
            try:
                # 1. 在庫推移リストの読み込み
                df_zai_raw = pd.read_csv(file_zai)
                headers = df_zai_raw.iloc[2].values
                df_zai_fixed = df_zai_raw.iloc[3:].copy()
                df_zai_fixed.columns = headers

                df_zai_fixed['品目コード'] = df_zai_fixed['品目コード'].ffill().str.strip()
                df_zai_fixed['品目名'] = df_zai_fixed['品目名'].ffill().str.strip()
                df_zai_fixed['安全在庫数'] = df_zai_fixed['安全在庫数'].ffill()

                df_zai_in_zai = df_zai_fixed[df_zai_fixed['種類'] == '在'].copy()
                df_zai_in_zai['安全在庫数'] = pd.to_numeric(df_zai_in_zai['安全在庫数'], errors='coerce')
                date_cols = [c for c in df_zai_in_zai.columns if '(日)' in str(c)]
                base_date = date_cols[0]
                df_zai_in_zai['現在の在庫'] = pd.to_numeric(df_zai_in_zai[base_date], errors='coerce')
                df_zai_in_zai['安全割れ不足数'] = df_zai_in_zai['安全在庫数'] - df_zai_in_zai['現在の在庫']
                df_zai_in_zai['安全割れ不足数'] = df_zai_in_zai['安全割れ不足数'].apply(lambda x: max(0, x))

                # 2. 月間製造計画書の読み込み
                df_monthly_raw = pd.read_csv(file_gekkan)
                df_m = df_monthly_raw.iloc[2:].copy()
                df_m_clean = pd.DataFrame({
                    '品目コード': df_m.iloc[:, 0].str.strip(),
                    '品目名_計画書': df_m.iloc[:, 1].str.strip(),
                    '6月_製造予定': pd.to_numeric(df_m.iloc[:, 46], errors='coerce').fillna(0),
                    '6月_製造実績': pd.to_numeric(df_m.iloc[:, 47], errors='coerce').fillna(0)
                })
                df_m_clean['6月_計画残数'] = df_m_clean['6月_製造予定'] - df_m_clean['6月_製造実績']
                df_m_clean['6月_計画残数'] = df_m_clean['6月_計画残数'].apply(lambda x: max(0, x))
                df_m_distinct = df_m_clean[df_m_clean['品目コード'].notna() & (df_m_clean['品目コード'] != '')].drop_duplicates(subset=['品目コード'])

                # 3. 統合（大きい方を採用）
                all_codes = set(df_zai_in_zai['品目コード']).union(set(df_m_distinct[df_m_distinct['6月_計画残数'] > 0]['品目コード']))
                master_list = []
                for code in all_codes:
                    if code == '合計': continue
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

                # 4. 容量の抽出
                def get_volume(name):
                    match = re.search(r'(\d+)\s*[LLｌｌＬＬ]', str(name))
                    if match: return int(match.group(1))
                    if '特大袋' in str(name): return 55
                    return 0

                df_master_combined['容量_L'] = df_master_combined['品目名'].apply(get_volume)
                df_master_combined['ベース必要容量_L'] = df_master_combined['採用ベース数量'] * df_master_combined['容量_L']

                # 5. BOMマスタから配合特定
                df_bom = pd.read_csv(file_bom)
                def extract_content_code(item_code):
                    sub_bom = df_bom[df_bom['親品目コード'] == item_code]
                    if sub_bom.empty: return item_code
                    bh_items = sub_bom[sub_bom['子品目コード'].astype(str).str.startswith('BH')]
                    if not bh_items.empty: return bh_items['子品目コード'].iloc[0]
                    sub_bom_numeric = sub_bom.copy()
                    sub_bom_numeric['員数'] = pd.to_numeric(sub_bom_numeric['員数'], errors='coerce')
                    small_items = sub_bom_numeric[sub_bom_numeric['員数'] < 1.0]
                    if not small_items.empty: return small_items['子品目コード'].iloc[0]
                    return sub_bom['子品目コード'].iloc[0]

                df_master_combined['中身設計コード'] = df_master_combined['品目コード'].apply(extract_content_code)

                # 6. 【最新：×0.9 ロス計算】プラントバッチ計算
                grouped = df_master_combined.groupby('中身設計コード').agg({'ベース必要容量_L': 'sum'}).reset_index()
                # 投入量に対して90%が製品化されるため、必要量 / 0.9 で逆算
                grouped['純計算_m3_ロス込'] = (grouped['ベース必要容量_L'] / 0.9) / 1000

                def apply_final_batch_rule(m3):
                    if m3 <= 0: return 0.0
                    if m3 <= 5.0: return 5.0
                    if m3 <= 10.0: return 10.0
                    return float(math.ceil(m3 / 10.0) * 10.0)

                grouped['製造決定_m3'] = grouped['純計算_m3_ロス込'].apply(apply_final_batch_rule)

                # 7. 最終予定袋数の確定（決定m3 * 1000 * 0.9 が実際の製品化L数）
                df_final = df_master_combined.merge(grouped[['中身設計コード', '製造決定_m3']], on='中身設計コード', how='left')
                total_volume_by_recipe = df_final.groupby('中身設計コード')['ベース必要容量_L'].transform('sum')
                df_final['分配比率'] = df_final['ベース必要容量_L'] / total_volume_by_recipe
                df_final['分配比率'] = df_final['分配比率'].fillna(1.0)

                df_final['製品化容量_L'] = (df_final['製造決定_m3'] * 1000 * 0.9) * df_final['分配比率']
                df_final['計画製造袋数'] = (df_final['製品化容量_L'] / df_final['容量_L']).round().astype(int)

                # 8. 製造ラインの判定
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

                # 9. 時間計算
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

                # 10. カレンダー割り当て
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

                # 11. openpyxlでのエクセル作成
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

                # メモリ上に出力してダウンロード
                excel_data = io.BytesIO()
                wb.save(excel_data)
                excel_data.seek(0)

                st.success("🎉 製造計画の作成が完了しました！以下のボタンからエクセルを保存してください。")
                st.download_button(
                    label="📊 【確定版】製造指示スケジュール表をダウンロード",
                    data=excel_data,
                    file_name="【確定版】当月日次製造指示スケジュール表.xlsx",
                    mime="application/vnd.openpyxlformats-officedocument.spreadsheetml.sheet"
                )

                # 画面上に簡易的なプレビューを表示
                st.markdown("### 🔍 1日目の製造指示（プレビュー）")
                df_preview = pd.DataFrame(full_schedule)
                st.dataframe(df_preview[df_preview['稼働日'] == '1日目'][['製造ライン', '品目名', '指示数量(袋)', '合計拘束時間(分)']], use_container_width=True)

            except Exception as e:
                st.error(f"計算中にエラーが発生しました。ファイルの形式が正しいか確認してください。エラー詳細: {str(e)}")
