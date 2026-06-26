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

st.title("🚜 製造計画全自動スケジュールシステム (現場完全同期版)")
st.markdown("### エクセルを置くだけで、ペア稼働・サイズ近接・1時間刻みのタイムテーブルを完全自動生成します")

st.sidebar.markdown("## 🏢 工場の選択")
factory_mode = st.sidebar.selectbox("対象の工場を選択してください", ["本社", "関西工場"])

st.sidebar.markdown("---")
st.sidebar.markdown("## 1. ファイルのアップロード")
file_zai = st.sidebar.file_uploader("① 在庫推移リスト (Excel形式: .xlsx)", type=["xlsx"])

if factory_mode == "本社":
    file_gekkan = st.sidebar.file_uploader("② 本社 月間製造計画書 (Excel形式: .xlsx)", type=["xlsx"])
else:
    file_gekkan = st.sidebar.file_uploader("② 関西工場 月間製造計画書 (Excel形式: .xlsx)", type=["xlsx"])

file_bom = st.sidebar.file_uploader("③ [任意] 新しいBOM構成表マスタ (ExcelまたはCSV)", type=["xlsx", "csv"])

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ 現場同期・固定ルール")
st.sidebar.info(
    f"・選択中の工場: {factory_mode}\n"
    "・定時稼働時間: 415分/日\n"
    "・ロス率: 10% (投入量の90%を製品化)\n"
    "・最小5m3 / 基本10m3 バッチルール\n"
    "・段取り替え: 異配合時は直前と「最も近いサイズ」を優先\n"
    "・人員制約: 全稼働を最優先、不可時は[2-6ペア] [3-5ペア]で連動\n"
    "・足切り: 欠品時を除き100袋未満は自動スキップ"
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

# 🌟【新ロジック：サイズ近接ソート関数】
def sort_jobs_by_size_proximity(df_line):
    unprocessed = df_line.to_dict('records')
    if not unprocessed: return []
    processed = []
    
    # 最初はキューの先頭（最も緊急度が高い配合グループ）
    first_job = unprocessed[0]
    first_recipe = first_job['中身設計コード']
    same_recipe_jobs = [j for j in unprocessed if j['中身設計コード'] == first_recipe]
    same_recipe_jobs.sort(key=lambda x: x['容量_L'], reverse=True) # 同配合内は大->小
    processed.extend(same_recipe_jobs)
    for j in same_recipe_jobs: unprocessed.remove(j)
        
    while unprocessed:
        last_job = processed[-1]
        last_vol = last_job['容量_L']
        
        # 残っている商品の中から、直前のサイズと「最も容量差が小さい」配合グループを貪欲に探索
        min_diff = float('inf')
        best_idx = -1
        for idx, j in enumerate(unprocessed):
            diff = abs(j['容量_L'] - last_vol)
            if diff < min_diff:
                min_diff = diff
                best_idx = idx
            elif diff == min_diff:
                # サイズ差が同じなら緊急度が高い方を優先
                if j['グループ緊急度'] < unprocessed[best_idx]['グループ緊急度']:
                    best_idx = idx
                    
        next_job = unprocessed[best_idx]
        next_recipe = next_job['中身設計コード']
        same_recipe_jobs = [j for j in unprocessed if j['中身設計コード'] == next_recipe]
        same_recipe_jobs.sort(key=lambda x: x['容量_L'], reverse=True)
        processed.extend(same_recipe_jobs)
        for j in same_recipe_jobs: unprocessed.remove(j)
        
    return processed

if st.sidebar.button("🚀 製造計画スケジュールを生成する"):
    if not file_zai or not file_gekkan:
        st.error(f"エラー: ①在庫推移リスト と ②{factory_mode}月間製造計画書 のエクセルファイルをアップロードしてください。")
    elif df_bom is None:
        st.error("エラー: 構成表マスタがシステム内に見つかりません。")
    else:
        with st.spinner("現在、サイズ近接段取りとオペレーターペア連動を考慮して24時間時間軸パズルを演算中です..."):
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
                    st.error("エラー: ①のファイル内に『品目コード』という見出し列が見つかりません。")
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

                # 2. 各工場の月間製造計画書の読み込み
                if factory_mode == "本社":
                    sheet_keywords = ["本社 月間製造計画書", "月間製造計画書", "月間計画", "本社"]
                else:
                    sheet_keywords = ["関西工場 月間製造計画書", "関西工場", "関西製造計画", "計画"]

                df_monthly_raw = load_excel_sheet_smart(file_gekkan, sheet_keywords)
                item_row_idx = None
                for i in range(min(15, len(df_monthly_raw))):
                    row_vals = [str(v).strip() for v in df_monthly_raw.iloc[i].values]
                    if any(kw in row_vals for kw in ['商品CD', '商品コード', '品目コード', '品目ｺｰﾄﾞ', '品目ｃｄ']):
                        item_row_idx = i
                        break
                if item_row_idx is None: item_row_idx = 1

                target_month_col = None
                for r in range(item_row_idx + 1):
                    row_vals = [str(v).strip() for v in df_monthly_raw.iloc[r].values]
                    for c_idx, val in enumerate(row_vals):
                        if '6月' in val or '当月' in val or '製造予定' in val:
                            target_month_col = c_idx
                            break
                    if target_month_col is not None: break
                
                plan_col_idx, actual_col_idx = None, None
                if target_month_col is not None:
                    for c in range(target_month_col, min(target_month_col + 8, len(df_monthly_raw.columns))):
                        col_val = str(df_monthly_raw.iloc[item_row_idx, c]).strip()
                        if '予定' in col_val or '製造予定' in col_val: plan_col_idx = c
                        elif '実績' in col_val or '製造実績' in col_val: actual_col_idx = c
                
                if plan_col_idx is None: plan_col_idx = 46
                if actual_col_idx is None: actual_col_idx = 47

                code_col_idx, name_col_idx = 0, 1
                for c_idx in range(min(5, len(df_monthly_raw.columns))):
                    val = str(df_monthly_raw.iloc[item_row_idx, c_idx]).strip()
                    if any(kw in val for kw in ['商品CD', '品目コード', 'コード', '商品']): code_col_idx = c_idx
                    elif any(kw in val for kw in ['商品名', '品目名', '名', '品名']): name_col_idx = c_idx

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

                # 3. アウター合流
                all_codes = set(df_zai_in_zai['品目コード']).union(set(df_m_distinct[df_m_distinct['6月_計画残数'] > 0]['品目コード']))
                master_list = []
                for code in all_codes:
                    if code in ['合計', 'nan', '商品CD', '品目コード', 'None', '商品ｃｄ']: continue
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
                df_master_combined['容量_L'] = df_master_combined['品目名'].apply(lambda n: int(re.search(r'(\d+)\s*[LLｌｌＬＬ]', str(n)).group(1)) if re.search(r'(\d+)\s*[LLｌｌＬＬ]', str(n)) else (55 if '特大袋' in str(n) else 0))
                df_master_combined['ベース必要容量_L'] = df_master_combined['採用ベース数量'] * df_master_combined['容量_L']

                # 5. マスタ列の自動認識して配合コード特定
                df_bom.columns = [str(c).strip() for c in df_bom.columns]
                parent_col = "商品CODE" if "商品CODE" in df_bom.columns else (df_bom.columns[2] if len(df_bom.columns) > 2 else df_bom.columns[0])
                child_col = "配合CODE" if "配合CODE" in df_bom.columns else df_bom.columns[0]
                
                def extract_content_code(item_code):
                    sub_bom = df_bom[df_bom[parent_col].astype(str).str.strip() == item_code]
                    if sub_bom.empty: return item_code
                    bh_items = sub_bom[sub_bom[child_col].astype(str).str.startswith('BH')]
                    return bh_items[child_col].iloc[0] if not bh_items.empty else sub_bom[child_col].iloc[0]

                df_master_combined['中身設計コード'] = df_master_combined['品目コード'].apply(extract_content_code)

                # 6. バッチ計算 (投入量に対して90%製品化)
                grouped = df_master_combined.groupby('中身設計コード').agg({'ベース必要容量_L': 'sum'}).reset_index()
                grouped['純計算_m3_ロス込'] = (grouped['ベース必要容量_L'] / 0.9) / 1000
                grouped['製造決定_m3'] = grouped['純計算_m3_ロス込'].apply(lambda m3: 5.0 if m3 <= 5.0 else (10.0 if m3 <= 10.0 else float(math.ceil(m3 / 10.0) * 10.0)))

                # 7. 最終指示袋数の確定
                df_final = df_master_combined.merge(grouped[['中身設計コード', '製造決定_m3']], on='中身設計コード', how='left')
                total_volume_by_recipe = df_final.groupby('中身設計コード')['ベース必要容量_L'].transform('sum')
                df_final['分配比率'] = (df_final['ベース必要容量_L'] / total_volume_by_recipe).fillna(1.0)
                df_final['製品化容量_L'] = (df_final['製造決定_m3'] * 1000 * 0.9) * df_final['分配比率']
                df_final['計画製造袋数'] = (df_final['製品化容量_L'] / df_final['容量_L']).round().astype(int)
                df_final['製造理由'] = df_final.apply(lambda r: '在庫不足（安全割れ）' if r['安全割れ不足数'] > 0 else '計画未達', axis=1)

                # 足切りルール
                df_final['計画製造袋数'] = df_final.apply(lambda r: 0 if r['製造理由'] == '計画未達' and r['計画製造袋数'] < 100 else r['計画製造袋数'], axis=1)

                # 8. 製造ライン判定
                df_final['堆肥・腐葉土フラグ'] = df_final['品目名'].apply(lambda n: '腐葉土' in str(n) or '堆肥' in str(n) or '特大袋' in str(n))
                def determine_line_advanced(row):
                    if row['品目コード'] == 'H0620030' or '再生材' in row['品目名'] or 'もう一土元気' in row['品目名'] or row['堆肥・腐葉土フラグ']: return '3号機'
                    if row['容量_L'] <= 12: return '5号機'
                    if 14 <= row['容量_L'] <= 20: return '2号機'
                    if row['容量_L'] >= 25: return '6号機'
                    return '要確認'
                df_final['製造ライン'] = df_final.apply(determine_line_advanced, axis=1)

                # 9. スピードと所要時間の計算
                def calc_duration_mins(row):
                    if row['計画製造袋数'] <= 0: return 0.0
                    if row['製造ライン'] == '2号機': speed = 500
                    elif row['製造ライン'] == '3号機': speed = 300
                    elif row['製造ライン'] == '5号機': speed = 1000 if row['容量_L'] == 14 else (700 if row['容量_L'] == 25 else 750)
                    elif row['製造ライン'] == '6号機': speed = 300
                    else: speed = 400
                    return (row['計画製造袋数'] / speed) * 60
                df_final['製造所要時間_分'] = df_final.apply(calc_duration_mins, axis=1)

                # ソート用緊急度
                df_final['緊急度'] = df_final.apply(lambda r: (r['現在の在庫'] - r['安全在庫数']) if not pd.isna(row['現在の在庫']) else 500, axis=1)
                group_urgency = df_final.groupby('中身設計コード')['緊急度'].min().to_dict()
                df_final['グループ緊急度'] = df_final['中身設計コード'].map(group_urgency)

                # 🌟【各号機ごとに、サイズ順＆大→小切り替え優先でソート】
                df_final_sorted = df_final[df_final['計画製造袋数'] > 0].sort_values(
                    by=['製造ライン', 'グループ緊急度', '中身設計コード', '容量_L'], ascending=[True, True, True, False]
                ).copy()

                # 🌟 10. 【最新機能：複数人対応・オペレーターペア稼働カレンダーハメ込み】
                queues = {}
                for line in ['2号機', '3号機', '5号機', '6号機']:
                    # ここで各ラインに「サイズ近接ソート」を適用！
                    line_df = df_final_sorted[df_final_sorted['製造ライン'] == line]
                    if not line_df.empty:
                        queues[line] = sort_jobs_by_size_proximity(line_df)
                    else:
                        queues[line] = []

                # 進捗ポインタと残り袋数の管理
                current_job_idx = {l: 0 for l in queues}
                for l in queues:
                    for j in queues[l]: j['remaining_bags'] = j['計画製造袋数']

                day = 1
                full_schedule = []

                while True:
                    # その日、各ラインに仕事が残っているかチェック
                    active_lines = []
                    for l in ['2号機', '3号機', '5号機', '6号機']:
                        idx = current_job_idx[l]
                        if idx < len(queues[l]) and queues[l][idx]['remaining_bags'] > 0:
                            active_lines.append(l)
                    
                    if not active_lines: break # 全ラインの仕事が終了したらシミュレーション完了

                    # ペア稼働の選別ロジック（4ライン揃うならフル、揃わないなら2-6、3-5のペア稼働）
                    has_pair_2_6 = ('2号機' in active_lines) or ('6号機' in active_lines)
                    has_pair_3_5 = ('3号機' in active_lines) or ('5号機' in active_lines)
                    
                    if '2号機' in active_lines and '3号機' in active_lines and '5号機' in active_lines and '6号機' in active_lines:
                        lines_to_run_today = ['2号機', '3号機', '5号機', '6号機']
                    elif has_pair_2_6 and not has_pair_3_5:
                        lines_to_run_today = [l for l in ['2号機', '6号機'] if l in active_lines]
                    elif has_pair_3_5 and not has_pair_2_6:
                        lines_to_run_today = [l for l in ['3号機', '5号機'] if l in active_lines]
                    else:
                        # 4ライン揃わないが、両方のペアに仕事が残っている場合、総残り袋数の多いペアを本日の稼働として優先
                        weight_2_6 = sum([queues[l][current_job_idx[l]]['remaining_bags'] for l in ['2号機', '6号機'] if l in active_lines])
                        weight_3_5 = sum([queues[l][current_job_idx[l]]['remaining_bags'] for l in ['3号機', '5号機'] if l in active_lines])
                        if weight_2_6 >= weight_3_5:
                            lines_to_run_today = [l for l in ['2号機', '6号機'] if l in active_lines]
                        else:
                            lines_to_run_today = [l for l in ['3号機', '5号機'] if l in active_lines]

                    # 選ばれたラインの時間を1日分（415分）流す
                    for line in lines_to_run_today:
                        time_spent = 0.0
                        prev_recipe = None
                        prev_vol = None
                        
                        while time_spent < 415.0:
                            idx = current_job_idx[line]
                            if idx >= len(queues[line]): break
                            job = queues[line][idx]
                            
                            switch_time = 0.0
                            if time_spent > 0.0 and prev_recipe is not None:
                                if prev_recipe == job['中身設計コード']:
                                    switch_time = 5.0 if (prev_vol and prev_vol > job['容量_L']) else 10.0
                                else:
                                    switch_time = 10.0
                            
                            available_time = 415.0 - time_spent - switch_time
                            if available_time <= 5.0: break
                            
                            vol = job['容量_L']
                            if line == '2号機': speed_per_min = 500 / 60
                            elif line == '3号機': speed_per_min = 300 / 60
                            elif line == '5号機': speed_per_min = 1000 / 60 if vol == 14 else (700 / 60 if vol == 25 else 750 / 60)
                            else: speed_per_min = 300 / 60
                            
                            max_bags_today = available_time * speed_per_min
                            start_time_current = time_spent + switch_time
                            
                            if job['remaining_bags'] <= max_bags_today:
                                bags_to_make = job['remaining_bags']
                                job_duration = bags_to_make / speed_per_min
                                time_spent += switch_time + job_duration
                                
                                full_schedule.append({
                                    '稼働日': f"{day}日目", '製造ライン': line, '配合コード': job['中身設計コード'],
                                    '品目コード': job['品目コード'], '品目名': job['品目名'], '指示数量(袋)': int(bags_to_make),
                                    '開始時間_分': start_time_current, '製造時間(分)': round(job_duration, 1),
                                    '切り替え(分)': round(switch_time, 1), '合計拘束時間(分)': round(switch_time + job_duration, 1), '備考': '全量完了'
                                })
                                job['remaining_bags'] = 0
                                current_job_idx[line] += 1
                                prev_recipe = job['中身設計コード']
                                prev_vol = job['容量_L']
                            else:
                                bags_to_make = math.floor(max_bags_today)
                                if bags_to_make <= 0: break
                                job_duration = bags_to_make / speed_per_min
                                time_spent += switch_time + job_duration
                                
                                full_schedule.append({
                                    '稼働日': f"{day}日目", '製造ライン': line, '配合コード': job['中身設計コード'],
                                    '品目コード': job['品目コード'], '品目名': job['品目名'], '指示数量(袋)': int(bags_to_make),
                                    '開始時間_分': start_time_current, '製造時間(分)': round(job_duration, 1),
                                    '切り替え(分)': round(switch_time, 1), '合計拘束時間(分)': round(switch_time + job_duration, 1), '備考': '翌日へ分割継続'
                                })
                                job['remaining_bags'] -= bags_to_make
                                prev_recipe = job['中身設計コード']
                                prev_vol = job['容量_L']
                                break
                    day += 1
                    if day > 40: break # 安全ブレーキ

                # 11. 高級エクセル自動装飾・出力
                wb = Workbook()
                wb.remove(wb.active)

                # シート①: 集計マスタ
                ws_summary = wb.create_sheet(title=f"製造品目・バッチ集計({factory_mode})")
                ws_summary.views.sheetView[0].showGridLines = True
                ws_summary.append(["品目コード", "品目名", "製造ライン", "配合レシピ", "現在の在庫", "安全在庫数", "安全割れ不足数", "今月の計画残数", "決定製造m3", "最終製造総数(袋)", "製造理由"])
                for idx, row in df_final_sorted.iterrows():
                    ws_summary.append([
                        row['品目コード'], row['品目名'], row['製造ライン'], row['中身設計コード'],
                        row['現在の在庫'], row['安全在庫数'], row['安全割れ不足数'], row['今月の計画残数'],
                        row['製造決定_m3'], row['計画製造袋数'], row['製造理由']
                    ])

                # シート②: 日次指示一覧
                ws_daily = wb.create_sheet(title=f"日別・号機別製造計画({factory_mode})")
                ws_daily.views.sheetView[0].showGridLines = True
                ws_daily.append(["稼働日", "製造ライン", "配合コード", "品目コード", "品目名", "指示数量(袋)", "製造時間(分)", "切り替え(分)", "合計拘束時間(分)", "備考"])
                for job in full_schedule:
                    ws_daily.append([
                        job['稼働日'], job['製造ライン'], job['配合コード'], job['品目コード'], job['品目名'],
                        job['指示数量(袋)'], job['製造時間(分)'], job['切り替え(分)'], job['合計拘束時間(分)'], job['備考']
                    ])

                # 🌟【シート③：2026.06形式 1時間刻みタイムテーブルシートの自動生成】
                ws_timeline = wb.create_sheet(title="日別時間軸タイムテーブル")
                ws_timeline.views.sheetView[0].showGridLines = True
                
                # 時間枠ヘッダー定義 (415分を1時間単位で7分割)
                time_slots = ["8:00〜", "9:00〜", "10:00〜", "11:00〜", "13:00〜", "14:00〜", "15:00〜"]
                max_days_generated = max([int(j['稼働日'].replace("日目", "")) for j in full_schedule]) if full_schedule else 1
                
                # 1〜2行目の日付と時間ヘッダーの組み立て
                row1 = ["予備", "号機"]
                row2 = ["", ""]
                for d in range(1, max_days_generated + 1):
                    row1.extend([f"{d}日目", "", "", "", "", "", ""])
                    row2.extend(time_slots)
                
                ws_timeline.append(row1)
                ws_timeline.append(row2)
                
                # 号機ごとのセル初期化
                line_map = {"2号機": "NO.2", "3号機": "NO.3", "5号機": "NO.5", "6号機": "NO.6"}
                timeline_rows = {l: ["", line_map[l]] + [""] * (max_days_generated * 7) for l in line_map}
                
                # ジョブを時間セルにマッピングしてハメ込む
                for job in full_schedule:
                    d_idx = int(job['稼働日'].replace("日目", "")) - 1
                    l_key = job['製造ライン']
                    if l_key not in timeline_rows: continue
                    
                    start_m = job['開始時間_分']
                    end_m = start_m + job['製造時間(分)']
                    
                    # 開始〜終了時間が属するスロット(0〜6)を割り出して文字を追記
                    for slot_idx in range(7):
                        slot_start = slot_idx * 60
                        slot_end = slot_start + 60 if slot_idx < 6 else 415.0
                        
                        # ジョブの時間がこの1時間枠に少しでも重なっているか
                        if max(start_m, slot_start) < min(end_m, slot_end):
                            cell_pos = 2 + (d_idx * 7) + slot_idx
                            txt = f"{job['品目名']}\n({job['指示数量(袋)']}袋)\n"
                            if timeline_rows[l_key][cell_pos] == "":
                                timeline_rows[l_key][cell_pos] = txt
                            else:
                                timeline_rows[l_key][cell_pos] += "＋\n" + txt
                
                for l in ["2号機", "3号機", "5号機", "6号機"]:
                    ws_timeline.append(timeline_rows[l])

                # スタイリングの適用
                navy_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                zebra_fill = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
                header_fill_tl = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                white_font = Font(name="Meiryo UI", size=11, bold=True, color="FFFFFF")
                regular_font = Font(name="Meiryo UI", size=10)
                bold_font = Font(name="Meiryo UI", size=10, bold=True)
                thin_side = Side(border_style="thin", color="D9D9D9")
                border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

                # 通常シートの装飾
                for ws in [ws_summary, ws_daily]:
                    ws.row_dimensions[1].height = 26
                    for cell in ws[1]:
                        cell.fill = navy_fill; cell.font = white_font; cell.alignment = Alignment(horizontal="center", vertical="center")
                    for row_idx in range(2, ws.max_row + 1):
                        ws.row_dimensions[row_idx].height = 20
                        is_zebra = (row_idx % 2 == 0)
                        for cell in ws[row_idx]:
                            cell.font = regular_font; cell.border = border_all
                            if is_zebra: cell.fill = zebra_fill
                            if isinstance(cell.value, (int, float)):
                                cell.number_format = "#,##0" if isinstance(cell.value, int) else "#,##0.0"
                                cell.alignment = Alignment(horizontal="right", vertical="center")
                            else:
                                cell.alignment = Alignment(horizontal="left", vertical="center")
                    for col in ws.columns:
                        max_len = max([sum([2 if ord(c) > 128 else 1 for c in str(cell.value or '')]) for cell in col])
                        ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 12)
                    ws.freeze_panes = "A2"

                # 🌟 タイムテーブルシート専用の超美麗装飾
                ws_timeline.row_dimensions[1].height = 24
                ws_timeline.row_dimensions[2].height = 20
                for cell in ws_timeline[1]:
                    cell.fill = navy_fill; cell.font = white_font; cell.alignment = Alignment(horizontal="center", vertical="center")
                for cell in ws_timeline[2]:
                    cell.fill = header_fill_tl; cell.font = bold_font; cell.alignment = Alignment(horizontal="center", vertical="center")
                
                # 日付セルのマージ結合 (7列ずつマージしてスッキリさせる)
                for d in range(max_days_generated):
                    start_col = 3 + (d * 7)
                    ws_timeline.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=start_col+6)

                for row_idx in range(3, ws_timeline.max_row + 1):
                    ws_timeline.row_dimensions[row_idx].height = 75 # 改行テキストが見えるように高さを広げる
                    for cell in ws_timeline[row_idx]:
                        cell.font = regular_font; cell.border = border_all
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True) # 自動改行
                
                # タイムテーブルの列幅設定
                ws_timeline.column_dimensions['A'].width = 8
                ws_timeline.column_dimensions['B'].width = 10
                for c in range(3, ws_timeline.max_column + 1):
                    ws_timeline.column_dimensions[get_column_letter(c)].width = 24

                excel_data = io.BytesIO()
                wb.save(excel_data)
                excel_data.seek(0)

                st.success(f"🎉 {factory_mode}用の現場完全同期計画が完了しました！以下のボタンからダウンロードしてください。")
                st.download_button(
                    label=f"📊 {factory_mode} 製造指示スケジュール表(.xlsx)をダウンロード",
                    data=excel_data,
                    file_name=f"【確定完成版】{factory_mode}_日次製造指示スケジュール表.xlsx",
                    mime="application/vnd.openpyxlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"エラーが発生しました。エクセルの形式が正しいか確認してください。詳細: {str(e)}")
