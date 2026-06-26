import streamlit as st
import pandas as pd
import numpy as np
import math
import re
import io
import os
import copy
import datetime

# 画面のデザイン設定
st.set_page_config(page_title="製造計画自動スケジュールシステム", page_icon="🚜", layout="wide")

st.title("🚜 製造計画全自動スケジュールシステム (カレンダー完全同期版)")
st.markdown("### エクセルを置くだけで、土日祝・金曜メンテ・30分刻みの横型ガント指示書を自動生成します")

st.sidebar.markdown("## 🏢 工場の選択")
factory_mode = st.sidebar.selectbox("対象の工場を選択してください", ["本社", "関西工場"])

st.sidebar.markdown("---")
st.sidebar.markdown("## 📅 カレンダー・目標設定")
target_days = st.sidebar.number_input("当月の目標稼働日数 (この日数以内に作り切る)", min_value=1, max_value=31, value=20)
target_month = st.sidebar.selectbox("計画対象の月度を選択してください", ["6月", "7月", "8月", "9月", "10月"])

# 🌟【新機能：作業日の翌日スタート＆祝日休業日の動的指定UI】
default_start = datetime.date.today() + datetime.timedelta(days=1)
start_date = st.sidebar.date_input("🚜 製造スケジュール開始日", default_start)
holidays_input = st.sidebar.multiselect(
    "🛑 平日の祝祭日・工場休業日を指定してください（自動スキップされます）",
    options=[start_date + datetime.timedelta(days=x) for x in range(45)],
    format_func=lambda x: x.strftime("%Y/%m/%d (%a)")
)

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
    f"・対象月度: {target_month}度計画\n"
    f"・開始日: {start_date.strftime('%Y/%m/%d')}\n"
    "・定時時間: 月〜木 430分(16:30終) / 金曜 400分(16:00終・メンテ)\n"
    "・残業最適化: 労務管理優先、必ず30分刻みジャストで終了探索\n"
    "・カレンダー: 土日および指定された休業日は自動的にスキップ\n"
    "・仕事融通：高速ライン(5号機等)が他ラインを自動応援製造\n"
    "・休憩ロック: 10:00(10分), 12:00(60分), 15:00(10分)"
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

# サイズ近接ソート関数
def sort_jobs_by_size_proximity(df_line):
    unprocessed = df_line.to_dict('records')
    if not unprocessed: return []
    processed = []
    first_job = unprocessed[0]
    first_recipe = first_job['中身設計コード']
    same_recipe_jobs = [j for j in unprocessed if j['中身設計コード'] == first_recipe]
    same_recipe_jobs.sort(key=lambda x: x['容量_L'], reverse=True) 
    processed.extend(same_recipe_jobs)
    for j in same_recipe_jobs: unprocessed.remove(j)
    while unprocessed:
        last_job = processed[-1]
        last_vol = last_job['容量_L']
        min_diff = float('inf')
        best_idx = -1
        for idx, j in enumerate(unprocessed):
            diff = abs(j['容量_L'] - last_vol)
            if diff < min_diff:
                min_diff = diff; best_idx = idx
            elif diff == min_diff:
                if j['グループ緊急度'] < unprocessed[best_idx]['グループ緊急度']: best_idx = idx
        next_job = unprocessed[best_idx]
        next_recipe = next_job['中身設計コード']
        same_recipe_jobs = [j for j in unprocessed if j['中身設計コード'] == next_recipe]
        same_recipe_jobs.sort(key=lambda x: x['容量_L'], reverse=True)
        processed.extend(same_recipe_jobs)
        for j in same_recipe_jobs: unprocessed.remove(j)
    return processed

# マスタ読込ロジック
df_bom = None
if os.path.exists("bom_master.xlsx"): df_bom = pd.read_excel("bom_master.xlsx")
elif os.path.exists("bom_master.csv"):
    try: df_bom = pd.read_csv("bom_master.csv", encoding='utf-8')
    except UnicodeDecodeError: df_bom = pd.read_csv("bom_master.csv", encoding='cp932')
elif 'bom_data' in st.session_state: df_bom = st.session_state['bom_data']

if file_bom is not None:
    if file_bom.name.endswith('.csv'):
        try: df_bom = pd.read_csv(file_bom, encoding='utf-8')
        except UnicodeDecodeError:
            file_bom.seek(0); df_bom = pd.read_csv(file_bom, encoding='cp932')
    else: df_bom = load_excel_sheet_smart(file_bom, ["マスタ", "BOM", "BomMaster"])
    st.session_state['bom_data'] = df_bom

def extract_content_code(item_code):
    if df_bom is None: return item_code
    parent_col = "商品CODE" if "商品CODE" in df_bom.columns else (df_bom.columns[2] if len(df_bom.columns) > 2 else df_bom.columns[0])
    child_col = "配合CODE" if "配合CODE" in df_bom.columns else df_bom.columns[0]
    sub_bom = df_bom[df_bom[parent_col].astype(str).str.strip() == item_code]
    if sub_bom.empty: return item_code
    bh_items = sub_bom[sub_bom[child_col].astype(str).str.startswith('BH')]
    return bh_items[child_col].iloc[0] if not bh_items.empty else sub_bom[child_col].iloc[0]

if df_bom is not None: st.sidebar.success("🟢 構成表マスタ: 読込済み (入力不要)")
else: st.sidebar.warning("⚠️ 構成表マスタが未登録です。")

if st.sidebar.button("🚀 製造計画スケジュールを生成する"):
    if not file_zai or not file_gekkan:
        st.error(f"エラー: 必要ファイルをアップロードしてください。")
    elif df_bom is None:
        st.error("エラー: 構成表マスタが見つかりません。")
    else:
        with st.spinner("現在、カレンダー・曜日・金曜メンテ・30分残業探索ルールを同期させてパズルを組み立てています..."):
            try:
                # 1. 在庫推移リストの読み込み
                df_zai_raw = load_excel_sheet_smart(file_zai, ["在庫推移リスト", "在庫推移"])
                header_idx = None
                for i in range(len(df_zai_raw)):
                    row_vals = [str(v).strip() for v in df_zai_raw.iloc[i].values]
                    if any(kw in row_vals for kw in ['品目コード', '品目ｺｰﾄﾞ', '商品コード', '商品CD']):
                        header_idx = i; break
                if header_idx is None: st.error("エラー: 見出し列が見つかりません。"); st.stop()

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

                df_zai_in_zai = df_zai_fixed[df_zai_fixed['種類'] == '放' or df_zai_fixed['種類'] == '在'].copy()
                df_zai_in_zai = df_zai_fixed[df_zai_fixed['種類'] == '在'].copy()
                df_zai_in_zai['安全在庫数'] = pd.to_numeric(df_zai_in_zai['安全在庫数'], errors='coerce')
                date_cols = [c for c in df_zai_in_zai.columns if '(日)' in str(c)]
                base_date = date_cols[0]
                df_zai_in_zai['現在の在庫'] = pd.to_numeric(df_zai_in_zai[base_date], errors='coerce')
                df_zai_in_zai['安全割れ不足数'] = (df_zai_in_zai['安全在庫数'] - df_zai_in_zai['現在の在庫']).apply(lambda x: max(0, x))

                # 2. 月間製造計画書の読み込み
                df_monthly_raw = load_excel_sheet_smart(file_gekkan, ["本社 月間製造計画書", "月間製造計画書", "月間計画", "本社"] if factory_mode == "本社" else ["関西工場 月間製造計画書", "関西工場", "関西製造計画", "計画"])
                item_row_idx = 1
                for i in range(min(15, len(df_monthly_raw))):
                    row_vals = [str(v).strip() for v in df_monthly_raw.iloc[i].values]
                    if any(kw in row_vals for kw in ['商品CD', '商品コード', '品目コード', '品目ｺｰﾄﾞ', '品目ｃｄ']):
                        item_row_idx = i; break

                plan_col_idx = None; actual_col_idx = None
                for r in range(item_row_idx + 1):
                    for c_idx in range(len(df_monthly_raw.columns)):
                        cell_val = str(df_monthly_raw.iloc[r, c_idx]).strip()
                        if target_month in cell_val:
                            for search_c in range(c_idx, min(c_idx + 6, len(df_monthly_raw.columns))):
                                col_text = "".join([str(df_monthly_raw.iloc[row, search_c]) for row in range(item_row_idx + 1)])
                                if ('予定' in col_text or '計画' in col_text) and plan_col_idx is None: plan_col_idx = search_c
                                elif '実績' in col_text and actual_col_idx is None: actual_col_idx = search_c

                if plan_col_idx is None or actual_col_idx is None:
                    try:
                        month_num = int(target_month.replace("月", ""))
                        base_offset = (month_num - 6) * 2
                        plan_col_idx = 46 + base_offset; actual_col_idx = 47 + base_offset
                    except: plan_col_idx, actual_col_idx = 46, 47

                code_col_idx, name_col_idx = (0, 1)
                for c_idx in range(min(5, len(df_monthly_raw.columns))):
                    val = str(df_monthly_raw.iloc[item_row_idx, c_idx]).strip()
                    if any(kw in val for kw in ['商品CD', '品目コード', 'コード', '商品']): code_col_idx = c_idx
                    elif any(kw in val for kw in ['商品名', '品目名', '名', '品名']): name_col_idx = c_idx

                df_m = df_monthly_raw.iloc[item_row_idx+1:].copy()
                df_m_clean = pd.DataFrame({
                    '品目コード': df_m.iloc[:, code_col_idx].astype(str).str.strip(),
                    '品目名_計画書': df_m.iloc[:, name_col_idx].astype(str).str.strip(),
                    '選択月_製造予定': pd.to_numeric(df_m.iloc[:, plan_col_idx], errors='coerce').fillna(0),
                    '選択月_製造実績': pd.to_numeric(df_m.iloc[:, actual_col_idx], errors='coerce').fillna(0)
                })
                df_m_clean['選択月_計画残数'] = (df_m_clean['選択月_製造予定'] - df_m_clean['選択月_製造実績']).apply(lambda x: max(0, x))
                df_m_distinct = df_m_clean[df_m_clean['品目コード'].notna() & (df_m_clean['品目コード'] != 'nan') & (df_m_clean['品目コード'] != '')].drop_duplicates(subset=['品目コード'])

                # 3. アウター合流
                all_codes = set(df_zai_in_zai['品目コード']).union(set(df_m_distinct[df_m_distinct['選択月_計画残数'] > 0]['品目コード']))
                master_list = []
                for code in all_codes:
                    if code in ['合計', 'nan', '商品CD', '品目コード', 'None']: continue
                    zai_row = df_zai_in_zai[df_zai_in_zai['品目コード'] == code]
                    plan_row = df_m_distinct[df_m_distinct['品目コード'] == code]
                    name = zai_row['品目名'].iloc[0] if not zai_row.empty else (plan_row['品目名_計画書'].iloc[0] if not plan_row.empty else "不明")
                    safety_gap = zai_row['安全割れ不足数'].iloc[0] if not zai_row.empty else 0.0
                    current_stock = zai_row['現在の在庫'].iloc[0] if not zai_row.empty else np.nan
                    safety_stock = zai_row['安全在庫数'].iloc[0] if not zai_row.empty else np.nan
                    plan_gap = plan_row['選択月_計画残数'].iloc[0] if not plan_row.empty else 0.0
                    
                    master_list.append({
                        '品目コード': code, '品目名': name, '現在の在庫': current_stock,
                        '安全在庫数': safety_stock, '安全割れ不足数': safety_gap, '今月の計画残数': plan_gap
                    })

                df_master_combined = pd.DataFrame(master_list)
                df_master_combined['採用ベース数量'] = df_master_combined[['安全割れ不足数', '今月の計画残数']].max(axis=1)
                df_master_combined = df_master_combined[df_master_combined['採用ベース数量'] > 0].copy()

                df_master_combined['容量_L'] = df_master_combined['品目名'].apply(lambda n: int(re.search(r'(\d+)\s*[LLｌｌＬＬ]', str(n)).group(1)) if re.search(r'(\d+)\s*[LLｌｌＬＬ]', str(n)) else (55 if '特大袋' in str(n) else 0))
                df_master_combined['ベース必要容量_L'] = df_master_combined['採用ベース数量'] * df_master_combined['容量_L']

                df_master_combined['中身設計コード'] = df_master_combined['品目コード'].apply(extract_content_code)

                grouped = df_master_combined.groupby('中身設計コード').agg({'ベース必要容量_L': 'sum'}).reset_index()
                grouped['純計算_m3_ロス込'] = (grouped['ベース必要容量_L'] / 0.9) / 1000
                grouped['製造決定_m3'] = grouped['純計算_m3_ロス込'].apply(lambda m3: 5.0 if m3 <= 5.0 else (10.0 if m3 <= 10.0 else float(math.ceil(m3 / 10.0) * 10.0)))

                df_final = df_master_combined.merge(grouped[['中身設計コード', '製造決定_m3']], on='中身設計コード', how='left')
                total_volume_by_recipe = df_final.groupby('中身設計コード')['ベース必要容量_L'].transform('sum')
                df_final['分配比率'] = (df_final['ベース必要容量_L'] / total_volume_by_recipe).fillna(1.0)
                df_final['製品化容量_L'] = (df_final['製造決定_m3'] * 1000 * 0.9) * df_final['分配比率']
                df_final['計画製造袋数'] = (df_final['製品化容量_L'] / df_final['容量_L']).round().astype(int)

                def determine_reason_advanced(row_item):
                    curr = row_item['現在の在庫']
                    if not pd.isna(curr) and curr < 0: return '現在庫がマイナス'
                    elif row_item['安全割れ不足数'] > 0: return '安全在庫割れ'
                    else: return '計画未達'
                df_final['製造理由'] = df_final.apply(determine_reason_advanced, axis=1)

                df_final['計画製造袋数'] = df_final.apply(lambda r: 0 if r['製造理由'] == '計画未達' and r['計画製造袋数'] < 100 else r['計画製造袋数'], axis=1)

                df_final['堆肥・腐葉土フラグ'] = df_final['品目名'].apply(lambda n: '腐葉土' in str(n) or '堆肥' in str(n) or '特大袋' in str(n))
                df_final['製造ライン'] = df_final.apply(lambda row_item: '3号機' if (row_item['品目コード'] == 'H0620030' or '再生材' in row_item['品目名'] or 'もう一土元気' in row_item['品目名'] or row_item['堆肥・腐葉土フラグ']) else ('5号機' if row_item['容量_L'] <= 12 else ('2号機' if 14 <= row_item['容量_L'] <= 20 else ('6号機' if row_item['容量_L'] >= 25 else '要確認'))), axis=1)

                df_final_sorted = df_final[df_final['計画製造袋数'] > 0].sort_values(by=['製造ライン', 'グループ緊急度', '中身設計コード', '容量_L'], ascending=[True, True, True, False]).copy()

                queues_base = {}
                for line in ['2号機', '3号機', '5号機', '6号機']:
                    line_df = df_final_sorted[df_final_sorted['製造ライン'] == line]
                    queues_base[line] = sort_jobs_by_size_proximity(line_df) if not line_df.empty else []

                # 日付・休日のヘルパー関数
                def get_next_working_date(current_date):
                    next_d = current_date
                    while True:
                        if next_d.weekday() >= 5 or next_d in holidays_input:
                            next_d += datetime.timedelta(days=1)
                        else:
                            break
                    return next_d

                # 🌟【カレンダー曜日連動対応型シミュレーションパズル関数】
                def run_calendar_simulation(overtime_block_mins):
                    queues = copy.deepcopy(queues_base)
                    current_job_idx = {l: 0 for l in queues}
                    for l in queues:
                        for j in queues[l]: j['remaining_bags'] = j['計画製造袋数']
                    
                    loop_date = get_next_working_date(start_date)
                    day_count = 1
                    schedule = []
                    
                    while True:
                        active_lines = []
                        for l in ['5号機', '2号機', '6号機', '3号機']:
                            has_work = False
                            if current_job_idx[l] < len(queues[l]) and queues[l][current_job_idx[l]]['remaining_bags'] > 0: has_work = True
                            else:
                                for ol in ['6号機', '2号機', '5号機', '3号機']:
                                    if ol == l: continue
                                    for job in queues[ol]:
                                        if job['remaining_bags'] > 0:
                                            if l == '5号機' and job['容量_L'] <= 25 and not job['堆肥・腐葉土フラグ']: has_work = True
                                            if l == '2号機' and job['容量_L'] <= 30 and not job['堆肥・腐葉土フラグ']: has_work = True
                            if has_work: active_lines.append(l)
                        
                        if not active_lines: break
                        
                        # 🌟【新機能：金曜・土曜のペア稼働・スピード連動ロジック】
                        has_pair_2_6 = ('2号機' in active_lines) or ('6号機' in active_lines)
                        has_pair_3_5 = ('3号機' in active_lines) or ('5号機' in active_lines)
                        
                        if '2号機' in active_lines and '3号機' in active_lines and '5号機' in active_lines and '6号機' in active_lines:
                            lines_to_run_today = ['2号機', '3号機', '5号機', '6号機']
                        elif has_pair_3_5 and ('5号機' in active_lines or not has_pair_2_6):
                            lines_to_run_today = [l for l in ['3号機', '5号機'] if l in active_lines]
                        else:
                            lines_to_run_today = [l for l in ['2号機', '6号機'] if l in active_lines]
                        
                        # 🌟【新機能：曜日による定時・残業容量の自動前倒し制御】
                        is_friday = (loop_date.weekday() == 4)
                        daily_base_capacity = 400.0 if is_friday else 430.0 # 金曜日は30分短いメンテ定時
                        capacity_limit_today = daily_base_capacity + overtime_block_mins
                        
                        weekday_kanji = ["月", "火", "水", "木", "金", "土", "日"][loop_date.weekday()]
                        date_str = loop_date.strftime("%Y/%m/%d")
                        
                        for line in lines_to_run_today:
                            time_spent = 0.0
                            prev_recipe, prev_vol = None, None
                            
                            while time_spent < capacity_limit_today:
                                idx = current_job_idx[line]
                                if idx < len(queues[line]):
                                    job = queues[line][idx]
                                    switch_time = 0.0
                                    if time_spent > 0.0 and prev_recipe is not None:
                                        switch_time = 5.0 if (prev_recipe == job['中身設計コード'] and prev_vol and prev_vol > job['容量_L']) else 10.0
                                    
                                    available_time = capacity_limit_today - time_spent - switch_time
                                    if available_time <= 5.0: break
                                    
                                    vol = job['容量_L']
                                    if line == '2号機': speed_per_min = 400 / 60
                                    elif line == '3号機': speed_per_min = (70 if vol == 55 else (100 if vol == 30 else 250)) / 60
                                    elif line == '5号機': speed_per_min = (730 if vol in [12, 14] else 650) / 60
                                    else: speed_per_min = 260 / 60
                                    
                                    max_bags_today = available_time * speed_per_min
                                    start_time_current = time_spent + switch_time
                                    
                                    if job['remaining_bags'] <= max_bags_today:
                                        bags_to_make = job['remaining_bags']
                                        job_duration = bags_to_make / speed_per_min
                                        t_start = start_time_current; t_end = t_start + job_duration
                                        time_spent += switch_time + job_duration
                                        schedule.append({
                                            '稼働日': f"{day_count}日目", '製造日': date_str, '曜日': weekday_kanji,
                                            '製造ライン': line, '配合コード': job['中身設計コード'],
                                            '品目コード': job['品目コード'], '品目名': job['品目名'], '指示数量(袋)': int(bags_to_make),
                                            '開始時間_分': start_time_current, '製造時間(分)': round(job_duration, 1),
                                            '切り替え(分)': round(switch_time, 1), '合計拘束時間(分)': round(switch_time + job_duration, 1), 
                                            '製造理由': job['製造理由'], '備考': '全量完了', 't_start': t_start, 't_end': t_end
                                        })
                                        job['remaining_bags'] = 0; current_job_idx[line] += 1
                                        prev_recipe, prev_vol = job['中身設計コード'], job['容量_L']
                                    else:
                                        bags_to_make = math.floor(max_bags_today)
                                        if bags_to_make <= 0: break
                                        job_duration = bags_to_make / speed_per_min
                                        t_start = start_time_current; t_end = t_start + job_duration
                                        time_spent += switch_time + job_duration
                                        schedule.append({
                                            '稼働日': f"{day_count}日目", '製造日': date_str, '曜日': weekday_kanji,
                                            '製造ライン': line, '配合コード': job['中身設計コード'],
                                            '品目コード': job['品目コード'], '品目名': job['品目名'], '指示数量(袋)': int(bags_to_make),
                                            '開始時間_分': start_time_current, '製造時間(分)': round(job_duration, 1),
                                            '切り替え(分)': round(switch_time, 1), '合計拘束時間(分)': round(switch_time + job_duration, 1), 
                                            '製造理由': job['製造理由'], '備考': '翌日へ分割継続', 't_start': t_start, 't_end': t_end
                                        })
                                        job['remaining_bags'] -= bags_to_make
                                        prev_recipe, prev_vol = job['中身設計コード'], job['容量_L']
                                        break
                                else:
                                    supported_job_found = False
                                    for other_line in ['6号機', '2号機', '5号機', '3号機']:
                                        if other_line == line: continue
                                        for job in queues[other_line]:
                                            if job['remaining_bags'] <= 0: continue
                                            can_support = False
                                            if line == '5号機' and job['容量_L'] <= 25 and not job['堆肥・腐葉土フラグ']: can_support = True
                                            if line == '2号機' and job['容量_L'] <= 30 and not job['堆肥・腐葉土フラグ']: can_support = True
                                            
                                            if can_support:
                                                switch_time = 10.0
                                                available_time = capacity_limit_today - time_spent - switch_time
                                                if available_time <= 5.0: break
                                                vol = job['容量_L']
                                                if line == '5号機': speed_per_min = (730 if vol in [12, 14] else 650) / 60
                                                else: speed_per_min = 400 / 60
                                                
                                                max_bags_today = available_time * speed_per_min
                                                start_time_current = time_spent + switch_time
                                                
                                                if job['remaining_bags'] <= max_bags_today:
                                                    bags_to_make = job['remaining_bags']
                                                    job_duration = bags_to_make / speed_per_min
                                                    t_start = start_time_current; t_end = t_start + job_duration
                                                    time_spent += switch_time + job_duration
                                                    schedule.append({
                                                        '稼働日': f"{day_count}日目", '製造日': date_str, '曜日': weekday_kanji,
                                                        '製造ライン': line, '配合コード': job['中身設計コード'],
                                                        '品目コード': job['品目コード'], '品目名': job['品目名'], '指示数量(袋)': int(bags_to_make),
                                                        '開始時間_分': start_time_current, '製造時間(分)': round(job_duration, 1),
                                                        '切り替え(分)': round(switch_time, 1), '合計拘束時間(分)': round(switch_time + job_duration, 1), 
                                                        '製造理由': job['製造理由'], '備考': f"★{other_line}の応援製造(全量完了)", 't_start': t_start, 't_end': t_end
                                                    })
                                                    job['remaining_bags'] = 0
                                                else:
                                                    bags_to_make = math.floor(max_bags_today)
                                                    if bags_to_make <= 0: break
                                                    job_duration = bags_to_make / speed_per_min
                                                    t_start = start_time_current; t_end = t_start + job_duration
                                                    time_spent += switch_time + job_duration
                                                    schedule.append({
                                                        '稼働日': f"{day_count}日目", '製造日': date_str, '曜日': weekday_kanji,
                                                        '製造ライン': line, '配合コード': job['中身設計コード'],
                                                        '品目コード': job['品目コード'], '品目名': job['品目名'], '指示数量(袋)': int(bags_to_make),
                                                        '開始時間_分': start_time_current, '製造時間(分)': round(job_duration, 1),
                                                        '切り替え(分)': round(switch_time, 1), '合計拘束時間(分)': round(switch_time + job_duration, 1), 
                                                        '製造理由': job['製造理由'], '備考': f"★{other_line}の応援製造(一部継続)", 't_start': t_start, 't_end': t_end
                                                    })
                                                    job['remaining_bags'] -= bags_to_make
                                                supported_job_found = True
                                                prev_recipe, prev_vol = job['中身設計コード'], job['容量_L']
                                                break
                                        if supported_job_found: break
                                    if not supported_job_found: break
                        
                        # 日付を進める
                        loop_date = get_next_working_date(loop_date + datetime.timedelta(days=1))
                        day_count += 1
                        if day_count > 45: break
                    return schedule, (day_count - 1)

                # 30分刻みの最小残業探索
                overtime_mins = 0
                full_schedule, generated_days = run_calendar_simulation(0)
                
                if generated_days > target_days:
                    for test_ov in [30, 60, 90, 120, 150, 180, 210]:
                        test_sched, test_days = run_calendar_simulation(test_ov)
                        if test_days <= target_days:
                            overtime_mins = test_ov
                            full_schedule = test_sched
                            generated_days = test_days
                            break

                if overtime_mins > 0:
                    st.warning(f"📢 【残業アラート】計画を目標の{target_days}日以内に終わらせるため、毎日一律 【{overtime_mins}分】の残業が必要です！")
                else:
                    st.success(f"🟢 【残業は一切不要です】金曜短縮・土日祝を考慮しても、通常の定時稼働のまま 【{generated_days}日間】ですべて安全に作り切れます！")

                # エクセル出力
                wb = Workbook()
                wb.remove(wb.active)

                ws_summary = wb.create_sheet(title="製造品目・バッチ集計")
                ws_summary.views.sheetView[0].showGridLines = True
                ws_summary.append(["品目コード", "品目名", "製造ライン", "配合レシピ", "現在の在庫", "安全在庫数", "安全割れ不足数", "今月の計画残数", "決定製造m3", "最終製造総数(袋)", "製造理由"])
                for idx, row_item in df_final_sorted.iterrows():
                    ws_summary.append([row_item['品目コード'], row_item['品目名'], row_item['製造ライン'], row_item['中身設計コード'], row_item['現在の在庫'], row_item['安全在庫数'], row_item['安全割れ不足数'], row_item['今月の計画残数'], row_item['製造決定_m3'], row_item['計画製造袋数'], row_item['製造理由']])

                ws_daily = wb.create_sheet(title="日別・号機別製造計画")
                ws_daily.views.sheetView[0].showGridLines = True
                # 🌟【新機能：見出し列に「製造日」と「曜日」を追加】
                ws_daily.append(["稼働日", "製造日", "曜日", "製造ライン", "配合コード", "品目コード", "品目名", "指示数量(袋)", "製造時間(分)", "切り替え(分)", "合計拘束時間(分)", "備考", "製造理由"])
                for job in full_schedule:
                    ws_daily.append([job['稼働日'], job['製造日'], job['曜日'], job['製造ライン'], job['配合コード'], job['品目コード'], job['品目名'], job['指示数量(袋)'], job['製造時間(分)'], job['切り替え(分)'], job['合計拘束時間(分)'], job['備考'], job['製造理由']])

                ws_timeline = wb.create_sheet(title="日別・30分刻みタイムテーブル")
                ws_timeline.views.sheetView[0].showGridLines = True
                
                time_slots = [
                    "8:00〜8:30", "8:30〜9:00", "9:00〜9:30", "9:30〜10:00", 
                    "10:00〜10:10(休憩)", "10:10〜10:30", "10:30〜11:00", "11:00〜11:30", "11:30〜12:00", 
                    "12:00〜13:00(昼休憩)", "13:00〜13:30", "13:30〜14:00", "14:00〜14:30", "14:30〜15:00", 
                    "15:00〜15:10(休憩)", "15:10〜15:30", "15:30〜16:00", "16:00〜16:30", 
                    "16:30〜17:00", "17:00〜17:30", "17:30〜18:00", "18:00〜18:30", "18:30〜19:00", "19:00〜19:30", "19:30〜20:00"
                ]
                # 🌟【新機能：タイムテーブルの見出しの先頭を拡張】
                ws_timeline.append(["稼働日", "製造日", "製造ライン"] + time_slots)
                
                # マトリクス作成用のユニークキーを抽出
                unique_days = []
                seen = set()
                for job in full_schedule:
                    k = (job['稼働日'], job['製造日'])
                    if k not in seen:
                        seen.add(k)
                        unique_days.append(k)
                
                matrix_rows = []
                for (d_str, date_str) in unique_days:
                    # 曜日を逆算
                    d_obj = datetime.datetime.strptime(date_str, "%Y/%m/%d")
                    w_kanji = ["月", "火", "水", "木", "金", "土", "日"][d_obj.weekday()]
                    for line in ["2号機", "3号機", "5号機", "6号機"]:
                        matrix_rows.append({
                            'day_str': d_str, 'date_disp': f"{date_str} ({w_kanji})", 'line_name': line,
                            'line_disp': {"2号機": "NO.2", "3号機": "NO.3", "5号機": "NO.5", "6号機": "NO.6"}[line],
                            'slots': [""] * 25
                        })
                
                slot_ranges = {}
                for s_idx in range(25):
                    if s_idx < 4: slot_ranges[s_idx] = (s_idx * 30, (s_idx + 1) * 30)
                    elif s_idx == 4: slot_ranges[s_idx] = (None, "小休憩")
                    elif s_idx == 5: slot_ranges[s_idx] = (120, 140)
                    elif s_idx < 9: slot_ranges[s_idx] = (140 + (s_idx - 6) * 30, 140 + (s_idx - 5) * 30)
                    elif s_idx == 9: slot_ranges[s_idx] = (None, "昼休憩")
                    elif s_idx < 14: slot_ranges[s_idx] = (230 + (s_idx - 10) * 30, 230 + (s_idx - 9) * 30)
                    elif s_idx == 14: slot_ranges[s_idx] = (None, "小休憩")
                    elif s_idx == 15: slot_ranges[s_idx] = (350, 370)
                    else: slot_ranges[s_idx] = (370 + (s_idx - 16) * 30, 370 + (s_idx - 15) * 30)

                for job in full_schedule:
                    d_str = job['稼働日']
                    l_key = job['製造ライン']
                    start_m = job['t_start']
                    end_m = job['t_end']
                    
                    for row_item in matrix_rows:
                        if row_item['day_str'] == d_str and row_item['line_name'] == l_key:
                            for s_idx in range(25):
                                if s_idx in [4, 14]: row_item['slots'][s_idx] = "小休憩"; continue
                                if s_idx == 9: row_item['slots'][s_idx] = "昼休憩"; continue
                                
                                s_start, s_end = slot_ranges[s_idx]
                                if s_idx == 24: s_end = 640.0
                                if s_start is not None and s_end is not None:
                                    if max(start_m, s_start) < min(end_m, s_end) - 1e-5:
                                        prefix = "★(応援)\n" if "応援製造" in str(job.get('備考', '')) else ""
                                        txt = f"{prefix}{job['品目名']}\n({job['指示数量(袋)']}袋)\n"
                                        if row_item['slots'][s_idx] == "": row_item['slots'][s_idx] = txt
                                        else: row_item['slots'][s_idx] += "＋\n" + txt

                for r_item in matrix_rows:
                    ws_timeline.append([r_item['day_str'], r_item['date_disp'], r_item['line_disp']] + r_item['slots'])

                # スタイリング
                navy_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                zebra_fill = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
                header_fill_tl = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                break_fill = PatternFill(start_color="E4DFEC", end_color="E4DFEC", fill_type="solid") 
                short_break_fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid") 
                white_font = Font(name="Meiryo UI", size=11, bold=True, color="FFFFFF")
                regular_font = Font(name="Meiryo UI", size=10); bold_font = Font(name="Meiryo UI", size=10, bold=True)
                thin_side = Side(border_style="thin", color="D9D9D9"); border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

                for ws in [ws_summary, ws_daily]:
                    ws.row_dimensions[1].height = 26
                    for cell in ws[1]: cell.fill = navy_fill; cell.font = white_font; cell.alignment = Alignment(horizontal="center", vertical="center")
                    for row_idx in range(2, ws.max_row + 1):
                        ws.row_dimensions[row_idx].height = 20
                        is_zebra = (row_idx % 2 == 0)
                        for cell in ws[row_idx]:
                            cell.font = regular_font; cell.border = border_all
                            if is_zebra: cell.fill = zebra_fill
                            if isinstance(cell.value, (int, float)):
                                cell.number_format = "#,##0"; cell.alignment = Alignment(horizontal="right", vertical="center")
                            else: cell.alignment = Alignment(horizontal="left", vertical="center")
                    for col in ws.columns:
                        max_len = max([sum([2 if ord(c) > 128 else 1 for c in str(cell.value or '')]) for cell in col])
                        ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 12)
                    ws.freeze_panes = "A2"

                ws_timeline.row_dimensions[1].height = 26
                for cell in ws_timeline[1]: cell.fill = navy_fill; cell.font = white_font; cell.alignment = Alignment(horizontal="center", vertical="center")
                
                # 縦セルの結合 🌟【タイポバグを100%きれいに消し去りました】
                for d in range(len(unique_days)):
                    ws_timeline.merge_cells(start_row=2+(d*4), start_column=1, end_row=2+(d*4)+3, end_column=1)
                    ws_timeline.merge_cells(start_row=2+(d*4), start_column=2, end_row=2+(d*4)+3, end_column=2)

                for row_idx in range(2, ws_timeline.max_row + 1):
                    ws_timeline.row_dimensions[row_idx].height = 65 
                    is_even_day = ((row_idx - 2) // 4 % 2 == 0)
                    for cell in ws_timeline[row_idx]:
                        cell.font = regular_font; cell.border = border_all
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        if cell.column in [2, 3]: cell.fill = header_fill_tl; cell.font = bold_font
                        elif cell.column in [8, 18]: cell.fill = short_break_fill; cell.font = bold_font 
                        elif cell.column == 13: cell.fill = break_fill; cell.font = bold_font 
                        elif cell.column > 3 and is_even_day: cell.fill = zebra_fill

                ws_timeline.column_dimensions['A'].width = 12; ws_timeline.column_dimensions['B'].width = 16; ws_timeline.column_dimensions['C'].width = 12
                for c in range(4, ws_timeline.max_column + 1): ws_timeline.column_dimensions[get_column_letter(c)].width = 24
                ws_timeline.freeze_panes = "D2" 

                excel_data = io.BytesIO()
                wb.save(excel_data)
                excel_data.seek(0)

                st.success(f"🎉 カレンダー・土日祝・金曜メンテ完全適応版のスケジュール表が完成しました！")
                st.download_button(
                    label="📊 製造指示スケジュール表(.xlsx)をダウンロード",
                    data=excel_data, file_name=f"【確定完成版】{target_month}度_カレンダー連動スケジュール表.xlsx",
                    mime="application/vnd.openpyxlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e: st.error(f"エラーが発生しました。詳細: {str(e)}")
