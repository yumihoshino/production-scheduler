import streamlit st
import pandas as pd
import numpy as np
import math
import re
import io
import os
import copy
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 画面のデザイン設定
st.set_page_config(page_title="製造計画自動スケジュールシステム", page_icon="🚜", layout="wide")

st.title("🚜 製造計画全自動スケジュールシステム (カレンダー完全同期版)")
st.markdown("### エクセルを置くだけで、過去実績から【号機×商品別】のスピードを自動適用して最適化します")

st.sidebar.markdown("## 🏢 工場の選択")
factory_mode = st.sidebar.selectbox("対象の工場を選択してください", ["本社", "関西工場"])

st.sidebar.markdown("---")
st.sidebar.markdown("## 📅 カレンダー・目標設定")
target_days = st.sidebar.number_input("当月の目標稼働日数 (この日数以内に作り切る)", min_value=1, max_value=31, value=20)
target_month = st.sidebar.selectbox("計画対象の月度を選択してください", ["6月", "7月", "8月", "9月", "10月"])

# 作業日の翌日スタート＆祝日休業日の動的指定UI
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

# 工場ごとの解説テキスト切り替え
if factory_mode == "本社":
    rule_info = "・定時時間: 月〜木 430分(16:30終) / 金曜 400分(16:00終・メンテ)\n・稼働ライン: 2号機、3号機、5号機、6号機"
else:
    rule_info = "・定時時間: 月〜木 430分(16:30終) / 金曜 400分(16:00終・メンテ)\n・稼働ライン: 1号, 2号, 3号, 5号, 6号, その他\n・マスタ同期: 🌟関西独自の配合コード【BK】と変則見出しを100%完全名寄せ合流！"

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ 現場同期・固定ルール")
st.sidebar.info(
    f"・選択中の工場: {factory_mode}\n"
    f"・対象月度: {target_month}度計画\n"
    f"・開始日: {start_date.strftime('%Y/%m/%d')}\n"
    f"{rule_info}\n"
    "・永続ストレージ: 一度入れたマスタデータはリロードしても消えずに自動復元されます\n"
    "・残業最適化: 労務管理優先、必ず30分刻みジャストで終了探索\n"
    "・製造理由: [現在庫がマイナス] [安全在庫割れ] [計画未達] の3種仕分け\n"
    "・休憩ロック: 10:00(10分), 12:00(60分), 15:00(10分)"
)

# 複数シートから該当する計画書をすべて集めて縦に自動結合するスマート関数
def load_excel_sheets_merged(file, keywords):
    xl = pd.ExcelFile(file)
    matched_sheets = [sheet for sheet in xl.sheet_names if any(kw in sheet for kw in keywords)]
    if not matched_sheets:
        return pd.read_excel(xl, sheet_name=xl.sheet_names[0], header=None)
    
    base_df = pd.read_excel(xl, sheet_name=matched_sheets[0], header=None)
    item_row_idx = 1
    for i in range(min(15, len(base_df))):
        row_vals = [str(v).strip() for v in base_df.iloc[i].values]
        if any(k in row_vals for k in ['品目コード', '品目ｺｰﾄﾞ', '商品コード', '商品CD', '商品CODE']):
            item_row_idx = i; break
            
    for sheet in matched_sheets[1:]:
        add_df = pd.read_excel(xl, sheet_name=sheet, header=None)
        if len(add_df) > item_row_idx + 1:
            data_to_add = add_df.iloc[item_row_idx + 1:]
            base_df = pd.concat([base_df, data_to_add], ignore_index=True)
    return base_df

# 🌟【最重要機能：構成表マスタのどんな変則位置の見出しも自動検出して名寄せクリーンアップする防弾パーサー】
def clean_bom_master(df_raw_bom):
    if df_raw_bom is None or df_raw_bom.empty: return None
    h_row = 0
    for i in range(min(15, len(df_raw_bom))):
        row_vals = [str(v).strip() for v in df_raw_bom.iloc[i].values]
        if any(k in row_vals for k in ['品目コード', '商品コード', '商品CODE', '配合CODE', '配合コード']):
            h_row = i; break
    df_clean = df_raw_bom.iloc[h_row+1:].copy()
    df_clean.columns = [str(c).strip() for c in df_raw_bom.iloc[h_row].values]
    return df_clean

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

# マスタ読込・自動無条件復元ロジック
df_bom = None
if os.path.exists("bom_master_local.csv"):
    try: df_bom = pd.read_csv("bom_master_local.csv", encoding='utf-8')
    except UnicodeDecodeError: df_bom = pd.read_csv("bom_master_local.csv", encoding='cp932')
elif os.path.exists("bom_master.xlsx"): 
    df_bom = pd.read_excel("bom_master.xlsx", header=None)
    df_bom = clean_bom_master(df_bom)
elif os.path.exists("bom_master.csv"):
    try: df_bom = pd.read_csv("bom_master.csv", encoding='utf-8', header=None)
    except UnicodeDecodeError: df_bom = pd.read_csv("bom_master.csv", encoding='cp932', header=None)
    df_bom = clean_bom_master(df_bom)
elif 'bom_data' in st.session_state: df_bom = st.session_state['bom_data']

if file_bom is not None:
    if file_bom.name.endswith('.csv'):
        try: df_bom = pd.read_csv(file_bom, encoding='utf-8', header=None)
        except UnicodeDecodeError:
            file_bom.seek(0); df_bom = pd.read_csv(file_bom, encoding='cp932', header=None)
    else: df_bom = load_excel_sheets_merged(file_bom, ["マスタ", "BOM", "BomMaster", "ﾏｽﾀ"])
    
    df_bom = clean_bom_master(df_bom) # 👈 ここで正式に見出しクリーンアップ
    
    if df_bom is not None:
        df_bom.to_csv("bom_master_local.csv", index=False, encoding='utf-8')
        st.session_state['bom_data'] = df_bom

# 🌟【大進化：本社「BH」・関西「BK」のどちらから始まっても名寄せできる完全紐付け頭脳】
def extract_content_code(item_code):
    if df_bom is None or df_bom.empty: return item_code
    parent_col = None
    for c in df_bom.columns:
        if c in ['商品CODE', '商品コード', '品目コード', '商品CD']: parent_col = c; break
    if parent_col is None: parent_col = df_bom.columns[2] if len(df_bom.columns) > 2 else df_bom.columns[0]
    
    child_col = None
    for c in df_bom.columns:
        if c in ['配合CODE', '配合コード', '配合CD', '中身コード']: child_col = c; break
    if child_col is None: child_col = df_bom.columns[0]

    sub_bom = df_bom[df_bom[parent_col].astype(str).str.strip() == str(item_code).strip()]
    if sub_bom.empty: return item_code
    
    # 🌟 本社の'BH'、関西の'BK'のどちらから始まっても確実に検知！
    bh_items = sub_bom[sub_bom[child_col].astype(str).str.startswith(('BH', 'BK'))]
    return bh_items[child_col].iloc[0] if not bh_items.empty else sub_bom[child_col].iloc[0]

if df_bom is not None: st.sidebar.success("🟢 構成表マスタ: 読込済み (入力不要)")
else: st.sidebar.warning("⚠️ 構成表マスタが未登録です。")

if st.sidebar.button("🚀 製造計画スケジュールを生成する"):
    if not file_zai or not file_gekkan:
        st.error(f"エラー: 必要ファイルをアップロードしてください。")
    elif df_bom is None:
        st.error("エラー: 構成表マスタが見つかりません。")
    else:
        with st.spinner(f"現在、最新のライン別・商品別巡航スピードを同期させて最適化パズルを解いています..."):
            try:
                # 1. 在庫推移リストの読み込み
                df_zai_raw = load_excel_sheets_merged(file_zai, ["在庫推移リスト", "在庫推移"])
                header_idx = None
                for i in range(len(df_zai_raw)):
                    row_vals = [str(v).strip() for v in df_zai_raw.iloc[i].values]
                    if any(kw in row_vals for kw in ['品目コード', '品目ｺｰﾄﾞ', '商品コード', '商品CD', '商品CODE']):
                        header_idx = i; break
                if header_idx is None: st.error("エラー: 見出し列が見つかりません。"); st.stop()

                raw_headers = [str(h).strip() for h in df_zai_raw.iloc[header_idx].values]
                standard_headers = []
                for h in raw_headers:
                    if h in ['品目コード', '品目ｺｰﾄﾞ', '商品コード', '商品CD', '商品CODE']: standard_headers.append('品目コード')
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
                df_zai_in_zai['安全割れ不足数'] = (df_zai_in_zai['安全在庫数'] - df_zai_in_zai['現在の在庫']).apply(lambda x: max(0, x))

                # 2. 月間製造計画書の読み込み
                df_monthly_raw = load_excel_sheets_merged(file_gekkan, ["本社 月間製造計画書", "月間製造計画書", "月間計画", "本社"] if factory_mode == "本社" else ["関西工場 月間製造計画書", "関西工場", "関西製造計画", "計画", "月間製造計画書"])
                item_row_idx = 1
                for i in range(min(15, len(df_monthly_raw))):
                    row_vals = [str(v).strip() for v in df_monthly_raw.iloc[i].values]
                    if any(kw in row_vals for kw in ['商品CD', '商品コード', '品目コード', '品目ｺｰﾄﾞ', '品目ｃｄ', '商品CODE']):
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
                    if any(kw in val for kw in ['商品CD', '品目コード', 'コード', '商品', '商品CODE']): code_col_idx = c_idx
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
                    if code in ['合計', 'nan', '商品CD', '品目コード', 'None', '商品CODE']: continue
                    
                    # 関西工場モードのときだけ、Hから始まる本社商品コードを一発で完全無視して除外する
