import streamlit as st
import pandas as pd
import numpy as np
import math
import re
import io
import os
import copy
import datetime
import gc
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="製造計画自動スケジュールシステム", page_icon="🚜", layout="wide")

st.title("🚜 製造計画全自動スケジュールシステム (カレンダー完全同期版)")
st.markdown("### エクセルを置くだけで、過去実績からスピードを学習し、超高速で指示書を出力します")

st.sidebar.markdown("## 🏢 工場の選択")
factory_mode = st.sidebar.selectbox("対象の工場を選択してください", ["本社", "関西工場"])

st.sidebar.markdown("---")
st.sidebar.markdown("## 📅 カレンダー・目標設定")
target_days = st.sidebar.number_input("当月の目標稼働日数 (この日数以内に作り切る)", min_value=1, max_value=31, value=20)
target_month = st.sidebar.selectbox("計画対象の月度を選択してください", ["6月", "7月", "8月", "9月", "10月"])

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

if factory_mode == "本社":
    rule_info = "・定時時間: 月〜木 430分(16:30終) / 金曜 400分(16:00終・メンテ)\n・稼働ライン: 2号機、3号機、5号機、6号機"
else:
    rule_info = "・定時時間: 月〜木 430分(16:30終) / 金曜 400分(16:00終・メンテ)\n・稼働ライン: 1号, 2号, 3号, 5号, 6号, その他\n・誤爆防止: 🌟古い本社マスタを強制看破して排除する関西専用検問ロック搭載！"

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ 現場同期・固定ルール")
st.sidebar.info(
    f"・選択中の工場: {factory_mode}\n"
    f"・対象月度: {target_month}度計画\n"
    f"・開始日: {start_date.strftime('%Y/%m/%d')}\n"
    f"{rule_info}\n"
    "・完全自動化: 人間による手動データ加工を一切排除した現場直結仕様\n"
    "・残業最適化: 労務管理優先、必ず30分刻みジャストで終了探索\n"
    "・製造理由: [現在庫がマイナス] [安全在庫割れ] [計画未達] の3種仕分け\n"
    "・休憩ロック: 10:00(10分), 12:00(60分), 15:00(10分)"
)

# =====================================================================
# 🌟 最上流グローバル・セーフティ関数群
# =====================================================================

def safe_seek(f):
    if hasattr(f, 'seek'): f.seek(0)

def load_excel_sheets_merged(file, keywords):
    safe_seek(file)
    xl = pd.ExcelFile(file)
    matched_sheets = [sheet for sheet in xl.sheet_names if any(kw in sheet for kw in keywords)]
    if not matched_sheets:
        safe_seek(file)
        df_single = pd.read_excel(file, sheet_name=0, header=None)
        return df_single
    
    base_df = pd.read_excel(xl, sheet_name=matched_sheets[0], header=None)
    item_row_idx = 1
    for i in range(min(15, len(base_df))):
        row_vals = [str(v).strip() for v in base_df.iloc[i].values]
        if any(k in row_vals for k in ['品目コード', '品目ｺｰﾄﾞ', '商品コード', '商品CD', '商品CODE']):
            item_row_idx = i; break
            
    for sheet in matched_sheets[1:]:
        add_df = pd.read_excel(xl, sheet_name=sheet, header=None)
        if len(add_df) > item_row_idx + 1:
            tmp_df = add_df.iloc[item_row_idx + 1:].copy()
            del add_df
            base_df = pd.concat([base_df, tmp_df], ignore_index=True)
            del tmp_df
            gc.collect()
            
    del xl
    gc.collect()
    return base_df

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

def extract_volume_safe(name_str):
    n_str = str(name_str)
    match = re.search(r'(\d+(?:\.\d+)?)\s*(?:[LLｌｌＬＬ]|[kKｋＫ][gGｇＧ]?)', n_str)
    if match:
        try: return int(float(match.group(1)))
        except: return 14
    elif '特大袋' in n_str: return 55
    else: return 14

def sort_jobs_by_size_proximity(df_line):
    unprocessed = df_line.to_dict('records')
    if not unprocessed: return []
    processed = []
    first_recipe = unprocessed[0]['中身設計コード']
    same_recipe_jobs = [j for j in unprocessed if j['中身設計コード'] == first_recipe]
    same_recipe_jobs.sort(key=lambda x: x['容量_L'], reverse=True) 
    processed.extend(same_recipe_jobs)
    for j in same_recipe_jobs: unprocessed.remove(j)
    while unprocessed:
        last_vol = processed[-1]['容量_L']
        min_diff = float('inf'); best_idx = -1
        for idx, j in enumerate(unprocessed):
            diff = abs(j['容量_L'] - last_vol)
            if diff < min_diff: min_diff = diff; best_idx = idx
            elif diff == min_diff:
                if j['グループ緊急度'] < unprocessed[best_idx]['グループ緊急度']: best_idx = idx
        next_recipe = unprocessed[best_idx]['中身設計コード']
        same_recipe_jobs = [j for j in unprocessed if j['中身設計コード'] == next_recipe]
        same_recipe_jobs.sort(key=lambda x: x['容量_L'], reverse=True)
        processed.extend(same_recipe_jobs)
        for j in same_recipe_jobs: unprocessed.remove(j)
    return processed

def job_can_support(l_key, job_item, f_mode):
    if f_mode == "本社": return (l_key == '5号機' and job_item['容量_L'] <= 25 or l_key == '2号機' and job_item['容量_L'] <= 30) and not job_item['堆肥・腐葉土フラグ']
    else: return (l_key == '5号機' and job_item['容量_L'] <= 14 or l_key == '2号機' and job_item['容量_L'] <= 25) and not job_item['堆肥・腐葉土フラグ']

def get_next_w_date(cur, holidays_list):
    nd = cur
    while nd.weekday() >= 5 or nd in holidays_list: nd += datetime.timedelta(days=1)
    return nd

def get_sp(line, vol, f_mode):
    if f_mode == "本社": return 400 if line == '2号機' else ((70 if vol == 55 else (100 if vol == 30 else 250)) if line == '3号機' else ((730 if vol in [12, 14] else 650) if line == '5号機' else 260))
    else: return 388 if line == '1号機' else (500 if line == '2号機' else ((70 if vol == 55 else (100 if vol == 30 else 191)) if line == '3号機' else (646 if line == '5号機' else (480 if line == '6号機' else 107))))

# =====================================================================
# 🌟 マスタスタンバイチェック
# =====================================================================

has_local_master = os.path.exists("bom_master_local.csv") or os.path.exists("bom_master.xlsx") or os.path.exists("bom_master.csv") or ('bom_data' in st.session_state) or (file_bom is not None)

has_master_in_gekkan = False
if not has_local_master and file_gekkan is not None:
    try:
        safe_seek(file_gekkan)
        xl_peek = pd.ExcelFile(file_gekkan)
        if any(any(k in s for k in ["マスタ", "BOM", "BomMaster", "ﾏｽﾀ"]) for s in xl_peek.sheet_names):
            has_master_in_gekkan = True
    except: pass

if has_local_master or has_master_in_gekkan:
    st.sidebar.success("🟢 構成表マスタ: 読込済み (スタンバイOK)")
else:
    st.sidebar.warning("⚠️ 構成表マスタが未登録です。")

# =====================================================================

if st.sidebar.button("🚀 製造計画スケジュールを生成する"):
    if not file_zai or not file_gekkan:
        st.error("エラー: 必要ファイルをアップロードしてください。")
    else:
        with st.spinner("⚡ 裏側でマスタを展開し、エコ・ハッシュエンジンで計画ファイルを出力中..."):
            try:
                df_bom = None
                if file_bom is not None:
                    safe_seek(file_bom)
                    if file_bom.name.endswith('.csv'):
                        try: df_bom = clean_bom_master(pd.read_csv(file_bom, encoding='utf-8', header=None))
                        except:
                            safe_seek(file_bom)
                            df_bom = clean_bom_master(pd.read_csv(file_bom, encoding='cp932', header=None))
                    else: df_bom = clean_bom_master(load_excel_sheets_merged(file_bom, ["マスタ", "BOM", "BomMaster", "ﾏｽﾀ"]))
                elif 'bom_data' in st.session_state:
                    df_bom = st.session_state['bom_data']
                elif os.path.exists("bom_master_local.csv"):
                    try: df_bom = pd.read_csv("bom_master_local.csv", encoding='utf-8')
                    except: df_bom = pd.read_csv("bom_master_local.csv", encoding='cp932')
                elif os.path.exists("bom_master.xlsx"):
                    df_bom = clean_bom_master(pd.read_excel("bom_master.xlsx", header=None))
                elif os.path.exists("bom_master.csv"):
                    try: df_bom = clean_bom_master(pd.read_csv("bom_master.csv", encoding='utf-8', header=None))
                    except: df_bom = clean_bom_master(pd.read_csv("bom_master.csv", encoding='cp932', header=None))

                # 🌟【ここが真の防波堤：関西マスタ誤爆チェック検問】
                # もし過去の古いマスタがロードされても、関西モードなのにBKコードが無ければ強制破棄！
                if df_bom is not None and not df_bom.empty and factory_mode == "関西工場":
                    has_bk = False
                    for col in df_bom.columns:
                        if df_bom[col].astype(str).str.contains('BK').any():
                            has_bk = True; break
                    if not has_bk:
                        df_bom = None  # 本社データの誤爆を検知して抹殺

                # 🌟 誤爆マスタが破棄されて空(None)になった場合、ここで計画書エクセル内から100%確実に本物の関西マスタを吸い上げる！
                if df_bom is None and file_gekkan is not None:
                    try:
                        safe_seek(file_gekkan)
                        xl_g = pd.ExcelFile(file_gekkan)
                        m_sheets = [s for s in xl_g.sheet_names if any(k in s for k in ["マスタ", "BOM", "BomMaster", "ﾏｽﾀ"])]
                        if m_sheets:
                            df_bom = clean_bom_master(pd.read_excel(xl_g, sheet_name=m_sheets[0], header=None))
                    except: pass

                if df_bom is not None:
                    try:
                        df_bom.to_csv("bom_master_local.csv", index=False, encoding='utf-8')
                        st.session_state['bom_data'] = df_bom
                    except: pass

                if df_bom is None:
                    st.error("エラー: 構成表マスタが見つかりません。")
                    st.stop()

                bom_lookup_dict = {}
                if not df_bom.empty:
                    p_col = next((c for c in df_bom.columns if c in ['商品CODE', '商品コード', '品目コード', '商品CD']), df_bom.columns[2] if len(df_bom.columns) > 2 else df_bom.columns[0])
                    c_col = next((c for c in df_bom.columns if c in ['配合CODE', '配合コード', '配合CD', '中身コード']), df_bom.columns[0])
                    for _, r in df_bom.iterrows():
                        pv = str(r[p_col]).strip(); cv = str(r[c_col]).strip()
                        if pv not in bom_lookup_dict or cv.startswith(('BH', 'BK')):
                            bom_lookup_dict[pv] = cv

                def extract_content_code(item_code):
                    return bom_lookup_dict.get(str(item_code).strip(), item_code)

                # --- 在庫推移リスト読み込み ---
                df_zai_raw = load_excel_sheets_merged(file_zai, ["在庫推移リスト", "在庫推移"])
                header_idx = next((i for i in range(len(df_zai_raw)) if any(kw in [str(v).strip() for v in df_zai_raw.iloc[i].values] for kw in ['品目コード', '品目ｺｰﾄﾞ', '商品コード', '商品CD', '商品CODE'])), None)
                if header_idx is None: st.error("エラー: 見出し列が見つかりません。"); st.stop()

                raw_headers = [str(h).strip() for h in df_zai_raw.iloc[header_idx].values]
                standard_headers = ['品目コード' if h in ['品目コード', '品目ｺｰﾄﾞ', '商品コード', '商品CD', '商品CODE'] else ('品目名' if h in ['品目名', '商品名'] else ('安全在庫数' if h in ['安全在庫数', '安全在庫'] else ('種類' if h in ['種類', '区分'] else h))) for h in raw_headers]

                df_zai_fixed = df_zai_raw.iloc[header_idx+1:].copy()
                df_zai_fixed.columns = standard_headers
                df_zai_fixed['品目コード'] = df_zai_fixed['品目コード'].ffill().astype(str).str.strip()
                df_zai_fixed['品目名'] = df_zai_fixed['品目名'].ffill().astype(str).str.strip()
                df_zai_fixed['安全在庫数'] = df_zai_fixed['安全在庫数'].ffill()

                df_zai_in_zai = df_zai_fixed[df_zai_fixed['種類'] == '在'].copy()
                df_zai_in_zai['安全在庫数'] = pd.to_numeric(df_zai_in_zai['安全在庫数'], errors='coerce')
                base_date = next((c for c in df_zai_in_zai.columns if '(日)' in str(c)), df_zai_in_zai.columns[-1])
                df_zai_in_zai['現在の在庫'] = pd.to_numeric(df_zai_in_zai[base_date], errors='coerce')
                df_zai_in_zai['安全割れ不足数'] = (df_zai_in_zai['安全在庫数'] - df_zai_in_zai['現在の在庫']).apply(lambda x: max(0, x))

                del df_zai_raw, df_zai_fixed
                gc.collect()

                # --- 月間計画書読み込み ---
                df_monthly_raw = load_excel_sheets_merged(file_gekkan, ["本社 月間製造計画書", "月間製造計画書", "月間計画", "本社"] if factory_mode == "本社" else ["関西工場 月間製造計画書", "関西工場", "関西製造計画", "計画", "月間製造計画書"])
                item_row_idx = next((i for i in range(min(15, len(df_monthly_raw))) if any(kw in [str(v).strip() for v in df_monthly_raw.iloc[i].values] for kw in ['商品CD', '商品コード', '品目コード', '品目ｺｰﾄﾞ', '品目ｃｄ', '商品CODE'])), 1)

                plan_col_idx = None; actual_col_idx = None
                for search_c in range(len(df_monthly_raw.columns)):
                    col_text = "".join([str(df_monthly_raw.iloc[row, search_c]) for range(item_row_idx + 1)])
                    if ('予定' in col_text or '計画' in col_text) and plan_col_idx is None: plan_col_idx = search_c
                    elif '実績' in col_text and actual_col_idx is None: actual_col_idx = search_c

                if plan_col_idx is None or actual_col_idx is None:
                    try: plan_col_idx = 46 + (int(target_month.replace("月", "")) - 6) * 2; actual_col_idx = plan_col_idx + 1
                    except: plan_col_idx, actual_col_idx = 46, 47

                code_col_idx = next((c for c in range(min(5, len(df_monthly_raw.columns))) if any(kw in str(df_monthly_raw.iloc[item_row_idx, c]) for kw in ['商品CD', '品目コード', 'コード', '商品', '商品CODE'])), 0)
                name_col_idx = next((c for c in range(min(5, len(df_monthly_raw.columns))) if any(kw in str(df_monthly_raw.iloc[item_row_idx, c]) for kw in ['商品名', '品目名', '名', '品名'])), 1)

                df_m = df_monthly_raw.iloc[item_row_idx+1:].copy()
                df_m_clean = pd.DataFrame({
                    '品目コード': df_m.iloc[:, code_col_idx].astype(str).str.strip(),
                    '品目名_計画書': df_m.iloc[:, name_col_idx].astype(str).str.strip(),
                    '選択月_製造予定': pd.to_numeric(df_m.iloc[:, plan_col_idx], errors='coerce').fillna(0),
                    '選択月_製造実績': pd.to_numeric(df_m.iloc[:, actual_col_idx], errors='coerce').fillna(0)
                })
                df_m_clean['選択月_計画残数'] = (df_m_clean['選択月_製造予定'] - df_m_clean['選択月_製造実績']).apply(lambda x: max(0, x))
                df_m_distinct = df_m_clean[df_m_clean['品目コード'].notna() & (~df_m_clean['品目コード'].isin(['nan', '', 'None']))].drop_duplicates(subset=['品目コード'])

                del df_monthly_raw, df_m, df_m_clean
                gc.collect()

                all_codes = set(df_zai_in_zai['品目コード']).union(set(df_m_distinct[df_m_distinct['選択月_計画残数'] > 0]['品目コード']))
                master_list = []
                for code in all_codes:
                    if code in ['合計', 'nan', '商品CD', '品目コード', 'None', '商品CODE'] or (factory_mode == "関西工場" and str(code).startswith('H')): continue
                    zai_row = df_zai_in_zai[df_zai_in_zai['品目コード'] == code]
                    plan_row = df_m_distinct[df_m_distinct['品目コード'] == code]
                    master_list.append({
                        '品目コード': code,
                        '品目名': zai_row['品目名'].iloc[0] if not zai_row.empty else (plan_row['品目名_計画書'].iloc[0] if not plan_row.empty else "不明"),
                        '現在の在庫': zai_row['現在の在庫'].iloc[0] if not zai_row.empty else np.nan,
                        '安全在庫数': zai_row['安全在庫数'].iloc[0] if not zai_row.empty else np.nan,
                        '安全割れ不足数': zai_row['安全割れ不足数'].iloc[0] if not zai_row.empty else 0.0,
                        '今月の計画残数': plan_row['選択月_計画残数'].iloc[0] if not plan_row.empty else 0.0
                    })

                df_master_combined = pd.DataFrame(master_list)
                if df_master_combined.empty: st.warning("計画対象となる品目がありません。"); st.stop()

                df_master_combined['採用ベース数量'] = df_master_combined[['安全割れ不足数', '今月の計画残数']].max(axis=1)
                df_master_combined = df_master_combined[df_master_combined['採用ベース数量'] > 0].copy()

                df_master_combined['容量_L'] = df_master_combined['品目名'].apply(extract_volume_safe)
                df_master_combined['ベース必要容量_L'] = df_master_combined['採用ベース数量'] * df_master_combined['容量_L']
                df_master_combined['中身設計コード'] = df_master_combined['品目コード'].apply(extract_content_code)

                grouped = df_master_combined.groupby('中身設計コード').agg({'ベース必要容量_L': 'sum'}).reset_index()
                grouped['純計算_m3_ロス込'] = (grouped['ベース必要容量_L'] / 0.9) / 1000
                grouped['製造決定_m3'] = grouped['純計算_m3_ロス込'].apply(lambda m3: 5.0 if m3 <= 5.0 else (10.0 if m3 <= 10.0 else float(math.ceil(m3 / 10.0) * 10.0)))

                df_final = df_master_combined.merge(grouped[['中身設計コード', '製造決定_m3']], on='中身設計コード', how='left')
                df_final['製造決定_m3'] = df_final['製造決定_m3'].fillna(0.0) 
                
                total_vol_recipe = df_final.groupby('中身設計コード')['ベース必要容量_L'].transform('sum')
                df_final['分配比率'] = (df_final['ベース必要容量_L'] / total_vol_recipe).fillna(1.0)
                df_final['製品化容量_L'] = ((df_final['製造決定_m3'] * 1000 * 0.9) * df_final['分配比率']).fillna(0.0)
                df_final['計画製造袋数'] = (df_final['製品化容量_L'] / df_final['容量_L']).replace([np.inf, -np.inf], np.nan).fillna(0.0).round().astype(int)

                df_final['製造理由'] = df_final.apply(lambda r: '現在庫がマイナス' if not pd.isna(r['現在の在庫']) and r['現在の在庫'] < 0 else ('安全在庫割れ' if r['安全割れ不足数'] > 0 else '計画未達'), axis=1)
                df_final['計画製造袋数'] = df_final.apply(lambda r: 0 if r['製造理由'] == '計画未達' and r['計画製造袋数'] < 100 else r['計画製造袋数'], axis=1)
                df_final['堆肥・腐葉土フラグ'] = df_final['品目名'].apply(lambda n: any(k in str(n) for k in ['腐葉土', '堆肥', '特大袋']))

                if factory_mode == "本社":
                    df_final['製造ライン'] = df_final.apply(lambda r: '3号機' if r['品目コード'] == 'H0620030' or any(k in r['品目名'] for k in ['再生材', 'もう一土元気']) or r['堆肥・腐葉土フラグ'] else ('5号機' if r['容量_L'] <= 12 else ('2号機' if r['容量_L'] <= 20 else '6号機')), axis=1)
                else:
                    df_final['製造ライン'] = df_final.apply(lambda r: '3号機' if any(k in r['品目名'] for k in ['再生材', 'もう一土元気']) or r['堆肥・腐葉土フラグ'] else ('5号機' if r['容量_L'] <= 12 else ('1号機' if r['容量_L'] >= 25 else ('6号機' if r['容量_L'] <= 20 else '2号機'))), axis=1)

                df_final['製造所要時間_分'] = df_final.apply(lambda r: (r['計画製造袋数'] / get_sp(r['製造ライン'], r['容量_L'], factory_mode)) * 60 if r['計画製造袋数'] > 0 else 0.0, axis=1)
                df_final['緊急度'] = df_final.apply(lambda r: (r['現在の在庫'] - r['安全在庫数']) if not pd.isna(r['現在の在庫']) else 500, axis=1)
                df_final['グループ緊急度'] = df_final['中身設計コード'].map(df_final.groupby('中身設計コード')['緊急度'].min().to_dict())

                df_final_sorted = df_final[df_final['計画製造袋数'] > 0].sort_values(by=['製造ライン', 'グループ緊急度', '中身設計コード', '容量_L'], ascending=[True, True, True, False]).copy()

                del df_master_combined, df_final, grouped
                gc.collect()

                lines_list = ["1号機", "2号機", "3号機", "5号機", "6号機", "その他"] if factory_mode == "関西工場" else ["2号機", "3号機", "5号機", "6号機"]
                queues_base = {line: sort_jobs_by_size_proximity(df_final_sorted[df_final_sorted['製造ライン'] == line]) for line in lines_list}

                def run_sim(ov_mins):
                    queues = copy.deepcopy(queues_base)
                    cur_idx = {l: 0 for l in queues}
                    for l in queues:
                        for j in queues[l]: j['rem'] = j['計画製造袋数']

                    loop_d = get_next_w_date(start_date, holidays_input)
                    day_cnt = 1; sched = []
                    
                    while True:
                        active = [l for l in lines_list if cur_idx[l] < len(queues[l]) and queues[l][cur_idx[l]]['rem'] > 0 or any(j['rem'] > 0 and job_can_support(l, j, factory_mode) for ol in lines_list if ol != l for j in queues[ol])]
                        if not active: break
                        
                        run_today = active if factory_mode == "関西工場" else ([l for l in ['2号機', '3号機', '5号機', '6号機'] if l in active] if len(active) == 4 else ([l for l in ['3号機', '5号機'] if l in active] if any(l in active for l in ['3号機', '5号機']) and ('5号機' in active or not any(l in active for l in ['2号機', '6号機'])) else [l for l in ['2号機', '6号機'] if l in active]))
                        
                        cap_limit = (400.0 if loop_d.weekday() == 4 else 430.0) + ov_mins
                        w_kanji = ["月", "火", "水", "木", "金", "土", "日"][loop_d.weekday()]
                        d_str_disp = loop_d.strftime("%Y/%m/%d")
                        
                        for line in run_today:
                            spent = 0.0; p_rec = None; p_vol = None
                            while spent < cap_limit:
                                idx = cur_idx[line]
                                if idx < len(queues[line]):
                                    job = queues[line][idx]
                                    sw = 5.0 if spent > 0 and p_rec == job['中身設計コード'] and p_vol and p_vol > job['容量_L'] else (10.0 if spent > 0 else 0.0)
                                    avail = cap_limit - spent - sw
                                    if avail <= 5.0: break
                                    
                                    sp_min = get_sp(line, job['容量_L'], factory_mode) / 60
                                    max_b = avail * sp_min
                                    
                                    if job['rem'] <= max_b:
                                        b_make = job['rem']; dur = b_make / sp_min
                                        sched.append({'稼働日': f"{day_cnt}日目", '製造日': d_str_disp, '曜日': w_kanji, '製造ライン': line, '配合コード': job['中身設計コード'], '品目コード': job['品目コード'], '品目名': job['品目名'], '指示数量(袋)': int(b_make), '製造時間(分)': round(dur, 1), '切り替え(分)': round(sw, 1), '合計拘束時間(分)': round(sw + dur, 1), '備考': '全量完了', '製造理由': job['製造理由'], 't_start': spent + sw, 't_end': spent + sw + dur})
                                        spent += sw + dur; job['rem'] = 0; cur_idx[line] += 1
                                        p_rec = job['中身設計コード']; p_vol = job['容量_L']
                                    else:
                                        b_make = math.floor(max_b)
                                        if b_make <= 0: break
                                        dur = b_make / sp_min
                                        sched.append({'稼働日': f"{day_cnt}日目", '製造日': d_str_disp, '曜日': w_kanji, '製造ライン': line, '配合コード': job['中身設計コード'], '品目コード': job['品目コード'], '品目名': job['品目名'], '指示数量(袋)': int(b_make), '製造時間(分)': round(dur, 1), '切り替え(分)': round(sw, 1), '合計拘束時間(分)': round(sw + dur, 1), '備考': '翌日へ継続', '製造理由': job['製造理由'], 't_start': spent + sw, 't_end': spent + sw + dur})
                                        spent += sw + dur; job['rem'] -= b_make
                                        p_rec = job['中身設計コード']; p_vol = job['容量_L']; break
                                else:
                                    sup_found = False
                                    for o_line in lines_list:
                                        if o_line == line: continue
                                        for job in queues[o_line]:
                                            if job['rem'] > 0 and job_can_support(line, job, factory_mode):
                                                sw = 10.0; avail = cap_limit - spent - sw
                                                if avail <= 5.0: break
                                                sp_min = (646 if line == '5号機' else 500) / 60 if factory_mode == "関西工場" else ((730 if job['容量_L'] in [12, 14] else 650) if line == '5号機' else 400) / 60
                                                max_b = avail * sp_min
                                                if job['rem'] <= max_b:
                                                    b_make = job['rem']; dur = b_make / sp_min
                                                    sched.append({'稼働日': f"{day_cnt}日目", '製造日': d_str_disp, '曜日': w_kanji, '製造ライン': line, '配合コード': job['中身設計コード'], '品目コード': job['品目コード'], '品目名': job['品目名'], '指示数量(袋)': int(b_make), '製造時間(分)': round(dur, 1), '切り替え(分)': round(sw, 1), '合計拘束時間(分)': round(sw + dur, 1), '備考': f"★{o_line}応援(全量)", '製造理由': job['製造理由'], 't_start': spent + sw, 't_end': spent + sw + dur})
                                                    spent += sw + dur; job['rem'] = 0; sup_found = True; p_rec = job['中身設計コード']; p_vol = job['容量_L']; break
                                                else:
                                                    b_make = math.floor(max_b)
                                                    if b_make <= 0: break
                                                    dur = b_make / sp_min
                                                    sched.append({'稼働日': f"{day_cnt}日目", '製造日': d_str_disp, '曜日': w_kanji, '製造ライン': line, '配合コード': job['中身設計コード'], '品目コード': job['品目コード'], '品目名': job['品目名'], '指示数量(袋)': int(b_make), '製造時間(分)': round(dur, 1), '切り替え(分)': round(sw, 1), '合計拘束時間(分)': round(sw + dur, 1), '備考': f"★{o_line}応援(継続)", '製造理由': job['製造理由'], 't_start': spent + sw, 't_end': spent + sw + dur})
                                                    spent += sw + dur; job['rem'] -= b_make; sup_found = True; p_rec = job['中身設計コード']; p_vol = job['容量_L']; break
                                        if sup_found: break
                                    if not sup_found: break
                        loop_d = get_next_w_date(loop_d + datetime.timedelta(days=1), holidays_input)
                        day_cnt += 1
                        if day_cnt > 45: break
                    return sched, day_cnt - 1

                ov_res = 0
                full_sched, gen_days = run_sim(0)
                if gen_days > target_days:
                    for t_ov in [30, 60, 90, 120, 150, 180, 210]:
                        ts, td = run_sim(t_ov)
                        if td <= target_days: ov_res = t_ov; full_sched = ts; gen_days = td; break

                if ov_res > 0: st.warning(f"📢 目標の{target_days}日以内に終わらせるため、毎日一律【{ov_res}分】の残業が必要です。")
                else: st.success(f"🟢 【残業不要】通常の定時稼働のまま【{gen_days}日間】ですべて作り切れます！")

                wb = Workbook(); wb.remove(wb.active)
                navy = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                w_font = Font(name="Meiryo UI", size=11, bold=True, color="FFFFFF")
                r_font = Font(name="Meiryo UI", size=10); b_font = Font(name="Meiryo UI", size=10, bold=True)
                b_all = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))

                ws1 = wb.create_sheet(title="製造品目・バッチ集計"); ws1.views.sheetView[0].showGridLines = True
                ws1.append(["品目コード", "品目名", "製造ライン", "配合レシピ", "現在の在庫", "安全在庫数", "安全割れ不足数", "今月の計画残数", "決定製造m3", "最終製造総数(袋)", "製造理由"])
                for _, r in df_final_sorted.iterrows(): ws1.append([r['品目コード'], r['品目名'], r['製造ライン'], r['中身設計コード'], r['現在の在庫'], r['安全在庫数'], r['安全割れ不足数'], r['今月の計画残数'], r['製造決定_m3'], r['計画製造袋数'], r['製造理由']])

                ws2 = wb.create_sheet(title="日別・号機別製造計画"); ws2.views.sheetView[0].showGridLines = True
                ws2.append(["稼働日", "製造日", "曜日", "製造ライン", "配合コード", "品目コード", "品目名", "指示数量(袋)", "製造時間(分)", "切り替え(分)", "合計拘束時間(分)", "備考", "製造理由"])
                for j in full_sched: ws2.append([j['稼働日'], j['製造日'], j['曜日'], j['製造ライン'], j['配合コード'], j['品目コード'], j['品目名'], j['指示数量(袋)'], j['製造時間(分)'], j['切り替え(分)'], j['合計拘束時間(分)'], j['備考'], j['製造理由']])

                ws3 = wb.create_sheet(title="日別・30分タイムテーブル"); ws3.views.sheetView[0].showGridLines = True
                slots = ["8:00〜8:30", "8:30〜9:00", "9:00〜9:30", "9:30〜10:00", "10:00〜10:10(休憩)", "10:10〜10:30", "10:30〜11:00", "11:00〜11:30", "11:30〜12:00", "12:00〜13:00(昼休)", "13:00〜13:30", "13:30〜14:00", "14:00〜14:30", "14:30〜15:00", "15:00〜15:10(休憩)", "15:10〜15:30", "15:30〜16:00", "16:00〜16:30", "16:30〜17:00", "17:00〜17:30", "17:30〜18:00", "18:00〜18:30", "18:30〜19:00", "19:00〜19:30", "19:30〜20:00"]
                ws3.append(["稼働日", "製造日", "ライン"] + slots)

                u_days = []
                seen_d = set()
                for j in full_sched:
                    k = (j['稼働日'], j['製造日'])
                    if k not in seen_d: seen_d.add(k); u_days.append(k)

                mat_map = {}
                for (ds, ddt) in u_days:
                    wk = ["月", "火", "水", "木", "金", "土", "日"][datetime.datetime.strptime(ddt, "%Y/%m/%d").weekday()]
                    for ln in lines_list:
                        ldisp = {"1号機": "NO.1", "2号機": "NO.2", "3号機": "NO.3", "5号機": "NO.5", "6号機": "NO.6"}.get(ln, ln)
                        row_arr = [ds, f"{ddt} ({wk})", ldisp] + [""] * 25
                        ws3.append(row_arr)
                        mat_map[(ds, ln)] = ws3[ws3.max_row]

                s_rng = {0:(0,30), 1:(30,60), 2:(60,90), 3:(90,120), 4:(None,"休"), 5:(120,140), 6:(140,170), 7:(170,200), 8:(200,230), 9:(None,"昼"), 10:(230,260), 11:(260,290), 12:(290,320), 13:(320,350), 14:(None,"休"), 15:(350,370), 16:(370,400), 17:(400,430), 18:(430,460), 19:(460,490), 20:(490,520), 21:(520,550), 22:(550,580), 23:(580,610), 24:(610,640)}

                for j in full_sched:
                    t_cell_row = mat_map.get((j['稼働日'], j['製造ライン']))
                    if t_cell_row:
                        sm = j['t_start']; em = j['t_end']
                        for si in range(25):
                            if si in [4, 9, 14]:
                                t_cell_row[si+3].value = "小休憩" if si!=9 else "昼休憩"; continue
                            st_s, ed_s = s_rng[si]
                            if max(sm, st_s) < min(em, ed_s) - 1e-5:
                                pfx = "★(応援)\n" if "応援" in j['備考'] else ""
                                cur_v = t_cell_row[si+3].value or ""
                                add_t = f"{pfx}{j['品目名']}\n({j['指示数量(袋)']}袋)"
                                t_cell_row[si+3].value = f"{cur_v}＋\n{add_t}" if cur_v else add_t

                for sheet in [ws1, ws2, ws3]:
                    sheet.row_dimensions[1].height = 26
                    for c in sheet[1]: c.fill = navy; c.font = w_font; c.alignment = Alignment(horizontal="center", vertical="center")
                    for r_idx in range(2, sheet.max_row+1):
                        sheet.row_dimensions[r_idx].height = 20 if sheet!=ws3 else 60
                        is_z = (r_idx%2 == 0) if sheet!=ws3 else ((r_idx-2)//len(lines_list)%2 == 0)
                        for c in sheet[r_idx]:
                            c.font = r_font; c.border = b_all
                            if sheet==ws3:
                                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                                if c.column in [2,3]: c.fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid"); c.font = b_font
                                elif c.column in [8,13,18]: c.fill = PatternFill(start_color="E4DFEC" if c.column==13 else "EAEAEA", end_color="E4DFEC" if c.column==13 else "EAEAEA", fill_type="solid"); c.font = b_font
                                elif c.column>3 and is_z: c.fill = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
                            else:
                                if is_z: c.fill = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
                                c.alignment = Alignment(horizontal="right" if isinstance(c.value, (int,float)) else "left", vertical="center")
                                if isinstance(c.value, (int,float)): c.number_format = "#,##0"
                    if sheet!=ws3:
                        for col in sheet.columns: sheet.column_dimensions[get_column_letter(col[0].column)].width = max(max(sum(2 if ord(char)>128 else 1 for char in str(cell.value or '')) for cell in col)+3, 12)
                        sheet.freeze_panes = "A2"
                    else:
                        nl = len(lines_list)
                        for di in range(len(u_days)):
                            sheet.merge_cells(start_row=2+di*nl, start_column=1, end_row=2+di*nl+nl-1, end_column=1)
                            sheet.merge_cells(start_row=2+di*nl, start_column=2, end_row=2+di*nl+nl-1, end_column=2)
                        sheet.column_dimensions['A'].width=12; sheet.column_dimensions['B'].width=16; sheet.column_dimensions['C'].width=12
                        for ci in range(4, sheet.max_column+1): sheet.column_dimensions[get_column_letter(ci)].width = 23
                        sheet.freeze_panes = "D2"

                out_io = io.BytesIO(); wb.save(out_io); out_io.seek(0)
                st.download_button("📊 指示スケジュール表(.xlsx)をダウンロード", out_io, f"【確定版】{factory_mode}_{target_month}度_スケジュール表.xlsx")
            except Exception as e: st.error(f"計算実行エラー: {e}")
