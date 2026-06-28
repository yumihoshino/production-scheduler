# マスタ読込・自動無条件復元ロジック（初期状態）
df_bom = None
if os.path.exists("bom_master_local.csv"):
    try: df_bom = pd.read_csv("bom_master_local.csv", encoding='utf-8')
    except UnicodeDecodeError: df_bom = pd.read_csv("bom_master_local.csv", encoding='cp932')
elif os.path.exists("bom_master.xlsx"): 
    df_bom = pd.read_excel("bom_master.xlsx", header=None)
    df_bom = clean_bom_master(df_bom)
elif 'bom_data' in st.session_state: 
    df_bom = st.session_state['bom_data']

# もし③番の個別マスタアップローダーにファイルが置かれたら最優先適用
if file_bom is not None:
    if file_bom.name.endswith('.csv'):
        try: df_bom = pd.read_csv(file_bom, encoding='utf-8', header=None)
        except UnicodeDecodeError:
            file_bom.seek(0); df_bom = pd.read_csv(file_bom, encoding='cp932', header=None)
    else: df_bom = load_excel_sheets_merged(file_bom, ["マスタ", "BOM", "BomMaster", "ﾏｽﾀ"])
    df_bom = clean_bom_master(df_bom)
    if df_bom is not None:
        df_bom.to_csv("bom_master_local.csv", index=False, encoding='utf-8')
        st.session_state['bom_data'] = df_bom

# 🌟【超大進化：もし③番が空っぽでも、②番の計画書ファイルの中に「ﾏｽﾀ」や「BOM」シートがあれば全自動で吸い上げる！】
if df_bom is None and file_gekkan is not None:
    try:
        xl_gekkan_test = pd.ExcelFile(file_gekkan)
        m_sheets = [s for s in xl_gekkan_test.sheet_names if any(k in s for k in ["マスタ", "BOM", "BomMaster", "ﾏｽﾀ"])]
        if m_sheets:
            df_bom_auto = pd.read_excel(xl_gekkan_test, sheet_name=m_sheets[0], header=None)
            df_bom = clean_bom_master(df_bom_auto)
            if df_bom is not None:
                df_bom.to_csv("bom_master_local.csv", index=False, encoding='utf-8')
                st.session_state['bom_data'] = df_bom
    except:
        pass
