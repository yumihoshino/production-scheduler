"""
GEN ERP 連携部品：品目マスタ・取引先マスタをAPIで取得し、GitHubへ自動保存する。

mrp_link.py と同じ方式（GitHub Contents API + st.secrets["github"]）を使う。
アプリ側の app.py には、下の「組み込み方」の1行を呼ぶだけでOK。

【前提】Streamlit Cloud の Secrets に以下がすでに設定済みであること（mrp_link.pyと共通）
    [github]
    token = "ghp_xxxxxxxxxxxx"
    repo = "ユーザー名/リポジトリ名"

【追加で必要なSecrets】GENのAPIトークン
    GEN_API_TOKEN = "xxxxxxxxxxxxxxxx"

【組み込み方】
    import gen_master_sync

    if st.button("GENから品目・取引先マスタを取得してGitHubへ保存"):
        gen_master_sync.sync_gen_masters()
"""
import base64
import hashlib
import io
import time

import pandas as pd
import requests
import streamlit as st

GITHUB_API = "https://api.github.com"
GEN_BASE_URL = "https://setogahara.gen-cloud.jp"

# GitHub上の保存先パス（必要であれば変更してください）
ITEM_MASTER_PATH = "data/品目マスタ.csv"
CUSTOMER_MASTER_PATH = "data/取引先マスタ.csv"


# --- GitHub保存（mrp_link.pyの_saveと同じ方式） ---
def _save_to_github(token: str, repo: str, path: str, content: bytes, message: str):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    url = f"{GITHUB_API}/repos/{repo}/contents/{path}"
    sha = None
    r = requests.get(url, headers=headers, timeout=30)
    if r.status_code == 200:
        sha = r.json().get("sha")
    payload = {"message": message,
               "content": base64.b64encode(content).decode("ascii")}
    if sha:
        payload["sha"] = sha
    r = requests.put(url, headers=headers, json=payload, timeout=60)
    return r.status_code in (200, 201), r.json().get("message", "")


# --- GEN APIからCSV取得 ---
def _fetch_gen_csv(endpoint: str) -> bytes:
    """GENのCSVエクスポートAPIを叩いて生バイト列（CSV）を返す。"""
    token = st.secrets["GEN_API_TOKEN"]
    url = f"{GEN_BASE_URL}{endpoint}"
    headers = {"Authorization": f"Bearer {token}"}

    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()
    return response.content


def _normalize_leading_zeros(csv_bytes: bytes) -> bytes:
    """先頭ゼロが消えないよう、全列を文字列として読み込み→CSVとして書き戻す。"""
    try:
        text = csv_bytes.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = csv_bytes.decode("cp932")

    df = pd.read_csv(io.StringIO(text), dtype=str)
    out = io.StringIO()
    df.to_csv(out, index=False)
    return out.getvalue().encode("utf-8-sig")


# --- メイン処理 ---
def sync_gen_masters():
    """品目マスタ・取引先マスタをGENから取得し、GitHubへ保存する。"""
    gh = st.secrets.get("github", {})
    if not gh.get("token") or not gh.get("repo"):
        st.error("GitHub連携が未設定です（st.secrets['github']を確認してください）")
        return

    with st.spinner("GENから品目マスタを取得中..."):
        try:
            item_csv = _normalize_leading_zeros(_fetch_gen_csv("/api/itemMaster/csv"))
        except Exception as e:
            st.error(f"品目マスタの取得に失敗しました: {e}")
            return

    time.sleep(1.1)  # レート制限(1req/秒)対策

    with st.spinner("GENから取引先マスタを取得中..."):
        try:
            customer_csv = _normalize_leading_zeros(_fetch_gen_csv("/api/customerMaster/csv"))
        except Exception as e:
            st.error(f"取引先マスタの取得に失敗しました: {e}")
            return

    with st.spinner("GitHubへ保存中..."):
        ok1, msg1 = _save_to_github(
            gh["token"], gh["repo"], ITEM_MASTER_PATH, item_csv,
            "品目マスタ更新（GEN APIから自動取得）",
        )
        ok2, msg2 = _save_to_github(
            gh["token"], gh["repo"], CUSTOMER_MASTER_PATH, customer_csv,
            "取引先マスタ更新（GEN APIから自動取得）",
        )

    if ok1 and ok2:
        st.success("✅ 品目マスタ・取引先マスタをGitHubへ保存しました。"
                   "1〜2分後にアプリ側で最新版が使えます。")
    else:
        if not ok1:
            st.warning(f"品目マスタの保存に失敗しました: {msg1}")
        if not ok2:
            st.warning(f"取引先マスタの保存に失敗しました: {msg2}")
