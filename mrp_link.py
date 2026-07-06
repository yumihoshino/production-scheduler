"""MRPアプリ（発注リスケ提案ツール）へのスケジュール連携部品。

【製造計画アプリへの組み込み方】
1. このファイル（mrp_link.py）を製造計画アプリのリポジトリ直下に置く
2. requirements.txt に requests>=2.31 を追加（無ければ）
3. Streamlit CloudのSecretsに以下を登録（MRPアプリと同じトークンでOK）:
       [github]
       token = "github_pat_..."          # mrp-reschedule-app へのContents R/W権限
       repo = "yumihoshino/mrp-reschedule-app"
4. app.py（製造計画アプリ）の、確定版Excelをダウンロードさせている箇所の近くに:

       from mrp_link import mrp_link_button
       mrp_link_button(excel_bytes, site="本社")   # excel_bytes = 確定版ExcelのBytes
       # 関西版の画面なら site="関西工場"

   ※ excel_bytes は st.download_button に渡しているのと同じデータ（bytes）です。
"""
import base64
import requests
import streamlit as st

API = "https://api.github.com"


def _save(token: str, repo: str, path: str, content: bytes, message: str):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    url = f"{API}/repos/{repo}/contents/{path}"
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


def mrp_link_button(excel_bytes: bytes, site: str = "本社", key: str | None = None):
    """確定版スケジュールをMRPアプリへ連携するボタンを表示する。

    site: "本社" または "関西工場"
    """
    label = f"📤 MRPアプリへ連携（{site}）"
    if st.button(label, key=key or f"mrp_link_{site}"):
        gh = st.secrets.get("github", {})
        if not gh.get("token") or not gh.get("repo"):
            st.error("連携が未設定です。管理者にご連絡ください（Secretsにgithub設定が必要）。")
            return
        path = f"data/製造スケジュール_{site}.xlsx"
        ok, msg = _save(gh["token"], gh["repo"], path, excel_bytes,
                        f"製造スケジュール連携（{site}・製造計画アプリから）")
        if ok:
            st.success(f"MRPアプリへ連携しました（{site}）。"
                       "MRPアプリ側は1〜2分後の再読み込みで最新版が使えます。")
        else:
            st.error(f"連携に失敗しました: {msg}")
