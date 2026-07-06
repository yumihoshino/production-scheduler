"""MRPアプリ（発注リスケ提案ツール）へのスケジュール自動連携部品（v2）。

v1（ボタン方式）はStreamlitの再実行仕様により動作しないため、
確定版Excelの生成と同時に自動でMRPへ連携する方式に変更。

【組み込み方（v1からの変更）】
このファイルで production-scheduler の mrp_link.py を丸ごと上書きするだけ。
app.py の呼び出し（mrp_link_button）はそのままでも動きます。
"""
import base64
import hashlib
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


def mrp_auto_link(excel_bytes: bytes, site: str = "本社"):
    """確定版スケジュールを、生成と同時にMRPアプリへ自動連携する。

    同じ内容を同一セッションで二重送信しないよう、内容のハッシュで判定する。
    """
    gh = st.secrets.get("github", {})
    if not gh.get("token") or not gh.get("repo"):
        st.caption("（MRP連携は未設定です。管理者にご連絡ください）")
        return

    digest = hashlib.md5(excel_bytes).hexdigest()
    state_key = f"mrp_linked_{site}"
    if st.session_state.get(state_key) == digest:
        st.caption(f"✅ この確定版はMRPアプリへ連携済みです（{site}）")
        return

    path = f"data/製造スケジュール_{site}.xlsx"
    ok, msg = _save(gh["token"], gh["repo"], path, excel_bytes,
                    f"製造スケジュール連携（{site}・製造計画アプリから自動送信）")
    if ok:
        st.session_state[state_key] = digest
        st.success(f"📤 MRPアプリへ自動連携しました（{site}）。"
                   "MRPアプリ側は1〜2分後に最新版が使えます。")
    else:
        st.warning(f"MRP連携に失敗しました（スケジュール自体は正常です）: {msg}")


# 旧関数名でも動くように残す（v1からの移行用・app.pyの変更不要）
def mrp_link_button(excel_bytes: bytes, site: str = "本社", key: str | None = None):
    mrp_auto_link(excel_bytes, site)
