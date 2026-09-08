"""
求人広告営業用リスト作成アプリ
--------------------------------
- 求人ページURL（r.jina.ai 経由でMarkdown取得） or 求人原稿テキストの直接入力
- Anthropic API で企業情報・分析・架電トークを抽出/生成
- 電話番号 / HP URL が欠損していれば DuckDuckGo 検索で補完
- SQLite (recruitment_list.db) に保存し、一覧表示 & Excel ダウンロード
"""

import io
import re
import json
import sqlite3
from datetime import datetime

import requests
import pandas as pd
import streamlit as st
import anthropic

try:
    # 従来のパッケージ名
    from duckduckgo_search import DDGS
except ImportError:  # パッケージ改名後（ddgs）へのフォールバック
    from ddgs import DDGS


# --------------------------------------------------------------------------
# 設定
# --------------------------------------------------------------------------
DB_NAME = "recruitment_list.db"
JINA_PREFIX = "https://r.jina.ai/"

# 使用モデル（Messages API）。変更する場合はここだけ差し替える。
# このAPIキーでは Claude 3.x は提供終了。Sonnet系の後継として claude-sonnet-5 を使用。
# 他候補: claude-opus-5（高精度・高コスト） / claude-haiku-4-5（高速・低コスト）
MODEL = "claude-sonnet-5"

# 抽出結果のキー(英語) -> 表示名(日本語) の対応
FIELD_MAP = {
    "media_name": "抽出元メディア名",
    "company_name": "企業名",
    "job_title": "職種",
    "address": "住所",
    "phone_number": "電話番号",
    "homepage_url": "ホームページURL",
    "difficulty_score": "採用難易度スコア",
    "analysis": "分析結果",
    "sales_talk": "架電用トーク",
}

# 仕様で「必ず設定する」と指定された system_prompt（本文はそのまま）
BASE_SYSTEM_PROMPT = """あなたは求人広告業界のトップ営業であり、組織構造の課題解決に長けた採用コンサルタントです。
入力された求人原稿のテキストデータを分析し、以下のフォーマット（JSON）で情報を抽出・生成してください。
【分析フレームワーク】以下の「構造的なエラー」をチェックしてください。
1. 市場環境の不一致（ターゲット層が相対的に狭すぎないか）
2. 条件の劣位性（給与等が採用競合と比較して劣っていないか）
3. 情報の不明瞭さ（仕事内容やターゲットが曖昧でないか）
4. 訴求力の欠如（写真の意図不明瞭、情報不足など）
【トーク生成条件】挨拶の直後に構造的な課題を1つ端的に指摘し、顧客ニーズの仮説（応募数不足、有効応募不足など）に基づき相手の課題をヒアリングし、解決策のアイデアがあることを伝えアポを打診する30秒の自然な口語体の日本語。"""

# JSON の出力仕様を追記（キー名を固定して機械的にパースできるようにする）
OUTPUT_SPEC = """
【出力フォーマット】
必ず次のキーだけを持つJSONオブジェクトを1つだけ返してください。
前後の説明文やMarkdownのコードフェンス(```)は一切付けないでください。
{
  "media_name": "抽出元メディア名（求人媒体名。不明ならnull）",
  "company_name": "企業名",
  "job_title": "職種",
  "address": "住所",
  "phone_number": "電話番号（記載がなければnull）",
  "homepage_url": "ホームページURL（記載がなければnull）",
  "difficulty_score": 採用難易度スコア（1〜100の整数。高いほど採用が困難）,
  "analysis": "分析結果（上記フレームワークに基づく構造的課題の指摘。文章）",
  "sales_talk": "架電用トーク（上記トーク生成条件を満たす口語体の日本語）"
}
不明な項目の値は null にしてください。
"""

SYSTEM_PROMPT = BASE_SYSTEM_PROMPT + "\n" + OUTPUT_SPEC


# --------------------------------------------------------------------------
# データベース
# --------------------------------------------------------------------------
def init_db() -> None:
    conn = sqlite3.connect(DB_NAME)
    try:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS recruitment_list (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                media_name       TEXT,
                company_name     TEXT,
                job_title        TEXT,
                address          TEXT,
                phone_number     TEXT,
                homepage_url     TEXT,
                difficulty_score INTEGER,
                analysis         TEXT,
                sales_talk       TEXT,
                source_url       TEXT,
                created_at       TEXT
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def save_to_db(rec: dict, source_url: str) -> None:
    conn = sqlite3.connect(DB_NAME)
    try:
        conn.execute(
            """
            INSERT INTO recruitment_list
                (media_name, company_name, job_title, address, phone_number,
                 homepage_url, difficulty_score, analysis, sales_talk,
                 source_url, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                rec.get("media_name"),
                rec.get("company_name"),
                rec.get("job_title"),
                rec.get("address"),
                rec.get("phone_number"),
                rec.get("homepage_url"),
                rec.get("difficulty_score"),
                rec.get("analysis"),
                rec.get("sales_talk"),
                source_url or None,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ),
        )
        conn.commit()
    finally:
        conn.close()


def load_df() -> pd.DataFrame:
    conn = sqlite3.connect(DB_NAME)
    try:
        df = pd.read_sql_query(
            "SELECT * FROM recruitment_list ORDER BY id DESC", conn
        )
    finally:
        conn.close()

    rename = dict(FIELD_MAP)
    rename.update({"id": "ID", "source_url": "取得元URL", "created_at": "登録日時"})
    return df.rename(columns=rename)


# --------------------------------------------------------------------------
# テキスト取得（r.jina.ai）
# --------------------------------------------------------------------------
def fetch_text_from_url(url: str):
    """r.jina.ai の Reader API 経由で Markdown テキストを取得する。"""
    endpoint = JINA_PREFIX + url
    try:
        resp = requests.get(
            endpoint,
            timeout=60,
            headers={
                "User-Agent": "Mozilla/5.0 (compatible; RecruitmentListBot/1.0)",
                "X-Return-Format": "markdown",
            },
        )
        resp.raise_for_status()
        return resp.text, None
    except requests.RequestException as e:
        return "", str(e)


# --------------------------------------------------------------------------
# Anthropic API
# --------------------------------------------------------------------------
def call_claude(client: anthropic.Anthropic, system: str, user: str,
                max_tokens: int = 2048) -> str:
    resp = client.messages.create(
        model=MODEL,
        max_tokens=max_tokens,
        system=system,
        messages=[{"role": "user", "content": user}],
    )
    return "".join(b.text for b in resp.content if b.type == "text").strip()


def parse_json_object(text: str) -> dict:
    """モデル出力から JSON オブジェクト部分を取り出してパースする。"""
    text = text.strip()
    # コードフェンスを除去
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\n?", "", text)
        text = re.sub(r"\n?```$", "", text).strip()
    # 最初の { から最後の } までを抽出
    start, end = text.find("{"), text.rfind("}")
    if start != -1 and end != -1 and end > start:
        text = text[start : end + 1]
    return json.loads(text)


def extract_with_ai(client: anthropic.Anthropic, source_text: str) -> dict:
    # トークン超過対策として長すぎる原稿は先頭を優先して切り詰める
    max_chars = 24000
    if len(source_text) > max_chars:
        source_text = source_text[:max_chars]

    user = (
        "以下の求人原稿を分析し、指定のJSONフォーマットのみで返答してください。\n\n"
        "=== 求人原稿ここから ===\n"
        f"{source_text}\n"
        "=== 求人原稿ここまで ==="
    )
    return parse_json_object(call_claude(client, SYSTEM_PROMPT, user, max_tokens=2048))


# --------------------------------------------------------------------------
# 欠損情報の補完（DuckDuckGo 検索 + AI）
# --------------------------------------------------------------------------
def is_missing(value) -> bool:
    if value is None:
        return True
    s = str(value).strip().lower()
    return s in ("", "null", "none", "不明", "なし", "n/a")


def ddg_search(query: str, max_results: int = 5) -> list[str]:
    try:
        with DDGS() as ddgs:
            results = ddgs.text(query, max_results=max_results) or []
    except Exception:
        return []

    snippets = []
    for r in results:
        part = " ".join(
            x for x in (r.get("title"), r.get("body"), r.get("href")) if x
        )
        if part:
            snippets.append(part)
    return snippets


def supplement_missing_info(client: anthropic.Anthropic, rec: dict):
    """電話番号 / HP URL が欠損していれば Web 検索結果を AI に渡して補完する。"""
    if not (is_missing(rec.get("phone_number")) or is_missing(rec.get("homepage_url"))):
        return rec, None

    company = (rec.get("company_name") or "").strip()
    address = (rec.get("address") or "").strip()
    query = " ".join(x for x in (company, address, "電話番号") if x).strip()
    if not company and not address:
        return rec, "企業名・住所が不明なため検索できませんでした"

    snippets = ddg_search(query, max_results=5)
    if not snippets:
        return rec, "DuckDuckGo の検索結果が取得できませんでした"

    joined = "\n".join(f"- {s}" for s in snippets)
    user = (
        f"企業名: {company or '不明'}\n"
        f"住所: {address or '不明'}\n\n"
        "以下はWeb検索結果の上位スニペットです。\n"
        f"{joined}\n\n"
        "この中から該当企業の『電話番号』と『ホームページURL』を推定してください。\n"
        '出力は {"phone_number": "...", "homepage_url": "..."} のJSONのみ。'
        "判断できない項目は null にしてください。"
    )

    try:
        found = parse_json_object(
            call_claude(
                client,
                "あなたは企業情報を正確に抽出するアシスタントです。JSONのみを返します。",
                user,
                max_tokens=512,
            )
        )
    except Exception as e:  # noqa: BLE001
        return rec, f"補完AIの呼び出しに失敗しました: {e}"

    if is_missing(rec.get("phone_number")) and not is_missing(found.get("phone_number")):
        rec["phone_number"] = found.get("phone_number")
    if is_missing(rec.get("homepage_url")) and not is_missing(found.get("homepage_url")):
        rec["homepage_url"] = found.get("homepage_url")
    return rec, None


# --------------------------------------------------------------------------
# 正規化
# --------------------------------------------------------------------------
def normalize_record(raw: dict) -> dict:
    rec = {key: raw.get(key) for key in FIELD_MAP}

    # スコアを整数へ
    score = rec.get("difficulty_score")
    try:
        rec["difficulty_score"] = int(float(str(score)))
    except (TypeError, ValueError):
        rec["difficulty_score"] = None

    # 空文字は None に寄せる
    for key, val in rec.items():
        if isinstance(val, str) and not val.strip():
            rec[key] = None
    return rec


# --------------------------------------------------------------------------
# Excel 出力
# --------------------------------------------------------------------------
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="求人リスト")
    return buf.getvalue()


# --------------------------------------------------------------------------
# 認証 / シークレット
# --------------------------------------------------------------------------
def get_secret(key: str):
    """secrets.toml / Streamlit Cloud のシークレットから安全に取得する。"""
    try:
        return st.secrets[key]
    except Exception:  # noqa: BLE001  未設定・ファイルなしなど
        return None


def check_password() -> bool:
    """APP_PASSWORD と一致するまでアプリ本体を表示しない簡易ゲート。"""
    if st.session_state.get("password_ok"):
        return True

    st.title("🔒 ログイン")
    with st.form("login_form"):
        pw = st.text_input("パスワード", type="password")
        submitted = st.form_submit_button("ログイン")

    if submitted:
        expected = get_secret("APP_PASSWORD")
        if not expected:
            st.error("APP_PASSWORD が secrets に設定されていません。管理者に連絡してください。")
        elif pw == expected:
            st.session_state["password_ok"] = True
            st.rerun()
        else:
            st.error("パスワードが違います。")
    return False


# --------------------------------------------------------------------------
# UI
# --------------------------------------------------------------------------
def main() -> None:
    st.set_page_config(page_title="求人広告 営業リスト作成", layout="wide")

    # --- パスワード保護: 一致するまで以降を表示しない ---
    if not check_password():
        st.stop()

    api_key = get_secret("ANTHROPIC_API_KEY")
    if not api_key:
        st.error("ANTHROPIC_API_KEY が secrets に設定されていません。管理者に連絡してください。")
        st.stop()

    st.title("📋 求人広告 営業リスト作成アプリ")
    init_db()

    with st.sidebar:
        st.header("設定")
        st.caption(f"使用モデル: `{MODEL}`")
        if st.button("ログアウト"):
            st.session_state.pop("password_ok", None)
            st.rerun()

    st.subheader("1. 入力")
    url = st.text_input("求人ページのURL", placeholder="https://...")
    manuscript = st.text_area(
        "求人原稿テキストの直接入力（URL取得に失敗した場合のフォールバック）",
        height=200,
        placeholder="求人原稿をそのまま貼り付け",
    )
    run = st.button("実行", type="primary")

    if run:
        if not url.strip() and not manuscript.strip():
            st.error("URL または 求人原稿テキスト のどちらかを入力してください。")
            st.stop()

        client = anthropic.Anthropic(api_key=api_key)
        source_url = url.strip()
        source_text = ""

        # --- テキスト取得 ---
        if source_url:
            with st.spinner("r.jina.ai 経由でページを取得中..."):
                source_text, err = fetch_text_from_url(source_url)
            if err:
                st.warning(f"URL取得に失敗しました: {err}")
                if manuscript.strip():
                    st.info("フォールバック: 直接入力されたテキストを使用します。")
                    source_text = manuscript.strip()
                else:
                    st.error("フォールバック用のテキストが未入力のため中断します。")
                    st.stop()
        else:
            source_text = manuscript.strip()

        if not source_text.strip():
            st.error("解析対象のテキストが空です。")
            st.stop()

        with st.expander("取得テキスト（先頭2000文字）"):
            st.text(source_text[:2000])

        # --- AI 抽出 ---
        with st.spinner("AIが求人原稿を分析中..."):
            try:
                rec = normalize_record(extract_with_ai(client, source_text))
            except Exception as e:  # noqa: BLE001
                st.error(f"AI抽出に失敗しました: {e}")
                st.stop()

        # --- 欠損補完 ---
        if is_missing(rec.get("phone_number")) or is_missing(rec.get("homepage_url")):
            with st.spinner("DuckDuckGo で不足情報（電話番号 / HP URL）を検索・補完中..."):
                rec, sup_err = supplement_missing_info(client, rec)
            if sup_err:
                st.warning(f"補完処理: {sup_err}")

        # --- 保存 ---
        save_to_db(rec, source_url)
        st.success("データベースに保存しました。")
        st.subheader("2. 抽出結果")
        st.json({FIELD_MAP[k]: v for k, v in rec.items()})

    # ----------------------------------------------------------------------
    st.divider()
    st.subheader("3. 保存済みリスト")
    df = load_df()
    if df.empty:
        st.info("まだデータがありません。")
    else:
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button(
            "Excel（.xlsx）でダウンロード",
            data=to_excel_bytes(df),
            file_name=f"recruitment_list_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
