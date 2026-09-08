# 求人広告 営業リスト作成アプリ

求人ページのURL（または求人原稿テキスト）から、営業アプローチ用の情報を
AIで抽出・生成し、SQLiteに蓄積・Excel出力する Streamlit アプリ。

## 機能

- 求人ページURLを `https://r.jina.ai/` 経由で Markdown 取得（失敗時は原稿テキスト直接入力にフォールバック）
- Anthropic Messages API で以下を抽出/生成
  - 抽出元メディア名 / 企業名 / 職種 / 住所 / 電話番号 / ホームページURL
  - 採用難易度スコア(1〜100) / 分析結果 / 架電用トーク
- 電話番号・HP URL が欠損時は DuckDuckGo 検索の上位スニペットを AI に渡して補完
- 結果を SQLite (`recruitment_list.db`) に保存し、一覧表示
- 一覧を Excel (`.xlsx`) でダウンロード
- APP_PASSWORD による簡易パスワード保護

## セットアップ（ローカル）

```bash
pip install -r requirements.txt
```

`.streamlit/secrets.toml` に以下を設定（`.gitignore` で除外済み。コミットしないこと）:

```toml
ANTHROPIC_API_KEY = "sk-ant-..."
APP_PASSWORD = "任意のパスワード"
```

## 起動

```bash
streamlit run app.py
```

ブラウザで `http://localhost:8501` を開き、パスワードを入力するとアプリ本体が表示される。

## クラウド展開（Streamlit Community Cloud）

1. このリポジトリを GitHub に push
2. Streamlit Community Cloud で `app.py` を指定してデプロイ
3. アプリの **Settings → Secrets** に `ANTHROPIC_API_KEY` と `APP_PASSWORD` を登録

## 使用モデル

`app.py` の `MODEL` 定数で指定（既定: `claude-sonnet-5`）。
`claude-opus-5` / `claude-haiku-4-5` などに1行で切替可能。
