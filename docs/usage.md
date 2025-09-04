# 使い方

## セットアップ

- 前提: `uv` が導入済み。`.venv` を使用します。
- 依存追加はプロジェクトに固定済みです（`pyproject.toml` / `uv.lock`）。

```
# （初回）依存を解決
uv sync
```

## 接続情報

- `.env` に接続 URL を用意（SQLAlchemy 形式でも可）
  - 例: `INVENTORY_DB=postgresql+psycopg2://user:pass@localhost:5432/inventory_db_01`
- もしくは `--dsn` で `postgresql://` 形式や DSN 文字列を指定

## 実行

```
# ヘルプ
uv run pgsql-export --help

# .env のキーを指定して出力
uv run pgsql-export --env-key INVENTORY_DB --out meta.xlsx

# 直接 DSN/URI を渡す
uv run pgsql-export --dsn "postgresql://user:pass@localhost/db" --out meta.xlsx

# 行数も計測（重い可能性）
uv run pgsql-export --env-key INVENTORY_DB --with-counts --out meta.xlsx
```

出力: Excel に `Tables / Columns / Constraints / Indexes / Triggers / Functions / Views` シートを作成し、自動列幅・ヘッダ固定・フィルタを設定します。

## トラブルシュート

- 認証失敗/接続拒否: ユーザー/パスワード、ホスト、ポート、DB 名を確認。読み取り専用ロール推奨。
- 権限不足: 対象スキーマに `USAGE`、テーブルに `SELECT` 権限が必要。
- URL 形式エラー: `postgresql+psycopg2://...` は自動で `postgresql://...` に正規化します。
