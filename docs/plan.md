# 構築方針と進め方

目的: `.env` の接続情報でローカル PostgreSQL のスキーマ情報を読み取り、Excel（.xlsx）にマルチシートで出力するローカル実行ツールを構築する。

方針:
- パッケージ/仮想環境管理は `uv` を使用（`.venv` を活用）。
- 依存は `psycopg2-binary`, `pandas`, `xlsxwriter`, `python-dotenv` を採用。
- `.env` は SQLAlchemy 形式の URL（例: `postgresql+psycopg2://...`）にも対応し、`psycopg2` 接続可能な URI へ正規化する。
- SPEC.md のクエリを踏襲し、`Tables/Columns/Constraints/Indexes/Triggers/Functions/Views` を各シートで出力。
- CLI を用意し、`--env-key` or `--dsn` で接続先を指定。`--with-counts` は任意（重い）。

実装ステップ:
1) uv で依存追加・エントリポイント整備
2) 接続ヘルパー実装（.env ロード、URL 正規化）
3) エクスポータ実装（SPEC の SQL + Excel 出力）
4) CLI 実装・動作確認（--help, 接続チェック）
5) 実運用ガイド整備（docs/usage.md）

検証の進め方:
- ステップ毎に `uv run pgsql-export --help` など軽い動作確認を行う。
- 接続は `.env` の `--env-key` を指定して試験。DB 未起動/権限不足などの際はエラーメッセージで原因を明示する。

拡張余地（任意）:
- Info シート（生成日時/ git rev / alembic heads）
- 対象スキーマのフィルタ、ERD 補助情報の抽出
