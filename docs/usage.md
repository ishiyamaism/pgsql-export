# 使い方

Python 3.13 以上が必要です。

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

# 行数も計測
uv run pgsql-export --env-key INVENTORY_DB --with-counts --out meta.xlsx
```

`--with-counts` は、対象テーブルごとに正確な `COUNT(*)` を実行します。大規模なDBでは時間とI/Oがかかるため、既定では無効です。また、対象テーブルへの `SELECT` 権限が必要です。

## 出力内容

Excel に次の7シートを作成し、自動列幅・ヘッダ固定・フィルタを設定します。

| シート | 内容 |
| --- | --- |
| Tables | スキーマ、テーブル、所有者、任意の正確な行数、コメント |
| Columns | 列順、型、NULL可否、デフォルト、IDENTITY、コメント |
| Constraints | 主キー、UNIQUE、外部キー、CHECK制約の定義 |
| Indexes | インデックス方式、一意性、条件、定義 |
| Triggers | 実行タイミングとイベント、関数、定義 |
| Functions | 引数、戻り値、VOLATILE属性、セキュリティ、コメント |
| Views | ビュー名、コメント、定義 |

`information_schema` と名前が `pg_` から始まる内部スキーマは対象外です。DB内の文字列はそのまま保持し、`=1+1` のような文字列をExcel数式に変換したり、URLらしい文字列をハイパーリンクに変換したりしません。

## トラブルシュート

- 認証失敗/接続拒否: ユーザー/パスワード、ホスト、ポート、DB 名を確認。読み取り専用ロール推奨。
- `--with-counts` で権限エラー: 対象スキーマに `USAGE`、対象テーブルに `SELECT` 権限があるか確認。
- URL 形式エラー: `postgresql+psycopg2://...` は自動で `postgresql://...` に正規化します。
