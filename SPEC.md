OK、**完全ローカル**・**読み取り専用**で、PostgreSQLのスキーマ情報を**Excel（.xlsx）に多シート出力**する手順をまとめました。Alembic 管轄のDBそのものから直接読み出す方式です。ネットワーク不要、VPSと同じ構成をそのまま扱えます。

---

# ゴール

* `pg_catalog` から **Tables / Columns / Constraints / Indexes / Triggers / Functions / Views** を取得
* 1つのExcelに**シート別**で書き出し（フィルタ行・自動列幅・長文折返し・ヘッダ固定）
* 追加オプション：テーブル件数の計測（重いので任意）

---

# 1) 事前準備（読み取り専用ユーザー推奨）

```sql
-- 例：読み取り専用ロール
CREATE ROLE meta_reader LOGIN PASSWORD '******';
GRANT CONNECT ON DATABASE yourdb TO meta_reader;
GRANT USAGE ON SCHEMA public TO meta_reader;         -- 必要なスキーマ分
GRANT SELECT ON ALL TABLES IN SCHEMA public TO meta_reader;
ALTER DEFAULT PRIVILEGES IN SCHEMA public GRANT SELECT ON TABLES TO meta_reader;
```

---

# 2) エクスポート用スクリプト（1ファイルで完結）

**`export_pg_metadata_to_excel.py`**

```python
import os
import argparse
import math
import psycopg2
import psycopg2.extras
import pandas as pd

TABLES_SQL = """
SELECT
  n.nspname        AS schema,
  c.relname        AS table,
  pg_get_userbyid(c.relowner) AS owner,
  obj_description(c.oid, 'pg_class') AS comment
FROM pg_class c
JOIN pg_namespace n ON n.oid = c.relnamespace
WHERE c.relkind IN ('r','p')
  AND n.nspname NOT IN ('pg_catalog','information_schema')
ORDER BY 1,2;
"""

COLUMNS_SQL = """
SELECT
  n.nspname AS schema,
  c.relname AS table,
  a.attname AS column,
  pg_catalog.format_type(a.atttypid, a.atttypmod) AS data_type,
  NOT a.attnotnull AS nullable,
  pg_get_expr(ad.adbin, ad.adrelid) AS default,
  CASE WHEN a.attidentity<>'' THEN a.attidentity ELSE NULL END AS identity,
  col_description(c.oid, a.attnum) AS comment,
  a.attnum AS ordinal_position
FROM pg_attribute a
JOIN pg_class c ON c.oid = a.attrelid
JOIN pg_namespace n ON n.oid = c.relnamespace
LEFT JOIN pg_attrdef ad ON ad.adrelid = a.attrelid AND ad.adnum = a.attnum
WHERE a.attnum > 0 AND NOT a.attisdropped
  AND c.relkind IN ('r','p')
  AND n.nspname NOT IN ('pg_catalog','information_schema')
ORDER BY 1,2,a.attnum;
"""

CONSTRAINTS_SQL = """
WITH cons AS (
  SELECT
    n.nspname AS schema, c.relname AS table, con.conname AS constraint_name,
    con.contype, con.oid AS con_oid
  FROM pg_constraint con
  JOIN pg_class c ON c.oid = con.conrelid
  JOIN pg_namespace n ON n.oid = c.relnamespace
  WHERE n.nspname NOT IN ('pg_catalog','information_schema')
)
SELECT
  schema, table, constraint_name,
  CASE contype
    WHEN 'p' THEN 'PRIMARY KEY'
    WHEN 'u' THEN 'UNIQUE'
    WHEN 'f' THEN 'FOREIGN KEY'
    WHEN 'c' THEN 'CHECK'
    ELSE contype::text
  END AS type,
  pg_get_constraintdef(con_oid) AS definition
FROM cons
ORDER BY 1,2,3;
"""

INDEXES_SQL = """
SELECT
  ns.nspname AS schema,
  tbl.relname AS table,
  idx.relname AS index_name,
  am.amname  AS method,
  ix.indisunique AS is_unique,
  pg_get_indexdef(ix.indexrelid) AS definition,
  pg_get_expr(ix.indpred, ix.indrelid) AS predicate
FROM pg_index ix
JOIN pg_class idx ON idx.oid = ix.indexrelid
JOIN pg_class tbl ON tbl.oid = ix.indrelid
JOIN pg_namespace ns ON ns.oid = tbl.relnamespace
JOIN pg_am am ON am.oid = idx.relam
WHERE ns.nspname NOT IN ('pg_catalog','information_schema')
ORDER BY 1,2,3;
"""

TRIGGERS_SQL = """
SELECT
  ns.nspname AS schema,
  tbl.relname AS table,
  tg.tgname  AS trigger_name,
  CONCAT(
    CASE WHEN (tg.tgtype & 1)<>0 THEN 'BEFORE ' ELSE 'AFTER ' END,
    CASE WHEN (tg.tgtype &  4)<>0 THEN 'INSERT ' ELSE '' END,
    CASE WHEN (tg.tgtype &  8)<>0 THEN 'DELETE ' ELSE '' END,
    CASE WHEN (tg.tgtype & 16)<>0 THEN 'UPDATE ' ELSE '' END,
    CASE WHEN (tg.tgtype & 32)<>0 THEN 'TRUNCATE ' ELSE '' END
  ) AS timing_events,
  pg_get_triggerdef(tg.oid) AS definition,
  p.proname AS function
FROM pg_trigger tg
JOIN pg_class tbl ON tbl.oid = tg.tgrelid
JOIN pg_namespace ns ON ns.oid = tbl.relnamespace
JOIN pg_proc p ON p.oid = tg.tgfoid
WHERE NOT tg.tgisinternal
  AND ns.nspname NOT IN ('pg_catalog','information_schema')
ORDER BY 1,2,3;
"""

FUNCTIONS_SQL = """
SELECT
  ns.nspname AS schema,
  p.proname  AS name,
  pg_get_function_arguments(p.oid) AS arguments,
  pg_catalog.format_type(p.prorettype, NULL) AS return_type,
  CASE p.provolatile WHEN 'i' THEN 'IMMUTABLE' WHEN 's' THEN 'STABLE' ELSE 'VOLATILE' END AS volatility,
  CASE WHEN p.prosecdef THEN 'SECURITY DEFINER' ELSE 'SECURITY INVOKER' END AS security,
  obj_description(p.oid, 'pg_proc') AS comment
FROM pg_proc p
JOIN pg_namespace ns ON ns.oid = p.pronamespace
WHERE ns.nspname NOT IN ('pg_catalog','information_schema')
ORDER BY 1,2;
"""

VIEWS_SQL = """
SELECT
  ns.nspname AS schema,
  c.relname  AS view,
  pg_get_viewdef(c.oid, true) AS definition,
  obj_description(c.oid, 'pg_class') AS comment
FROM pg_class c
JOIN pg_namespace ns ON ns.oid = c.relnamespace
WHERE c.relkind = 'v'
  AND ns.nspname NOT IN ('pg_catalog','information_schema')
ORDER BY 1,2;
"""

def fetch_df(conn, sql):
    with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
        cur.execute(sql)
        rows = cur.fetchall()
    return pd.DataFrame(rows)

def auto_width(writer, sheet_name, df, wrap_cols=None):
    ws = writer.sheets[sheet_name]
    wrap_cols = wrap_cols or []
    # ヘッダ＋セルの最大幅を計算（上限は見やすさ優先で 80）
    for idx, col in enumerate(df.columns, start=1):
        max_len = max([len(str(col))] + [len(str(x)) for x in df[col].head(1000)])  # フル走査は重いので上限
        ws.set_column(idx-1, idx-1, min(max_len + 2, 80))
        if col in wrap_cols:
            fmt = writer.book.add_format({"text_wrap": True, "valign": "top"})
            ws.set_column(idx-1, idx-1, min(max_len + 2, 80), fmt)
    # 1行目固定＆フィルタ
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(0, len(df)), max(0, len(df.columns)-1))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dsn", help="psycopg2 DSN (例: host=127.0.0.1 dbname=yourdb user=meta_reader password=*** port=5432)")
    ap.add_argument("--out", default="pg_metadata.xlsx", help="出力 xlsx ファイル名")
    ap.add_argument("--with-counts", action="store_true", help="Tables に row_count を追加（重い可能性あり）")
    args = ap.parse_args()

    dsn = args.dsn or os.getenv("PG_DSN")
    if not dsn:
        raise SystemExit("DSN 未指定です。--dsn か PG_DSN 環境変数を設定してください。")

    conn = psycopg2.connect(dsn)

    try:
        tables = fetch_df(conn, TABLES_SQL)
        cols   = fetch_df(conn, COLUMNS_SQL)
        cons   = fetch_df(conn, CONSTRAINTS_SQL)
        idxs   = fetch_df(conn, INDEXES_SQL)
        trigs  = fetch_df(conn, TRIGGERS_SQL)
        funcs  = fetch_df(conn, FUNCTIONS_SQL)
        views  = fetch_df(conn, VIEWS_SQL)

        if args.with_counts and not tables.empty:
            # スキーマ別/テーブル別に count(*) を回す（注意：時間かかる）
            counts = []
            with conn.cursor() as cur:
                for _, r in tables.iterrows():
                    schema = r["schema"]
                    table  = r["table"]
                    cur.execute(f'SELECT COUNT(*) FROM "{schema}"."{table}"')
                    n = cur.fetchone()[0]
                    counts.append(n)
            tables.insert(3, "row_count", counts)

        # Excel 書き出し
        with pd.ExcelWriter(args.out, engine="xlsxwriter") as writer:
            def write_sheet(name, df, wrap_cols=None, order=None):
                if df is None or df.empty:
                    pd.DataFrame({"info": ["(no rows)"]}).to_excel(writer, index=False, sheet_name=name)
                else:
                    if order:
                        exist = [c for c in order if c in df.columns]
                        rest  = [c for c in df.columns if c not in exist]
                        df = df[exist + rest]
                    df.to_excel(writer, index=False, sheet_name=name)
                auto_width(writer, name, df if not df.empty else pd.DataFrame({"info":[""]}), wrap_cols=wrap_cols)

            write_sheet("Tables",     tables, wrap_cols=["comment"], order=["schema","table","owner","row_count","comment"])
            write_sheet("Columns",    cols,   wrap_cols=["comment","default"], order=["schema","table","ordinal_position","column","data_type","nullable","default","identity","comment"])
            write_sheet("Constraints",cons,   wrap_cols=["definition"], order=["schema","table","constraint_name","type","definition"])
            write_sheet("Indexes",    idxs,   wrap_cols=["definition","predicate"], order=["schema","table","index_name","method","is_unique","predicate","definition"])
            write_sheet("Triggers",   trigs,  wrap_cols=["definition"], order=["schema","table","trigger_name","timing_events","function","definition"])
            write_sheet("Functions",  funcs,  wrap_cols=["arguments","comment"], order=["schema","name","arguments","return_type","volatility","security","comment"])
            write_sheet("Views",      views,  wrap_cols=["definition","comment"], order=["schema","view","comment","definition"])

        print(f"OK: {args.out}")
    finally:
        conn.close()

if __name__ == "__main__":
    main()
```

---

## 使い方

### 依存ライブラリ

```bash
# uv をお使いなら
uv pip install pandas psycopg2-binary xlsxwriter
# もしくは
pip install pandas psycopg2-binary xlsxwriter
```

### 実行例

```bash
# DSNを環境変数で渡す例（Windows PowerShell）
$env:PG_DSN = 'host=127.0.0.1 dbname=yourdb user=meta_reader password=*** port=5432'
python .\export_pg_metadata_to_excel.py --out meta_local.xlsx

# 直接DSNを指定
python export_pg_metadata_to_excel.py --dsn "host=127.0.0.1 dbname=yourdb user=meta_reader password=***" --out meta_local.xlsx

# 件数も出したい場合（時間かかる可能性）
python export_pg_metadata_to_excel.py --dsn "..." --out meta_local.xlsx --with-counts
```

---

## 運用のコツ

* **Alembicコメント運用**：`COMMENT ON TABLE/COLUMN` をマイグレーションに必ず含めると、**仕様書としてExcelが生きる**（`comment`列に反映）。
* **バージョン刻印**：Excelに「生成日時」「git rev-parse HEAD」「alembic heads」などを別シートに書くと、**どの時点のスキーマか**が明確になります（必要なら追記コード差し込みます）。
* **環境ごと比較**：開発・検証・本番をそれぞれ `meta_dev.xlsx / meta_stg.xlsx / meta_prod.xlsx` として出力し、Excelの**Power Query**や**VLOOKUP**で差分チェックが楽です。
* **安全性**：読み取り専用ユーザーを使い、スクリプトは**完全ローカル実行**。外部送信なし。

---

## 代替（より軽量）

* **CSV派**：psqlで `\copy (SQL) TO 'xxx.csv' CSV HEADER` をタスク化 → Excelで開く（ただし列幅や折返し、複数シートの整形は手作業になりがち）。
* **ERD生成**：Excelではなく図が要る場合は、ローカルだけで SchemaSpy / dbdiagram.io(ローカルCLI) 等を使う方法もあります（必要なら手順出します）。

---

必要なら「生成日時・Git/Alembic情報をInfoシートに追加」「ER図用の依存関係（view→table、function→table）を追加」「特定スキーマのみ対象」など、**用途に合わせて拡張版**をすぐ出せます。どこまで入れますか？
