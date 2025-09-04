from __future__ import annotations

import pandas as pd
import psycopg2
import psycopg2.extras


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
  "schema", "table", constraint_name,
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


def fetch_df(conn: psycopg2.extensions.connection, sql: str) -> pd.DataFrame:
    with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
        cur.execute(sql)
        rows = cur.fetchall()
    return pd.DataFrame(rows)


def _auto_width(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame, wrap_cols=None):
    ws = writer.sheets[sheet_name]
    wrap_cols = set(wrap_cols or [])
    for idx, col in enumerate(df.columns, start=1):
        sample_vals = df[col].head(1000)
        max_len = max([len(str(col))] + [len(str(x)) for x in sample_vals]) if not df.empty else len(str(col))
        width = min(max_len + 2, 80)
        cell_fmt = None
        if col in wrap_cols:
            cell_fmt = writer.book.add_format({"text_wrap": True, "valign": "top"})
        ws.set_column(idx - 1, idx - 1, width, cell_fmt)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(0, len(df)), max(0, len(df.columns) - 1))


def write_excel(conn: psycopg2.extensions.connection, out_path: str, with_counts: bool = False):
    tables = fetch_df(conn, TABLES_SQL)
    cols = fetch_df(conn, COLUMNS_SQL)
    cons = fetch_df(conn, CONSTRAINTS_SQL)
    idxs = fetch_df(conn, INDEXES_SQL)
    trigs = fetch_df(conn, TRIGGERS_SQL)
    funcs = fetch_df(conn, FUNCTIONS_SQL)
    views = fetch_df(conn, VIEWS_SQL)

    if with_counts and not tables.empty:
        counts = []
        with conn.cursor() as cur:
            for _, r in tables.iterrows():
                schema = r["schema"]
                table = r["table"]
                cur.execute(f'SELECT COUNT(*) FROM "{schema}"."{table}"')
                # psycopg2's fetchone() returns tuple[Any, ...] | None per stubs.
                # COUNT(*) always returns a row, but guard for mypy and safety.
                row = cur.fetchone()
                if row is None:
                    counts.append(0)
                else:
                    counts.append(row[0])
        tables.insert(3, "row_count", counts)

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        def write_sheet(name: str, df: pd.DataFrame, wrap_cols=None, order=None):
            if df is None or df.empty:
                view_df = pd.DataFrame({"info": ["(no rows)"]})
                view_df.to_excel(writer, index=False, sheet_name=name)
                _auto_width(writer, name, view_df, wrap_cols=wrap_cols)
                return
            view_df = df.copy()
            if order:
                exist = [c for c in order if c in view_df.columns]
                rest = [c for c in view_df.columns if c not in exist]
                view_df = view_df[exist + rest]
            view_df.to_excel(writer, index=False, sheet_name=name)
            _auto_width(writer, name, view_df, wrap_cols=wrap_cols)

        write_sheet("Tables", tables, wrap_cols=["comment"], order=["schema", "table", "owner", "row_count", "comment"])
        write_sheet(
            "Columns",
            cols,
            wrap_cols=["comment", "default"],
            order=["schema", "table", "ordinal_position", "column", "data_type", "nullable", "default", "identity", "comment"],
        )
        write_sheet("Constraints", cons, wrap_cols=["definition"], order=["schema", "table", "constraint_name", "type", "definition"])
        write_sheet(
            "Indexes",
            idxs,
            wrap_cols=["definition", "predicate"],
            order=["schema", "table", "index_name", "method", "is_unique", "predicate", "definition"],
        )
        write_sheet(
            "Triggers",
            trigs,
            wrap_cols=["definition"],
            order=["schema", "table", "trigger_name", "timing_events", "function", "definition"],
        )
        write_sheet(
            "Functions",
            funcs,
            wrap_cols=["arguments", "comment"],
            order=["schema", "name", "arguments", "return_type", "volatility", "security", "comment"],
        )
        write_sheet("Views", views, wrap_cols=["definition", "comment"], order=["schema", "view", "comment", "definition"])
