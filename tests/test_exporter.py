from __future__ import annotations

import re
import sqlite3
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

import pandas as pd
from psycopg2 import sql

from pgsql_export import exporter


class _RecordingCursor:
    def __init__(self) -> None:
        self.executed: list[object] = []

    def __enter__(self) -> _RecordingCursor:
        return self

    def __exit__(self, *args: object) -> None:
        return None

    def execute(self, query: object) -> None:
        self.executed.append(query)

    def fetchone(self) -> tuple[int]:
        return (17,)


class _RecordingConnection:
    def __init__(self) -> None:
        self.cursor_instance = _RecordingCursor()

    def cursor(self) -> _RecordingCursor:
        return self.cursor_instance


class ExporterSqlTests(unittest.TestCase):
    def test_trigger_timing_uses_timing_bits_for_row_and_statement_triggers(self) -> None:
        timing_case = re.search(
            r"(CASE\s+WHEN \(tg\.tgtype & 64\).*?END),\s*CASE WHEN \(tg\.tgtype &\s*4\)",
            exporter.TRIGGERS_SQL,
            flags=re.DOTALL,
        )
        self.assertIsNotNone(timing_case)
        assert timing_case is not None

        db = sqlite3.connect(":memory:")
        self.addCleanup(db.close)

        cases = {
            4: "AFTER ",  # AFTER INSERT, statement-level
            5: "AFTER ",  # AFTER INSERT, row-level
            6: "BEFORE ",  # BEFORE INSERT, statement-level
            7: "BEFORE ",  # BEFORE INSERT, row-level
            81: "INSTEAD OF ",  # INSTEAD OF UPDATE, row-level
        }
        for trigger_type, expected in cases.items():
            with self.subTest(trigger_type=trigger_type):
                actual = db.execute(
                    f"SELECT {timing_case.group(1)} FROM (SELECT ? AS tgtype) AS tg",
                    (trigger_type,),
                ).fetchone()
                self.assertEqual(actual, (expected,))

    def test_metadata_queries_exclude_all_internal_schemas(self) -> None:
        queries_and_aliases = (
            (exporter.TABLES_SQL, "n"),
            (exporter.COLUMNS_SQL, "n"),
            (exporter.CONSTRAINTS_SQL, "n"),
            (exporter.INDEXES_SQL, "ns"),
            (exporter.TRIGGERS_SQL, "ns"),
            (exporter.FUNCTIONS_SQL, "ns"),
            (exporter.VIEWS_SQL, "ns"),
        )
        for query, alias in queries_and_aliases:
            with self.subTest(alias=alias, query=query[:30]):
                self.assertIn(f"{alias}.nspname !~ '^pg_'", query)
                self.assertIn(f"{alias}.nspname <> 'information_schema'", query)

    def test_row_count_query_composes_quoted_identifiers(self) -> None:
        connection = _RecordingConnection()
        tables = pd.DataFrame(
            [{"schema": 'reporting"archive', "table": 'monthly"totals'}]
        )

        counts = exporter._fetch_row_counts(connection, tables)  # type: ignore[arg-type]

        self.assertEqual(counts, [17])
        self.assertEqual(len(connection.cursor_instance.executed), 1)
        query = connection.cursor_instance.executed[0]
        self.assertIsInstance(query, sql.Composed)
        assert isinstance(query, sql.Composed)
        identifiers = [
            part.strings
            for part in query.seq
            if isinstance(part, sql.Identifier)
        ]
        self.assertEqual(
            identifiers,
            [('reporting"archive',), ('monthly"totals',)],
        )


class ExcelOutputTests(unittest.TestCase):
    def test_formula_and_url_like_text_are_not_auto_converted(self) -> None:
        tables = pd.DataFrame(
            [
                {
                    "schema": "public",
                    "table": "notes",
                    "owner": "https://example.invalid/owner",
                    "comment": "=1+1",
                }
            ]
        )
        query_results = [tables] + [pd.DataFrame()] * 6

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "metadata.xlsx"
            with patch.object(exporter, "fetch_df", side_effect=query_results):
                exporter.write_excel(object(), str(output))  # type: ignore[arg-type]

            with zipfile.ZipFile(output) as workbook:
                worksheet_xml = b"".join(
                    workbook.read(name)
                    for name in workbook.namelist()
                    if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
                )
                shared_strings = workbook.read("xl/sharedStrings.xml")

        self.assertNotIn(b"<f", worksheet_xml)
        self.assertNotIn(b"<hyperlink", worksheet_xml)
        self.assertIn(b"=1+1", shared_strings)
        self.assertIn(b"https://example.invalid/owner", shared_strings)


if __name__ == "__main__":
    unittest.main()
