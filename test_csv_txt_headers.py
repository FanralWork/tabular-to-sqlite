import sqlite3
import unittest
from contextlib import closing
from pathlib import Path
from tempfile import TemporaryDirectory

from data_processor_optimized import DataProcessorOptimized


class CsvTxtHeaderTests(unittest.TestCase):
    def convert(self, suffix, rows, chunk_size, **options):
        delimiter = "," if suffix == "csv" else "|"
        with TemporaryDirectory() as directory:
            source = Path(directory) / f"sample.{suffix}"
            target = Path(directory) / "sample.db"
            with source.open("w", encoding="utf-8") as stream:
                for row in rows:
                    stream.write(delimiter.join(row) + "\n")

            if suffix == "txt":
                options["txt_delimiter"] = delimiter
            processor = DataProcessorOptimized({"chunk_size": chunk_size})
            processor.convert_file_to_db(str(source), str(target), **options)

            with closing(sqlite3.connect(target)) as conn:
                with closing(conn.execute("SELECT * FROM sample ORDER BY rowid")) as cursor:
                    return [column[0] for column in cursor.description], cursor.fetchall()

    def test_header_supplies_column_names_only(self):
        rows = [("Label", "Value"), ("first", "1"), ("second", "2"), ("third", "3")]
        for suffix in ("csv", "txt"):
            for chunk_size in (1, 2):
                with self.subTest(suffix=suffix, chunk_size=chunk_size):
                    columns, records = self.convert(
                        suffix, rows, chunk_size,
                        use_first_row_as_header=True,
                        column_mapping={"Label": "Name"},
                    )
                    self.assertEqual(columns, ["Name", "Value"])
                    self.assertEqual(records, [("first", 1), ("second", 2), ("third", 3)])

    def test_no_header_keeps_first_record(self):
        rows = [("first", "1"), ("second", "2"), ("third", "3")]
        for suffix in ("csv", "txt"):
            with self.subTest(suffix=suffix):
                columns, records = self.convert(suffix, rows, chunk_size=2)
                self.assertEqual(columns, ["Column_0", "Column_1"])
                self.assertEqual(records, [("first", 1), ("second", 2), ("third", 3)])

    def test_header_only_creates_empty_table(self):
        for suffix in ("csv", "txt"):
            with self.subTest(suffix=suffix):
                columns, records = self.convert(
                    suffix, [("Label", "Value")], chunk_size=2, use_first_row_as_header=True,
                )
                self.assertEqual(columns, ["Label", "Value"])
                self.assertEqual(records, [])

    def test_header_keeps_column_with_missing_values_in_first_chunk(self):
        rows = [("Label", "Value"), ("first", ""), ("second", "2")]
        columns, records = self.convert("csv", rows, chunk_size=1, use_first_row_as_header=True)
        self.assertEqual(columns, ["Label", "Value"])
        self.assertEqual(records, [("first", None), ("second", 2)])
