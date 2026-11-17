import os
import sys
import tempfile
import unittest
from decimal import Decimal
from types import SimpleNamespace
from unittest.mock import MagicMock, call, patch

from openpyxl import Workbook

from excel_mysql_compare import (
    _validate_identifier,
    compare_rows,
    fetch_mysql_rows,
    read_excel_rows,
)


class CompareRowsTests(unittest.TestCase):
    def test_reports_matches_differences_and_rows_on_each_side(self):
        result = compare_rows(
            [
                {"id": 1, "name": "same", "amount": 2},
                {"id": 2, "name": "Excel", "amount": 4},
                {"id": 3, "name": "only Excel", "amount": 5},
            ],
            [
                {"id": Decimal("1"), "name": "same", "amount": Decimal("2.0")},
                {"id": 2, "name": "MySQL", "amount": 4},
                {"id": 4, "name": "only MySQL", "amount": 6},
            ],
            ["id"],
        )

        self.assertEqual(result.exact_matches, [{"id": 1}])
        self.assertEqual(
            result.differing_rows,
            [{"identity": {"id": 2}, "differences": {"name": {"excel": "Excel", "mysql": "MySQL"}}}],
        )
        self.assertEqual(result.excel_only, [{"id": 3, "name": "only Excel", "amount": 5}])
        self.assertEqual(result.mysql_only, [{"id": 4, "name": "only MySQL", "amount": 6}])

    def test_rejects_duplicate_or_empty_identities(self):
        with self.assertRaisesRegex(ValueError, "Duplicate identity"):
            compare_rows([{"id": 1}, {"id": 1}], [], ["id"])
        with self.assertRaisesRegex(ValueError, "empty identity"):
            compare_rows([{"id": None}], [], ["id"])
        with self.assertRaisesRegex(ValueError, "empty identity"):
            compare_rows([{"id": " "}], [], ["id"])

    def test_accepts_iterators(self):
        excel = iter([{"id": 1, "value": "a"}])
        mysql = iter([{"id": 1, "value": "a"}])
        result = compare_rows(excel, mysql, ["id"])
        self.assertEqual(result.exact_matches, [{"id": 1}])

    def test_rejects_unsafe_identifiers(self):
        self.assertEqual(_validate_identifier("report.items", qualified=True), "`report`.`items`")
        with self.assertRaisesRegex(ValueError, "Invalid SQL identifier"):
            _validate_identifier("items; DROP TABLE users")

    def test_reads_named_sheet_and_skips_blank_rows(self):
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Data"
        worksheet.append(["id", "value"])
        worksheet.append([1, "a"])
        worksheet.append([None, None])
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as file:
            path = file.name
        try:
            workbook.save(path)
            self.assertEqual(
                read_excel_rows(path, "Data"),
                [{"id": 1, "value": "a"}],
            )
        finally:
            workbook.close()
            os.unlink(path)


class FetchMysqlRowsTests(unittest.TestCase):
    def test_uses_read_only_transaction_and_quoted_select(self):
        cursor = MagicMock()
        cursor.fetchall.return_value = [{"id": 1}]
        connection = MagicMock()
        connection.cursor.return_value.__enter__.return_value = cursor
        mysql = SimpleNamespace(
            MySQLError=Exception,
            cursors=SimpleNamespace(DictCursor=object()),
            connect=MagicMock(return_value=connection),
        )

        with patch.dict(sys.modules, {"pymysql": mysql}):
            rows = fetch_mysql_rows(
                host="localhost",
                port=3306,
                database="example",
                table="data.records",
                user="reader",
                password="secret",
                columns=["record_id", "display_name"],
            )

        self.assertEqual(rows, [{"id": 1}])
        self.assertEqual(
            cursor.execute.call_args_list,
            [
                call("START TRANSACTION READ ONLY"),
                call("SELECT `record_id`, `display_name` FROM `data`.`records`"),
            ],
        )
        self.assertFalse(mysql.connect.call_args.kwargs["autocommit"])
        connection.close.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
