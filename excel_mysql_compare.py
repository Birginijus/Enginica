"""Compare worksheet records with rows in a MySQL table."""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import asdict, dataclass
from datetime import date, datetime
from decimal import Decimal
from typing import Any, Optional
from zipfile import BadZipFile


_IDENTIFIER = re.compile(r"^[A-Za-z_][A-Za-z0-9_$]*$")


@dataclass(frozen=True)
class ComparisonResult:
    exact_matches: list[dict[str, Any]]
    differing_rows: list[dict[str, Any]]
    excel_only: list[dict[str, Any]]
    mysql_only: list[dict[str, Any]]


def _validate_identifier(identifier: str, *, qualified: bool = False) -> str:
    parts = identifier.split(".") if qualified else [identifier]
    if (
        not parts
        or (qualified and len(parts) > 2)
        or any(not _IDENTIFIER.fullmatch(part) for part in parts)
    ):
        raise ValueError(f"Invalid SQL identifier: {identifier!r}")
    return ".".join(f"`{part}`" for part in parts)


def _comparable(value: Any) -> Any:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float, Decimal)):
        return Decimal(str(value))
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return value


def _index_rows(
    rows: Iterable[Mapping[str, Any]], identity_columns: Sequence[str], side: str
) -> dict[tuple[Any, ...], dict[str, Any]]:
    indexed: dict[tuple[Any, ...], dict[str, Any]] = {}
    for row_number, row in enumerate(rows, start=1):
        missing = [column for column in identity_columns if column not in row]
        if missing:
            raise ValueError(f"{side} row {row_number} is missing columns: {', '.join(missing)}")
        key = tuple(row[column] for column in identity_columns)
        if any(value is None or (isinstance(value, str) and not value.strip()) for value in key):
            raise ValueError(f"{side} row {row_number} has an empty identity value")
        comparable_key = tuple(_comparable(value) for value in key)
        if comparable_key in indexed:
            raise ValueError(f"Duplicate identity in {side} data: {key!r}")
        indexed[comparable_key] = dict(row)
    return indexed


def compare_rows(
    excel_rows: Iterable[Mapping[str, Any]],
    mysql_rows: Iterable[Mapping[str, Any]],
    identity_columns: Sequence[str],
) -> ComparisonResult:
    """Compare rows by identity, checking every column present in the worksheet."""
    if not identity_columns:
        raise ValueError("At least one identity column is required")
    # Materialize once so generators can be indexed and all worksheet headers retained.
    excel_rows = list(excel_rows) if not isinstance(excel_rows, list) else excel_rows
    mysql_rows = list(mysql_rows)
    columns = list(dict.fromkeys(column for row in excel_rows for column in row))
    missing_ids = [column for column in identity_columns if column not in columns]
    if missing_ids:
        raise ValueError(f"Excel data is missing identity columns: {', '.join(missing_ids)}")

    excel = _index_rows(excel_rows, identity_columns, "Excel")
    mysql = _index_rows(mysql_rows, identity_columns, "MySQL")
    for key, row in mysql.items():
        missing = [column for column in columns if column not in row]
        if missing:
            raise ValueError(f"MySQL row {key!r} is missing columns: {', '.join(missing)}")

    exact_matches = []
    differing_rows = []
    for key, excel_row in excel.items():
        mysql_row = mysql.get(key)
        if mysql_row is None:
            continue
        differences = {
            column: {"excel": excel_row.get(column), "mysql": mysql_row.get(column)}
            for column in columns
            if _comparable(excel_row.get(column)) != _comparable(mysql_row.get(column))
        }
        identity = dict(zip(identity_columns, [excel_row[column] for column in identity_columns]))
        if differences:
            differing_rows.append({"identity": identity, "differences": differences})
        else:
            exact_matches.append(identity)

    return ComparisonResult(
        exact_matches=exact_matches,
        differing_rows=differing_rows,
        excel_only=[row for key, row in excel.items() if key not in mysql],
        mysql_only=[row for key, row in mysql.items() if key not in excel],
    )


def read_excel_rows(path: str, sheet_name: str) -> list[dict[str, Any]]:
    from openpyxl import load_workbook
    from openpyxl.utils.exceptions import InvalidFileException

    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except (BadZipFile, InvalidFileException) as error:
        raise ValueError(f"Unable to read Excel workbook {path!r}: {error}") from error
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(
                f"Worksheet {sheet_name!r} not found. Available sheets: "
                f"{', '.join(workbook.sheetnames)}"
            )
        worksheet = workbook[sheet_name]
        iterator = worksheet.iter_rows(values_only=True)
        headers = next(iterator, None)
        if not headers or all(header is None for header in headers):
            raise ValueError("The worksheet must have a header row")
        names = [str(header).strip() if header is not None else "" for header in headers]
        if any(not name for name in names):
            raise ValueError("Worksheet headers must not be blank")
        if len(set(names)) != len(names):
            raise ValueError("Worksheet headers must be unique")

        rows = []
        for values in iterator:
            if all(value is None for value in values):
                continue
            rows.append(dict(zip(names, values)))
        return rows
    finally:
        workbook.close()


def fetch_mysql_rows(
    *,
    host: str,
    port: int,
    database: str,
    table: str,
    user: str,
    password: str,
    columns: Sequence[str],
) -> list[dict[str, Any]]:
    import pymysql

    quoted_table = _validate_identifier(table, qualified=True)
    if not columns:
        raise ValueError("No worksheet columns were found")
    quoted_columns = ", ".join(_validate_identifier(column) for column in columns)
    try:
        connection = pymysql.connect(
            host=host,
            port=port,
            user=user,
            password=password,
            database=database,
            cursorclass=pymysql.cursors.DictCursor,
            read_timeout=30,
            write_timeout=30,
            autocommit=False,
        )
    except pymysql.MySQLError as error:
        raise RuntimeError(f"MySQL connection failed: {error}") from error
    try:
        with connection.cursor() as cursor:
            cursor.execute("START TRANSACTION READ ONLY")
            cursor.execute(f"SELECT {quoted_columns} FROM {quoted_table}")
            return list(cursor.fetchall())
    except pymysql.MySQLError as error:
        raise RuntimeError(f"MySQL query failed: {error}") from error
    finally:
        connection.close()


def _json_value(value: Any) -> Any:
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return str(value)
    return value


def _format_text(result: ComparisonResult) -> str:
    lines = [
        "Comparison results",
        f"Exact matches: {len(result.exact_matches)}",
        f"Differing rows: {len(result.differing_rows)}",
        f"Excel-only rows: {len(result.excel_only)}",
        f"MySQL-only rows: {len(result.mysql_only)}",
    ]
    for item in result.differing_rows:
        lines.append(f"  Difference for {item['identity']}:")
        for column, values in item["differences"].items():
            lines.append(f"    {column}: Excel={values['excel']!r}, MySQL={values['mysql']!r}")
    for label, rows in (("Excel-only", result.excel_only), ("MySQL-only", result.mysql_only)):
        for row in rows:
            lines.append(f"  {label}: {row}")
    return "\n".join(lines)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Compare Excel worksheet rows with a MySQL table by identity columns."
    )
    parser.add_argument("--excel", required=True, help="Path to the Excel workbook")
    parser.add_argument("--sheet", required=True, help="Worksheet name")
    parser.add_argument("--table", required=True, help="MySQL table (optionally schema-qualified)")
    parser.add_argument(
        "--id-columns",
        required=True,
        help="Comma-separated identity column names, e.g. customer_id or tenant_id,customer_id",
    )
    parser.add_argument("--mysql-host", default=os.getenv("MYSQL_HOST", "localhost"))
    parser.add_argument("--mysql-port", type=int, default=int(os.getenv("MYSQL_PORT", "3306")))
    parser.add_argument("--mysql-database", default=os.getenv("MYSQL_DATABASE"))
    parser.add_argument("--mysql-user", default=os.getenv("MYSQL_USER"))
    parser.add_argument(
        "--output", choices=("text", "json"), default="text", help="Report format (default: text)"
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    missing = [
        name
        for name, value in (
            ("MYSQL_DATABASE / --mysql-database", args.mysql_database),
            ("MYSQL_USER / --mysql-user", args.mysql_user),
            ("MYSQL_PASSWORD environment variable", os.getenv("MYSQL_PASSWORD")),
        )
        if not value
    ]
    if missing:
        parser.error("Missing MySQL configuration: " + ", ".join(missing))

    identity_columns = [column.strip() for column in args.id_columns.split(",") if column.strip()]
    try:
        excel_rows = read_excel_rows(args.excel, args.sheet)
        columns = list(dict.fromkeys(column for row in excel_rows for column in row))
        mysql_rows = fetch_mysql_rows(
            host=args.mysql_host,
            port=args.mysql_port,
            database=args.mysql_database,
            table=args.table,
            user=args.mysql_user,
            password=os.environ["MYSQL_PASSWORD"],
            columns=columns,
        )
        result = compare_rows(excel_rows, mysql_rows, identity_columns)
    except (OSError, ValueError, ImportError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 2
    except RuntimeError as error:
        print(f"Database error: {error}", file=sys.stderr)
        return 2

    if args.output == "json":
        print(json.dumps(asdict(result), default=_json_value, indent=2))
    else:
        print(_format_text(result))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
