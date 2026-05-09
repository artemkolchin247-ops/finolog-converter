"""
Unit tests for the Finolog converter.

Run:
    pytest test_app.py -v
"""

import sys
from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO
from types import ModuleType

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook


class _StreamlitStub(ModuleType):
    """Small Streamlit stub so tests can import app.py without running UI."""

    def __init__(self):
        super().__init__("streamlit")
        self.messages = []

    def set_page_config(self, *args, **kwargs):
        return None

    def title(self, *args, **kwargs):
        return None

    def file_uploader(self, *args, **kwargs):
        return None

    def error(self, message, *args, **kwargs):
        self.messages.append(("error", message))

    def warning(self, message, *args, **kwargs):
        self.messages.append(("warning", message))

    def stop(self):
        raise RuntimeError("streamlit.stop() called")


sys.modules["streamlit"] = _StreamlitStub()

from app import (  # noqa: E402
    build_amount_columns,
    detect_header_row,
    make_unique_columns,
    norm_text,
    parse_amount,
    process_excel,
)


def _excel_bytes(rows):
    """Create an in-memory xlsx file from raw rows without pandas headers."""
    buf = BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False, header=False)
    buf.seek(0)
    return buf


# ============================================================
# 1. Amount parsing
# ============================================================


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("1'000.00", Decimal("1000.00")),
        ("1\u2019000.00", Decimal("1000.00")),
        ("(500)", Decimal("-500")),
        ("(1,234.56)", Decimal("-1234.56")),
        ("1.234,56", Decimal("1234.56")),
        ("1,234.56", Decimal("1234.56")),
        ("1 000", Decimal("1000")),
        ("1\u00A0000,50", Decimal("1000.50")),
        ("1\u202F000,50", Decimal("1000.50")),
        ("1\u2009000,50", Decimal("1000.50")),
        ("500-", Decimal("-500")),
        ("1000 ₽", Decimal("1000")),
        ("$1,000.00", Decimal("1000.00")),
        ("1000 руб.", Decimal("1000")),
        ("—500", Decimal("-500")),
        (42, Decimal("42")),
        (0.1, Decimal("0.1")),
    ],
)
def test_parse_amount_accepts_supported_formats(raw, expected):
    val, reason, _ = parse_amount(raw)

    assert reason is None
    assert val == expected
    assert isinstance(val, Decimal)


@pytest.mark.parametrize(
    ("raw", "expected_reason"),
    [
        (None, "EMPTY"),
        (float("nan"), "EMPTY"),
        ("", "EMPTY"),
        ("-", "EMPTY"),
        ("none", "EMPTY"),
        ("nan", "EMPTY"),
        ("Да", "NON_AMOUNT"),
        ("Нет", "NON_AMOUNT"),
        ("true", "NON_AMOUNT"),
        ("false", "NON_AMOUNT"),
        ("текст", "NON_AMOUNT"),
        ("шт", "NON_AMOUNT"),
        ("₽", "NON_AMOUNT"),
        ("---", "NON_AMOUNT"),
        ("!!!", "NON_AMOUNT"),
    ],
)
def test_parse_amount_rejects_empty_and_non_amount_values(raw, expected_reason):
    val, reason, _ = parse_amount(raw)

    assert val is None
    assert reason == expected_reason


@pytest.mark.parametrize("raw", ["++100", "1-2", "12.34.56", "1,2,3"])
def test_parse_amount_reports_parse_fail_for_malformed_numeric_strings(raw):
    val, reason, debug = parse_amount(raw)

    assert val is None
    assert reason == "PARSE_FAIL"
    assert debug["exception"]


# ============================================================
# 2. Text/header helpers
# ============================================================


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        (None, ""),
        ("hello\u00A0world", "hello world"),
        ("1\u202F000", "1 000"),
        ("1\u2009000", "1 000"),
        ("  a   b  ", "a b"),
    ],
)
def test_norm_text(raw, expected):
    assert norm_text(raw) == expected


def test_make_unique_columns_suffixes_duplicate_headers():
    assert make_unique_columns(["Касса", "Касса", " Касса "]) == [
        "Касса",
        "Касса #2",
        "Касса #3",
    ]


def test_detect_header_row_finds_date_operation_header_with_spaces_and_case():
    df_raw = pd.DataFrame([
        ["meta", ""],
        ["  дата\u00A0операции  ", "Касса"],
    ])

    assert detect_header_row(df_raw) == 1


def test_build_amount_columns_excludes_known_non_amount_columns():
    df = pd.DataFrame(columns=[
        "Дата операции",
        "Дата начисления",
        "Описание",
        "Статья",
        "Касса",
        "Банк",
    ])

    assert build_amount_columns(df) == ["Касса", "Банк"]


# ============================================================
# 3. Decimal precision and rounding
# ============================================================


def test_decimal_math_does_not_lose_cents():
    values = [Decimal("0.01") for _ in range(10_000)]

    assert sum(values, Decimal(0)) == Decimal("100.00")
    assert Decimal("0.1") + Decimal("0.2") == Decimal("0.3")
    assert 0.1 + 0.2 != 0.3


def test_parsed_amounts_sum_exactly():
    values = [parse_amount("0.01")[0] for _ in range(10_000)]

    assert sum(values, Decimal(0)) == Decimal("100.00")


@pytest.mark.parametrize(
    ("value", "income", "expense"),
    [
        (Decimal("10.005"), Decimal("10.01"), Decimal("0")),
        (Decimal("-10.005"), Decimal("0"), Decimal("10.01")),
        (Decimal("0.004"), Decimal("0.00"), Decimal("0")),
        (Decimal("-0.004"), Decimal("0"), Decimal("0.00")),
    ],
)
def test_round_half_up_rules_used_by_converter(value, income, expense):
    two_places = Decimal("0.01")
    zero = Decimal(0)

    actual_income = value.quantize(two_places, rounding=ROUND_HALF_UP) if value > zero else zero
    actual_expense = (-value).quantize(two_places, rounding=ROUND_HALF_UP) if value < zero else zero

    assert actual_income == income
    assert actual_expense == expense


# ============================================================
# 4. Full Excel processing
# ============================================================


def test_process_excel_converts_wide_amount_columns_to_operations_and_sorts_by_dates():
    uploaded_file = _excel_bytes([
        ["service row", "", "", "", "", "", ""],
        ["№ п.п.", "Дата операции", "Дата начисления", "Описание", "Статья", "Касса", "Банк"],
        [1, "2026-04-20", "2026-04-20", "april", "Маркетинг", "", "300"],
        [2, "2025-12-19", "2025-12-01", "dec early pnl", "Логистика", "100", ""],
        [3, "2025-09-04", "2025-08-31", "sep", "Логистика", "-50", ""],
        [4, "2025-12-19", "2025-12-20", "dec late pnl", "Маркетинг", "200", ""],
    ])

    display_df, export_df, error_df = process_excel(uploaded_file)

    assert error_df.empty
    assert list(export_df.columns) == [
        "Дата ДДС",
        "Дата P&L",
        "Приход",
        "Расход",
        "Статья операции",
        "Касса / Счет",
        "Комментарий",
    ]
    assert display_df["Дата ДДС"].tolist() == [
        "04.09.2025",
        "19.12.2025",
        "19.12.2025",
        "20.04.2026",
    ]
    assert display_df["Дата P&L"].tolist() == [
        "31.08.2025",
        "01.12.2025",
        "20.12.2025",
        "20.04.2026",
    ]
    assert display_df["Комментарий"].tolist() == [
        "sep",
        "dec early pnl",
        "dec late pnl",
        "april",
    ]
    assert export_df["Приход"].tolist() == [
        Decimal("0"),
        Decimal("100.00"),
        Decimal("200.00"),
        Decimal("300.00"),
    ]
    assert export_df["Расход"].tolist() == [
        Decimal("50.00"),
        Decimal("0"),
        Decimal("0"),
        Decimal("0"),
    ]
    assert export_df["Касса / Счет"].tolist() == ["Касса", "Касса", "Касса", "Банк"]


def test_process_excel_creates_one_operation_per_non_empty_amount_cell():
    uploaded_file = _excel_bytes([
        ["Дата операции", "Дата начисления", "Описание", "Статья", "Касса", "Банк"],
        ["2026-01-10", "2026-01-09", "split", "Доход", "100", "-25"],
    ])

    display_df, export_df, error_df = process_excel(uploaded_file)

    assert error_df.empty
    assert len(export_df) == 2
    operations_by_account = {
        row["Касса / Счет"]: (row["Приход"], row["Расход"])
        for row in export_df.to_dict("records")
    }
    assert operations_by_account == {
        "Касса": (Decimal("100.00"), Decimal("0")),
        "Банк": (Decimal("0"), Decimal("25.00")),
    }


def test_process_excel_logs_rounded_to_zero_and_excludes_operation():
    uploaded_file = _excel_bytes([
        ["Дата операции", "Дата начисления", "Описание", "Статья", "Касса"],
        ["2026-01-10", "2026-01-10", "tiny", "Доход", "0.004"],
    ])

    display_df, export_df, error_df = process_excel(uploaded_file)

    assert display_df.empty
    assert export_df.empty
    assert len(error_df) == 1
    assert error_df.loc[0, "Причина_пропуска"] == "ROUNDED_TO_ZERO"
    assert error_df.loc[0, "Сумма_числом"] == Decimal("0.004")


def test_process_excel_skips_text_amounts_without_logging_operations():
    uploaded_file = _excel_bytes([
        ["Дата операции", "Дата начисления", "Описание", "Статья", "Касса"],
        ["2026-01-10", "2026-01-10", "text", "Доход", "не сумма"],
    ])

    result = process_excel(uploaded_file)

    assert result is None


def test_process_excel_returns_none_when_header_is_missing():
    uploaded_file = _excel_bytes([
        ["Дата платежа", "Касса"],
        ["2026-01-10", "100"],
    ])

    assert process_excel(uploaded_file) is None


# ============================================================
# 5. Excel roundtrip
# ============================================================


def test_decimal_values_can_be_written_to_excel_as_numbers_and_read_back():
    values = [Decimal("0.01").quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) for _ in range(100)]
    expected_total = sum(values, Decimal(0))
    df = pd.DataFrame({"Приход": [float(v) for v in values]})

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Test")
    buf.seek(0)

    df_read = pd.read_excel(buf, sheet_name="Test")
    actual_total = Decimal(str(df_read["Приход"].sum()))

    assert abs(actual_total - expected_total) < Decimal("0.01")


def test_zero_hidden_format_preserves_numeric_cell_type():
    buf = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Приход"
    ws["A2"] = 0
    ws["A2"].number_format = '#,##0.00;-#,##0.00;""'
    ws["A3"] = 100.50
    ws["A3"].number_format = '#,##0.00;-#,##0.00;""'
    wb.save(buf)
    buf.seek(0)

    wb2 = load_workbook(buf)
    ws2 = wb2.active

    assert ws2["A2"].value == 0
    assert isinstance(ws2["A2"].value, (int, float))
    assert ws2["A3"].value == 100.50
    assert isinstance(ws2["A3"].value, float)
