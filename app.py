import pandas as pd
import streamlit as st
import re
from io import BytesIO
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation

# =========================
# Helpers
# =========================

def norm_text(x) -> str:
    """Normalize header/cell text for comparisons."""
    if x is None:
        return ""
    s = str(x)
    # normalize spaces (regular + NBSP + narrow NBSP + thin space)
    s = s.replace("\u00A0", " ").replace("\u202F", " ").replace("\u2009", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def norm_key(x) -> str:
    """Lowercased normalized key for header matching."""
    return norm_text(x).lower()

def get_scalar(value):
    """If pandas returns Series, take first element; else return as-is."""
    if isinstance(value, pd.Series):
        value = value.iloc[0]
    return value

def parse_amount(val):
    """
    Возвращает:
      value: Decimal | None
      reason: str | None   # None = успешно; EMPTY/NON_AMOUNT/PARSE_FAIL = не сумма
      debug: dict
    """
    debug = {
        "raw": val,
        "raw_str": None,
        "normalized": None,
        "rule": None,
        "exception": None,
    }

    # None/NaN
    if val is None or (isinstance(val, float) and pd.isna(val)):
        debug["rule"] = "EMPTY:None/NaN"
        return None, "EMPTY", debug

    # already numeric — через str, чтобы избежать float-артефактов
    if isinstance(val, (int, float)) and not pd.isna(val):
        try:
            d = Decimal(str(val))
            debug["raw_str"] = str(val)
            debug["normalized"] = str(d)
            debug["rule"] = "OK:already_numeric"
            return d, None, debug
        except InvalidOperation as e:
            debug["rule"] = "FAIL:decimal_cast"
            debug["exception"] = repr(e)
            return None, "PARSE_FAIL", debug

    s = norm_text(val)
    debug["raw_str"] = s

    if s == "" or s.lower() in {"-", "none", "nan"}:
        debug["rule"] = "EMPTY:blank_or_marker"
        return None, "EMPTY", debug

    # common non-amount markers
    if s.lower() in {"да", "нет", "true", "false"}:
        debug["rule"] = "NON_AMOUNT_MARKER"
        return None, "NON_AMOUNT", debug

    # normalize minus variants
    s = s.replace("\u2212", "-").replace("\u2013", "-").replace("\u2014", "-")

    # parentheses negative
    is_paren_negative = s.startswith("(") and s.endswith(")")
    if is_paren_negative:
        s = s[1:-1].strip()
        debug["rule"] = "RULE:paren_negative"

    # remove spaces (including NBSP variants + thin space)
    s = s.replace("\u00A0", "").replace("\u202F", "").replace("\u2009", "")
    s = s.replace(" ", "")

    # remove apostrophes (швейцарский формат 1'000)
    s = s.replace("\u2019", "").replace("'", "")

    # IMPORTANT: strip everything except digits, separators, and sign
    # this removes ₽ $ € and any other letters/symbols safely
    s = re.sub(r"[^0-9,\.\-\+]", "", s)

    # if nothing numeric left -> not an amount
    if not re.search(r"\d", s):
        debug["rule"] = "NON_AMOUNT:no_digits"
        debug["normalized"] = s
        return None, "NON_AMOUNT", debug

    # trailing minus: 1000- -> -1000
    if s.endswith("-") and re.fullmatch(r"[0-9\.,]+-", s):
        s = "-" + s[:-1]
        debug["rule"] = (debug["rule"] + "|RULE:trailing_minus") if debug["rule"] else "RULE:trailing_minus"

    # handle mixed separators: 1.234,56 vs 1,234.56
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            # EU: 1.234,56 -> 1234.56
            s = s.replace(".", "").replace(",", ".")
            tag = "RULE:EU_mixed_sep"
        else:
            # US: 1,234.56 -> 1234.56
            s = s.replace(",", "")
            tag = "RULE:US_mixed_sep"
        debug["rule"] = (debug["rule"] + "|" + tag) if debug["rule"] else tag
    else:
        # single sep: comma as decimal
        s = s.replace(",", ".")
        tag = "RULE:single_sep"
        debug["rule"] = (debug["rule"] + "|" + tag) if debug["rule"] else tag

    debug["normalized"] = s

    # final cast — Decimal
    try:
        num = Decimal(s)
        if is_paren_negative:
            num = -abs(num)
        debug["rule"] = ("OK:parsed|" + debug["rule"]) if debug["rule"] else "OK:parsed"
        return num, None, debug
    except (InvalidOperation, ValueError) as e:
        debug["rule"] = "FAIL:decimal_cast"
        debug["exception"] = repr(e)
        return None, "PARSE_FAIL", debug

def make_unique_columns(cols):
    seen = {}
    out = []

    for c in cols:
        c = norm_text(c)

        if c not in seen:
            seen[c] = 1
            out.append(c)
        else:
            seen[c] += 1
            out.append(f"{c} #{seen[c]}")

    return out
    
def detect_header_row(df_raw: pd.DataFrame) -> int | None:
    """
    Find row index that contains 'Дата операции' (case-insensitive, tolerant to spaces).
    """
    target = "дата операции"
    for i in range(len(df_raw)):
        row_vals = [norm_key(v) for v in df_raw.iloc[i].tolist()]
        if target in row_vals:
            return i
    return None

def build_amount_columns(df: pd.DataFrame, headers=None):
    exclude_cols = {
        "дата операции",
        "описание",
        "статья",
        "дата начисления",
        "",
        "nan",
    }
    return [c for c in df.columns if norm_key(c) not in exclude_cols]
# =========================
# Core processing
# =========================

def process_excel(uploaded_file):
    # Read raw without header
    df_raw = pd.read_excel(uploaded_file, header=None)

    header_idx = detect_header_row(df_raw)
    if header_idx is None:
        st.error("Не найдена строка с заголовками (нет 'Дата операции').")
        return None

    # Build headers from that row (keep original names, but strip)
    headers = df_raw.iloc[header_idx].apply(norm_text)
    headers = make_unique_columns(headers)

    df = df_raw.iloc[header_idx + 1:].copy()
    df.columns = headers

    # Drop known irrelevant cols (tolerant: compare by normalized key)
    cols_to_drop_keys = {
        "№  п.п.",
        "№ п.п.",
        "заказ",
        "проект",
        "контрагент",
        "реквизит контрагента",
        "компания",
        "id",
    }
    drop_cols = []
    for c in df.columns:
        if norm_key(c) in {norm_key(x) for x in cols_to_drop_keys}:
            drop_cols.append(c)
    if drop_cols:
        df = df.drop(columns=drop_cols, errors="ignore")

    # Map actual column names for key fields (tolerant)
    col_map = {norm_key(c): c for c in df.columns}
    col_date_op = col_map.get("дата операции")
    col_date_pl = col_map.get("дата начисления")
    col_desc = col_map.get("описание")
    col_article = col_map.get("статья")

    if not col_date_op:
        st.error("В таблице после заголовка нет колонки 'Дата операции' (проверь формат файла).")
        return None

    amount_cols = build_amount_columns(df)

    result_rows = []
    error_rows = []

    # Iterate source rows
    for src_idx, row in df.iterrows():
        date_dds = get_scalar(row.get(col_date_op, pd.NaT))
        date_pl = get_scalar(row.get(col_date_pl, pd.NaT)) if col_date_pl else pd.NaT
        description = norm_text(get_scalar(row.get(col_desc, ""))) if col_desc else ""
        article = norm_text(get_scalar(row.get(col_article, ""))) if col_article else ""

        # Iterate amount columns
        for col in amount_cols:
            if col in {col_date_op, col_date_pl, col_desc, col_article}:
                continue

            raw_val = get_scalar(row.get(col))

            value, reason, dbg = parse_amount(raw_val)

            if reason in {"EMPTY", "NON_AMOUNT"}:
                continue

            # если парсинг не удался — пишем точную причину и лог
            if reason is not None:
                error_rows.append({
                    "Источник_строка": int(src_idx) + 1,
                    "Источник_колонка": norm_text(col),
            
                    "Сумма_как_в_файле": dbg.get("raw_str"),
                    "Сумма_нормализованная": dbg.get("normalized"),
                    "Сумма_числом": "",
            
                    "Причина_пропуска": reason,
                    "Лог": f"{dbg.get('rule')} | exception={dbg.get('exception')}",
            
                    "Дата ДДС (как в файле)": date_dds,
                    "Дата P&L (как в файле)": date_pl,
                    "Статья": article,
                    "Комментарий": description,
                })
                continue
            
            _zero = Decimal(0)
            _two_places = Decimal("0.01")
            income = value.quantize(_two_places, rounding=ROUND_HALF_UP) if value > _zero else _zero
            expense = (-value).quantize(_two_places, rounding=ROUND_HALF_UP) if value < _zero else _zero
            
            # если число есть, но после округления стало 0 — логируем и НЕ добавляем в операции
            if income == _zero and expense == _zero and abs(value) > _zero:
                error_rows.append({
                    "Источник_строка": int(src_idx) + 1,
                    "Источник_колонка": norm_text(col),
            
                    "Сумма_как_в_файле": dbg.get("raw_str"),
                    "Сумма_нормализованная": dbg.get("normalized"),
                    "Сумма_числом": value,
            
                    "Причина_пропуска": "ROUNDED_TO_ZERO",
                    "Лог": f"value={value} -> income={income}, expense={expense} (round to 2 decimals)",
            
                    "Дата ДДС (как в файле)": date_dds,
                    "Дата P&L (как в файле)": date_pl,
                    "Статья": article,
                    "Комментарий": description,
                })
                continue
            
            result_rows.append({
                "Источник_строка": int(src_idx) + 1,
                "Источник_колонка": norm_text(col),
            
                "Дата ДДС": date_dds,
                "Дата P&L": date_pl,
                "Приход": income,
                "Расход": expense,
                "Статья операции": article,
                "Касса / Счет": norm_text(col),
                "Комментарий": description
            })

    if not result_rows and not error_rows:
        st.warning("Нет данных для отображения (пустой файл или формат не распознан).")
        return None

    result_df = pd.DataFrame(result_rows) if result_rows else pd.DataFrame(
    columns=["Источник_строка","Источник_колонка","Дата ДДС","Дата P&L","Приход","Расход","Статья операции","Касса / Счет","Комментарий"]
    )
    error_df = pd.DataFrame(error_rows) if error_rows else pd.DataFrame(
        columns=["Источник_строка","Источник_колонка","Сумма_как_в_файле","Сумма_нормализованная","Сумма_числом","Причина_пропуска","Лог","Дата ДДС (как в файле)","Дата P&L (как в файле)","Статья","Комментарий"]
    )
    
    # Normalize dates for sorting/export
    if not result_df.empty:
        result_df["Дата ДДС"] = pd.to_datetime(result_df["Дата ДДС"], errors="coerce")
        result_df["Дата P&L"] = pd.to_datetime(result_df["Дата P&L"], errors="coerce")
    
        # ====== СХЛОПЫВАНИЕ: совпадает всё, кроме Комментария и суммы ======
    
        # Строгая проверка: если в колонку попало не-Decimal — логируем и удаляем
        _zero_d = Decimal(0)
        _two_places = Decimal("0.01")

        for col_name in ["Приход", "Расход"]:
            invalid_mask = result_df[col_name].apply(
                lambda v: not isinstance(v, (Decimal, int)) or (isinstance(v, float) and pd.isna(v))
            )
            if invalid_mask.any():
                for idx in result_df[invalid_mask].index:
                    bad_row = result_df.loc[idx]
                    error_rows.append({
                        "Источник_строка": bad_row.get("Источник_строка", "N/A"),
                        "Источник_колонка": bad_row.get("Источник_колонка", "N/A"),
                        "Сумма_как_в_файле": str(bad_row[col_name]),
                        "Сумма_нормализованная": "",
                        "Сумма_числом": "",
                        "Причина_пропуска": "INVALID_TYPE_BEFORE_AGG",
                        "Лог": f"Колонка '{col_name}': тип {type(bad_row[col_name]).__name__}, значение={bad_row[col_name]}",
                        "Дата ДДС (как в файле)": bad_row.get("Дата ДДС", ""),
                        "Дата P&L (как в файле)": bad_row.get("Дата P&L", ""),
                        "Статья": bad_row.get("Статья операции", ""),
                        "Комментарий": bad_row.get("Комментарий", ""),
                    })
                result_df = result_df[~invalid_mask].copy()
    
        # убираем строки без суммы
        result_df = result_df[~((result_df["Приход"] == _zero_d) & (result_df["Расход"] == _zero_d))].copy()

        # нормализуем комментарий
        result_df["Комментарий"] = result_df["Комментарий"].fillna("").astype(str).str.strip()

        # ====== СЕЛЕКТИВНАЯ АГРЕГАЦИЯ ======
        # Крупные операции (≥1000) идут отдельными строками.
        # Мелкие (<1000) и "Перевод между счетами" — схлопываются.

        _agg_limit = Decimal("1000")
        result_df["_amount"] = result_df[["Приход", "Расход"]].max(axis=1)
        result_df["_is_agg"] = (
            (result_df["_amount"] < _agg_limit) |
            (result_df["Статья операции"] == "Перевод между счетами")
        )

        df_solo = result_df[~result_df["_is_agg"]].copy()
        df_to_agg = result_df[result_df["_is_agg"]].copy()

        # --- Агрегация мелких / переводов ---
        def join_comments(series: pd.Series) -> str:
            seen = set()
            out = []
            for x in series.tolist():
                x = (x or "").strip()
                if not x:
                    continue
                if x not in seen:
                    seen.add(x)
                    out.append(x)
            return "; ".join(out)

        def decimal_sum(series: pd.Series) -> Decimal:
            """Суммирование через Python sum — Decimal-safe, без конвертации в float."""
            return sum(series.tolist(), Decimal(0))

        group_keys = ["Дата ДДС", "Дата P&L", "Статья операции", "Касса / Счет"]

        if not df_to_agg.empty:
            df_grouped = (
                df_to_agg
                .groupby(group_keys, dropna=False, as_index=False)
                .agg({
                    "Приход": decimal_sum,
                    "Расход": decimal_sum,
                    "Комментарий": join_comments
                })
            )
            # Финальное округление агрегированных сумм
            df_grouped["Приход"] = df_grouped["Приход"].apply(
                lambda v: v.quantize(_two_places, rounding=ROUND_HALF_UP) if isinstance(v, Decimal) else Decimal(0)
            )
            df_grouped["Расход"] = df_grouped["Расход"].apply(
                lambda v: v.quantize(_two_places, rounding=ROUND_HALF_UP) if isinstance(v, Decimal) else Decimal(0)
            )
            # Источник_строка для агрегированных строк — пусто (несколько строк склеены)
            df_grouped["Источник_строка"] = ""
            df_grouped["Источник_колонка"] = ""
        else:
            df_grouped = pd.DataFrame()

        # --- Объединяем соло + агрегированные ---
        result_df = pd.concat([df_solo, df_grouped], ignore_index=True)

        # убираем нулевые строки после агрегации
        result_df = result_df[~((result_df["Приход"] == _zero_d) & (result_df["Расход"] == _zero_d))].copy()

        # убираем технические колонки
        result_df = result_df.drop(columns=["_amount", "_is_agg"], errors="ignore")

        # сортировка
        result_df = result_df.sort_values(by=["Дата ДДС", "Дата P&L"]).reset_index(drop=True)

    # Display df (human-friendly)
    # --- фиксируем порядок колонок ---
    desired_order = [
        "Источник_строка",
        "Дата ДДС",
        "Дата P&L",
        "Приход",
        "Расход",
        "Статья операции",
        "Касса / Счет",
        "Комментарий",
    ]
    
    # если какие-то колонки вдруг отсутствуют — не упадём
    existing = [c for c in desired_order if c in result_df.columns]
    result_df = result_df[existing]

    # Decimal -> float для отображения в Streamlit (st.dataframe не поддерживает Decimal)
    display_df = result_df.copy()
    if not display_df.empty:
        display_df["Дата ДДС"] = display_df["Дата ДДС"].dt.strftime("%d.%m.%Y").fillna("")
        display_df["Дата P&L"] = display_df["Дата P&L"].dt.strftime("%d.%m.%Y").fillna("")
        display_df["Приход"] = display_df["Приход"].apply(lambda x: "" if x == Decimal(0) else float(x))
        display_df["Расход"] = display_df["Расход"].apply(lambda x: "" if x == Decimal(0) else float(x))
    
    return display_df, result_df, error_df

# =========================
# Streamlit UI
# =========================

st.set_page_config(page_title="Конвертер Финолог → Управленческий учет", layout="wide")
st.title("Конвертер Финолог → Управленческий учет")

uploaded_file = st.file_uploader("📂 Загрузите Excel-файл", type=["xlsx"])

if uploaded_file:
    try:
        out = process_excel(uploaded_file)
        if out is None:
            st.stop()

        display_df, export_df, error_df = out

        # Tabs: results + errors
        tab1, tab2 = st.tabs(["Операции", f"Ошибки/пропуски ({len(error_df)})"])

        with tab1:
            if export_df.empty:
                st.info("Операции не сформированы. Проверь вкладку 'Ошибки/пропуски'.")
            else:
                st.dataframe(display_df, use_container_width=True)

        with tab2:
            if error_df.empty:
                st.success("Ошибок парсинга не найдено.")
            else:
                st.warning("Эти строки НЕ попали в операции, потому что сумму не удалось корректно распознать (или она округлилась в 0).")
                st.dataframe(error_df, use_container_width=True)

        # ========== Export to Excel ==========
        output_buffer = BytesIO()

        # Prepare export: даты как строки, суммы Decimal→float для openpyxl
        export_for_excel = export_df.copy()
        if not export_for_excel.empty:
            export_for_excel["Дата ДДС"] = pd.to_datetime(export_for_excel["Дата ДДС"], errors="coerce").dt.strftime("%d.%m.%Y").fillna("")
            export_for_excel["Дата P&L"] = pd.to_datetime(export_for_excel["Дата P&L"], errors="coerce").dt.strftime("%d.%m.%Y").fillna("")
            # Decimal → float (openpyxl не пишет Decimal), нули оставляем как 0
            export_for_excel["Приход"] = export_for_excel["Приход"].apply(
                lambda x: float(x) if isinstance(x, Decimal) else (x if x else 0)
            )
            export_for_excel["Расход"] = export_for_excel["Расход"].apply(
                lambda x: float(x) if isinstance(x, Decimal) else (x if x else 0)
            )

        # Filename from DDS dates
        if export_df.empty or export_df["Дата ДДС"].dropna().empty:
            file_name = "Операции_без_дат.xlsx"
        else:
            dds_dates = export_df["Дата ДДС"].dropna()
            first_date = dds_dates.min().strftime("%d.%m.%Y")
            last_date = dds_dates.max().strftime("%d.%m.%Y")
            file_name = f"Операции_{first_date}_{last_date}.xlsx"

        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            # Sheet 1: operations
            export_for_excel.to_excel(writer, index=False, sheet_name="Операции")

            # Sheet 2: errors
            if not error_df.empty:
                error_df.to_excel(writer, index=False, sheet_name="Ошибки")

            # --- Кастомный числовой формат: 0 отображается как пустая ячейка ---
            ws_ops = writer.sheets["Операции"]
            header_row = [cell.value for cell in ws_ops[1]]
            income_col_idx = (header_row.index("Приход") + 1) if "Приход" in header_row else None
            expense_col_idx = (header_row.index("Расход") + 1) if "Расход" in header_row else None
            # Формат: число с 2 знаками; 0 → визуально пусто; ячейка остается Numeric
            zero_hidden_fmt = '#,##0.00;-#,##0.00;""'
            for row_idx in range(2, ws_ops.max_row + 1):
                if income_col_idx:
                    cell = ws_ops.cell(row=row_idx, column=income_col_idx)
                    cell.number_format = zero_hidden_fmt
                if expense_col_idx:
                    cell = ws_ops.cell(row=row_idx, column=expense_col_idx)
                    cell.number_format = zero_hidden_fmt

            # Autosize columns (basic)
            for sheet_name in writer.sheets:
                ws = writer.sheets[sheet_name]
                for idx, col_cells in enumerate(ws.columns, 1):
                    header = ws.cell(row=1, column=idx).value
                    max_len = 0
                    for cell in col_cells:
                        v = cell.value
                        max_len = max(max_len, len(str(v)) if v is not None else 0)
                    col_letter = ws.cell(row=1, column=idx).column_letter
                    if header in {"Комментарий", "Лог", "Сумма_как_в_файле", "Сумма_нормализованная"}:
                        ws.column_dimensions[col_letter].width = 50
                    else:
                        ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

        output_buffer.seek(0)

        st.download_button(
            label="💾 Скачать результат (Excel)",
            data=output_buffer,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ошибка при обработке файла: {e}")
        st.exception(e)
