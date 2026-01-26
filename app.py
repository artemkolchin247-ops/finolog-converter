import pandas as pd
import streamlit as st
import re
from io import BytesIO
from datetime import datetime

# =========================
# Helpers
# =========================

def norm_text(x) -> str:
    """Normalize header/cell text for comparisons."""
    if x is None:
        return ""
    s = str(x)
    # normalize spaces (regular + NBSP + narrow NBSP)
    s = s.replace("\u00A0", " ").replace("\u202F", " ")
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
      value: float | None
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

    # already numeric
    if isinstance(val, (int, float)) and not pd.isna(val):
        f = float(val)
        debug["raw_str"] = str(val)
        debug["normalized"] = str(f)
        debug["rule"] = "OK:already_numeric"
        return f, None, debug

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
    s = s.replace("−", "-").replace("–", "-").replace("—", "-")

    # parentheses negative
    is_paren_negative = s.startswith("(") and s.endswith(")")
    if is_paren_negative:
        s = s[1:-1].strip()
        debug["rule"] = "RULE:paren_negative"

    # remove spaces (including NBSP variants)
    s = s.replace("\u00A0", " ").replace("\u202F", " ")
    s = s.replace(" ", "")

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

    # final cast
    try:
        num = float(s)
        if is_paren_negative:
            num = -abs(num)
        debug["rule"] = ("OK:parsed|" + debug["rule"]) if debug["rule"] else "OK:parsed"
        return num, None, debug
    except Exception as e:
        debug["rule"] = "FAIL:float_cast"
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
amount_cols = build_amount_columns(df)

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

    amount_cols = build_amount_columns(df, df.columns)

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
            
            income = round(value, 2) if value > 0 else 0.0
            expense = round(-value, 2) if value < 0 else 0.0
            
            # если число есть, но после округления стало 0 — фиксируем как отдельную причину
            if income == 0.0 and expense == 0.0 and abs(value) > 0:
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

            income = round(value, 2) if value > 0 else 0.0
            expense = round(-value, 2) if value < 0 else 0.0

            # Keep zeros only if the original value was non-zero but rounded to zero
            # (rare but avoids losing micro-values). If you want to drop strict zeros, revert this logic.
            if income == 0.0 and expense == 0.0 and abs(value) > 0:
                # store as warning record in errors instead of dropping
                error_rows.append({
                    "Источник_строка": int(src_idx) + 1,
                    "Источник_колонка": norm_text(col),
                    "Исходное значение": raw_val,
                    "Ошибка": f"ROUNDED_TO_ZERO: {value}",
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
    
        # гарантируем числа
        result_df["Приход"] = pd.to_numeric(result_df["Приход"], errors="coerce").fillna(0.0)
        result_df["Расход"] = pd.to_numeric(result_df["Расход"], errors="coerce").fillna(0.0)
    
        # тип строки, чтобы приход и расход не склеились
        result_df["Тип"] = result_df.apply(
            lambda r: "Приход" if r["Приход"] != 0 else ("Расход" if r["Расход"] != 0 else "Ноль"),
            axis=1
        )
    
        # убираем строки без суммы
        result_df = result_df[result_df["Тип"] != "Ноль"].copy()
    
        # нормализуем комментарий
        result_df["Комментарий"] = result_df["Комментарий"].fillna("").astype(str).str.strip()
    
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
    
        # ключ: всё кроме Комментария и сумм
        # Источник_* НЕ включаем
        group_keys = ["Тип", "Дата ДДС", "Дата P&L", "Статья операции", "Касса / Счет"]
    
        result_df = (
            result_df
            .groupby(group_keys, dropna=False, as_index=False)
            .agg({
                "Приход": "sum",
                "Расход": "sum",
                "Комментарий": join_comments
            })
        )
    
        result_df["Приход"] = result_df["Приход"].round(2)
        result_df["Расход"] = result_df["Расход"].round(2)
    
        # на всякий случай после суммирования убрать нули
        result_df = result_df[~((result_df["Приход"] == 0) & (result_df["Расход"] == 0))].copy()
    
        # сортировка после схлопывания
        result_df = result_df.sort_values(by=["Дата ДДС", "Дата P&L", "Тип"]).reset_index(drop=True)
    
        # убрать технический столбец
        result_df = result_df.drop(columns=["Тип"], errors="ignore")
    
    # Display df (human-friendly)
    # --- фиксируем порядок колонок ---
    desired_order = [
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

    display_df = result_df.copy()
    if not display_df.empty:
        display_df["Дата ДДС"] = display_df["Дата ДДС"].dt.strftime("%d.%m.%Y").fillna("")
        display_df["Дата P&L"] = display_df["Дата P&L"].dt.strftime("%d.%m.%Y").fillna("")
        display_df["Приход"] = display_df["Приход"].apply(lambda x: "" if x == 0 else x)
        display_df["Расход"] = display_df["Расход"].apply(lambda x: "" if x == 0 else x)
    
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

        # Prepare export: dates as strings, zeros as blanks (as you wanted)
        export_for_excel = export_df.copy()
        if not export_for_excel.empty:
            export_for_excel["Дата ДДС"] = pd.to_datetime(export_for_excel["Дата ДДС"], errors="coerce").dt.strftime("%d.%m.%Y").fillna("")
            export_for_excel["Дата P&L"] = pd.to_datetime(export_for_excel["Дата P&L"], errors="coerce").dt.strftime("%d.%m.%Y").fillna("")
            export_for_excel["Приход"] = export_for_excel["Приход"].apply(lambda x: "" if x == 0 else x)
            export_for_excel["Расход"] = export_for_excel["Расход"].apply(lambda x: "" if x == 0 else x)

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
                    if header in {"Комментарий", "Ошибка", "Сырье"}:
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
