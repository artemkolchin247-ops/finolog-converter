import pandas as pd
import streamlit as st
import re
from io import BytesIO

# === Функция для получения скалярного значения ===
def get_scalar(value):
    if isinstance(value, pd.Series):
        value = value.iloc[0]
    return value

# === Основная логика обработки ===
def process_excel(uploaded_file):
    df_raw = pd.read_excel(uploaded_file, header=None)

    # Поиск строки с "Дата операции"
    header_idx = None
    for i in range(len(df_raw)):
        row_values = df_raw.iloc[i].astype(str).str.strip().str.lower().tolist()
        if 'дата операции' in row_values:
            header_idx = i
            break
    if header_idx is None:
        st.error("Не найдена строка с заголовками (нет 'Дата операции')")
        return None

    headers = df_raw.iloc[header_idx].astype(str).str.strip()
    df = df_raw.iloc[header_idx + 1:].copy()
    df.columns = headers

    cols_to_drop = ['№  п.п.', '№ п.п.', 'Заказ', 'Проект', 'Контрагент', 'Реквизит контрагента', 'Компания', 'ID', 'id']
    df = df.drop(columns=[col for col in cols_to_drop if col in df.columns], errors='ignore')

    exclude_cols = {'Дата операции', 'Описание', 'Статья', 'Дата начисления', '', 'nan'}
    amount_cols = [col for col in df.columns if str(col).strip() not in exclude_cols]

    result_rows = []
    for src_idx, row in df.iterrows():
        date_dds = get_scalar(row.get('Дата операции', pd.NaT))
        date_pl = get_scalar(row.get('Дата начисления', pd.NaT))
        description = str(get_scalar(row.get('Описание', ''))) if pd.notna(row.get('Описание', None)) else ''
        article = str(get_scalar(row.get('Статья', ''))) if pd.notna(row.get('Статья', None)) else ''

        for col in amount_cols:
            val = get_scalar(row[col])
            if pd.isna(val) or str(val).strip() in ['', '-', 'None', 'nan']:
                continue
            try:
                cleaned = re.sub(r'[₽$€\s]', '', str(val))
                cleaned = cleaned.replace(',', '.')
                value = float(cleaned)
            except:
                continue

            income = round(value, 2) if value > 0 else 0
            expense = round(-value, 2) if value < 0 else 0
            if income == 0 and expense == 0:
                continue

            result_rows.append({
                "Источник_строка": int(src_idx) + 1,
                "Источник_колонка": str(col).strip(),

                "Дата ДДС": date_dds,
                "Дата P&L": date_pl,
                "Приход": income,
                "Расход": expense,
                "Статья операции": article,
                "Касса / Счет": str(col).strip(),
                "Комментарий": description
            })

    if not result_rows:
        st.warning("Нет данных для отображения")
        return None

    result_df = pd.DataFrame(result_rows)
    result_df['Дата ДДС'] = pd.to_datetime(result_df['Дата ДДС'], errors='coerce')
    result_df['Дата P&L'] = pd.to_datetime(result_df['Дата P&L'], errors='coerce')
    result_df = result_df.sort_values(by='Дата ДДС').reset_index(drop=True)

    # --- Для отображения в Streamlit: заменяем 0 на пусто ---
    display_df = result_df.copy()
    display_df['Дата ДДС'] = display_df['Дата ДДС'].dt.strftime('%d.%m.%Y').fillna('')
    display_df['Дата P&L'] = display_df['Дата P&L'].dt.strftime('%d.%m.%Y').fillna('')
    display_df['Приход'] = display_df['Приход'].apply(lambda x: '' if x == 0 else x)
    display_df['Расход'] = display_df['Расход'].apply(lambda x: '' if x == 0 else x)

    return display_df, result_df

# === Streamlit UI ===
st.set_page_config(page_title="Конвертер Финолог → Управленческий учет", layout="wide")
st.title("Конвертер Финолог → Управленческий учет")

uploaded_file = st.file_uploader("📂 Загрузите Excel-файл", type=["xlsx"])

if uploaded_file:
    try:
        output = process_excel(uploaded_file)
        if output is None:
            st.stop()
        display_df, export_df = output

        st.dataframe(display_df, use_container_width=True)

        # --- Подготовка к экспорту в Excel с нужным форматом дат и пустыми ячейками ---
        export_for_excel = export_df.copy()

        # Преобразуем даты в СТРОКИ формата "01.10.2025" (важно — не даты Excel!)
        export_for_excel['Дата ДДС'] = pd.to_datetime(export_for_excel['Дата ДДС'], errors='coerce').dt.strftime('%d.%m.%Y').fillna('')
        export_for_excel['Дата P&L'] = pd.to_datetime(export_for_excel['Дата P&L'], errors='coerce').dt.strftime('%d.%m.%Y').fillna('')

        # Заменяем 0 на пустую строку в Приход/Расход
        export_for_excel['Приход'] = export_for_excel['Приход'].apply(lambda x: '' if x == 0 else x)
        export_for_excel['Расход'] = export_for_excel['Расход'].apply(lambda x: '' if x == 0 else x)

        # --- Формируем имя файла ---
        dds_dates = export_df['Дата ДДС'].dropna()
        if dds_dates.empty:
            file_name = "Операции_без_дат.xlsx"
        else:
            first_date = dds_dates.min().strftime('%d.%m.%Y')
            last_date = dds_dates.max().strftime('%d.%m.%Y')
            file_name = f"Операции_{first_date}_{last_date}.xlsx"

        # --- Сохраняем в Excel ---
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            export_for_excel.to_excel(writer, index=False, sheet_name='Операции')
            worksheet = writer.sheets['Операции']

            # Настройка ширины столбцов
            for idx, col in enumerate(worksheet.columns, 1):
                max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
                col_letter = worksheet.cell(row=1, column=idx).column_letter
                if worksheet.cell(row=1, column=idx).value == "Комментарий":
                    worksheet.column_dimensions[col_letter].width = 50
                else:
                    worksheet.column_dimensions[col_letter].width = min(max_len + 2, 50)

        output_buffer.seek(0)

        st.download_button(
            label="💾 Скачать результат (Excel)",
            data=output_buffer,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ошибка при обработке файла: {e}")
        st.exception(e)  # для отладки (можно убрать в продакшене)
