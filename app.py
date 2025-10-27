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
    for _, row in df.iterrows():
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

            income = value if value > 0 else 0
            expense = -value if value < 0 else 0
            if income == 0 and expense == 0:
                continue

            result_rows.append({
                "Дата ДДС": date_dds,
                "Дата P&L": date_pl,
                "Приход": round(income, 2),
                "Расход": round(expense, 2),
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

    # Для отображения
    display_df = result_df.copy()
    display_df['Дата ДДС'] = display_df['Дата ДДС'].dt.strftime('%d.%m.%Y').fillna('')
    display_df['Дата P&L'] = display_df['Дата P&L'].dt.strftime('%d.%m.%Y').fillna('')
    display_df['Приход'] = display_df['Приход'].replace(0, '')
    display_df['Расход'] = display_df['Расход'].replace(0, '')

    return display_df, result_df

# === Streamlit UI ===
st.set_page_config(page_title="Конвертер Финолог → Управленческий учет", layout="wide")
st.title("Конвертер Финолог → Управленческий учет")

uploaded_file = st.file_uploader("📂 Загрузите Excel-файл", type=["xlsx"])

if uploaded_file:
    try:
        display_df, export_df = process_excel(uploaded_file)
        if display_df is not None:
            st.dataframe(display_df, use_container_width=True)

            # Кнопка скачивания
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                export_df.to_excel(writer, index=False, sheet_name='Операции')
                worksheet = writer.sheets['Операции']
                for idx, col in enumerate(worksheet.columns, 1):
                    max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                    if worksheet.cell(1, idx).value == "Комментарий":
                        worksheet.column_dimensions[chr(64 + idx)].width = 50
                    else:
                        worksheet.column_dimensions[chr(64 + idx)].width = min(max_len + 2, 50)

            output.seek(0)
            st.download_button(
                label="💾 Скачать результат (Excel)",
                data=output,
                file_name="Операции_финолог_результат.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Ошибка при обработке файла: {e}")