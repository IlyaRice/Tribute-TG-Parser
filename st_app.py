# st_app.py
import streamlit as st
from parser import process_data, save_to_excel

st.title("Tribute TG bot financial data parser")

uploaded_file = st.file_uploader("Выберите JSON файл", type="json")

if uploaded_file is not None:
    df = process_data(uploaded_file)
    output_files, quarter_summaries = save_to_excel(df)

    # Only display summaries for non-empty quarters
    for quarter, summary in quarter_summaries.items():
        st.subheader(quarter)
        st.text(summary)

    # Only provide download buttons for quarters with data
    for file_name, file_data in output_files:
        st.download_button(
            label=f"Скачать {file_name}",
            data=file_data.getvalue(),
            file_name=file_name,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
