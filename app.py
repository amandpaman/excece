import streamlit as st
import pandas as pd
import plotly.express as px
import io
from pandas.io.formats.style import Styler
from openpyxl import load_workbook

st.set_page_config(layout="wide")
st.title("📊 Advanced Excel Dashboard Viewer")

# --- File Upload ---
file = st.file_uploader("Upload Excel File", type=[".xlsx"])

if file:
    xls = pd.ExcelFile(file)
    sheet_name = st.selectbox("Select Sheet", xls.sheet_names)

    # --- Read with MultiIndex headers (3 levels max assumed) ---
    df = pd.read_excel(file, sheet_name=sheet_name, header=[0, 1, 2])

    # --- Prepare column selectors ---
    all_columns = df.columns.to_list()
    main_headers = sorted(set(col[0] for col in all_columns if not pd.isna(col[0])))

    selected_main = st.multiselect("Select Column Sections", main_headers)

    selected_sub = []
    if selected_main:
        sub_headers = [col for col in all_columns if col[0] in selected_main]
        search_query = st.text_input("Search Sub-Columns (Optional)")
        filtered_sub_headers = [col for col in sub_headers if search_query.lower() in str(col).lower()]
        selected_sub = st.multiselect("Select Sub-Columns", filtered_sub_headers, default=filtered_sub_headers)

    if selected_sub:
        filtered_df = df[selected_sub]
    elif selected_main:
        filtered_df = df[[col for col in all_columns if col[0] in selected_main]]
    else:
        filtered_df = df.copy()

    st.subheader("Filtered Data Preview")
    st.dataframe(filtered_df, use_container_width=True)

    # --- Download as CSV or Excel ---
    st.markdown("### 📥 Download Options")
    with st.expander("Download Filtered Data"):
        col1, col2 = st.columns(2)

        with col1:
            csv = filtered_df.to_csv(index=False).encode('utf-8')
            st.download_button("Download CSV", csv, "filtered_data.csv", "text/csv")

        with col2:
            excel_buf = io.BytesIO()
            filtered_df.to_excel(excel_buf, index=False)
            st.download_button("Download Excel", excel_buf.getvalue(), "filtered_data.xlsx")

    # --- Pivot Table Generator ---
    st.subheader("📈 Pivot Table Generator")
    with st.expander("Create Pivot Table"):
        col_pivot1, col_pivot2 = st.columns(2)
        with col_pivot1:
            pivot_index = st.multiselect("Rows", options=filtered_df.columns.to_list())
        with col_pivot2:
            pivot_columns = st.multiselect("Columns", options=filtered_df.columns.to_list())

        pivot_values = st.multiselect("Values", options=filtered_df.columns.to_list())
        aggfunc = st.selectbox("Aggregation Function", ["sum", "mean", "count", "max", "min"])

        if pivot_values:
            try:
                pivot_table = pd.pivot_table(
                    filtered_df,
                    index=pivot_index,
                    columns=pivot_columns,
                    values=pivot_values,
                    aggfunc=aggfunc
                )
                st.dataframe(pivot_table, use_container_width=True)
            except Exception as e:
                st.warning(f"Pivot creation error: {e}")

    # --- Auto Chart Generator ---
    st.subheader("📊 Auto Chart Generator")
    with st.expander("Generate Charts"):
        chart_col = st.selectbox("Column for Chart (Numeric Only)", [col for col in filtered_df.columns if pd.api.types.is_numeric_dtype(filtered_df[col])])

        chart_type = st.radio("Chart Type", ["Bar", "Pie", "Line", "Area"])

        chart_data = filtered_df[chart_col].value_counts().reset_index()
        chart_data.columns = ['Category', 'Value']

        if chart_type == "Bar":
            fig = px.bar(chart_data, x='Category', y='Value', title=f"Bar Chart of {chart_col}")
        elif chart_type == "Pie":
            fig = px.pie(chart_data, names='Category', values='Value', title=f"Pie Chart of {chart_col}")
        elif chart_type == "Line":
            fig = px.line(chart_data, x='Category', y='Value', title=f"Line Chart of {chart_col}")
        else:
            fig = px.area(chart_data, x='Category', y='Value', title=f"Area Chart of {chart_col}")

        st.plotly_chart(fig, use_container_width=True)

    # --- PDF Export ---
    st.subheader("📄 Export Summary Report")
    with st.expander("Generate PDF Summary"):
        try:
            import pdfkit
            from jinja2 import Template

            summary_html = f"""
            <h2>Excel Summary Report</h2>
            <p><b>Selected Columns:</b> {', '.join([str(c) for c in selected_sub])}</p>
            <p><b>Shape:</b> {filtered_df.shape}</p>
            <h3>Head</h3>
            {filtered_df.head().to_html(index=False)}
            """
            pdf = pdfkit.from_string(summary_html, False)
            st.download_button("Download PDF Report", pdf, file_name="summary_report.pdf")
        except Exception as e:
            st.error("PDF generation failed. Ensure `pdfkit` and `wkhtmltopdf` are installed.")
