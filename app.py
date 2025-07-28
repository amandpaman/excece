import streamlit as st
import pandas as pd
from io import BytesIO
from st_aggrid import GridOptionsBuilder, AgGrid, JsCode
import plotly.express as px
from fpdf import FPDF

# Title
st.set_page_config(page_title="Complex Excel Dashboard", layout="wide")
st.title("📊 Complex Excel Data Viewer & Dashboard")

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    # Get all sheet names
    xl = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("Select Sheet", xl.sheet_names)

    # Read selected sheet
    df = xl.parse(sheet_name, header=[0, 1] if st.checkbox("Enable Sub-columns (MultiIndex)", value=True) else 0)

    st.markdown("### Data Preview")
    st.dataframe(df, use_container_width=True)

    # Optional column filters
    st.sidebar.header("🔍 Filter Data")
    column_to_filter = st.sidebar.selectbox("Select Column to Filter", df.columns, index=0)
    if pd.api.types.is_numeric_dtype(df[column_to_filter]):
        min_val, max_val = st.sidebar.slider("Value Range", float(df[column_to_filter].min()), float(df[column_to_filter].max()),
                                             (float(df[column_to_filter].min()), float(df[column_to_filter].max())))
        filtered_df = df[(df[column_to_filter] >= min_val) & (df[column_to_filter] <= max_val)]
    else:
        filter_val = st.sidebar.text_input("Enter Filter Keyword")
        filtered_df = df[df[column_to_filter].astype(str).str.contains(filter_val, case=False, na=False)]

    st.markdown("### 🔎 Filtered Data")
    st.dataframe(filtered_df, use_container_width=True)

    # Pivot Table Section
    st.markdown("---")
    st.subheader("📌 Create Pivot Table")
    with st.expander("Generate Pivot Table"):
        pivot_index = st.multiselect("Rows", options=filtered_df.columns)
        pivot_columns = st.multiselect("Columns", options=filtered_df.columns)
        pivot_values = st.multiselect("Values", options=filtered_df.columns)
        aggfunc = st.selectbox("Aggregation Function", ["sum", "mean", "count", "min", "max"], index=0)

        if st.button("Generate Pivot Table"):
            try:
                pivot_table = pd.pivot_table(
                    filtered_df,
                    index=pivot_index,
                    columns=pivot_columns,
                    values=pivot_values,
                    aggfunc=aggfunc
                )
                st.dataframe(pivot_table)
            except Exception as e:
                st.error(f"❌ Error generating pivot table: {e}")

    # Export Section
    st.markdown("---")
    st.subheader("⬇️ Export Data")

    # Excel Export
    excel_buf = BytesIO()
    export_df = filtered_df.copy()

    if isinstance(export_df.columns, pd.MultiIndex):
        export_df.columns = [' - '.join([str(i) for i in col if str(i).strip()]) for col in export_df.columns]

    export_df.to_excel(excel_buf, index=False)
    excel_buf.seek(0)

    st.download_button(
        label="📤 Download Filtered Data as Excel",
        data=excel_buf,
        file_name="filtered_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # PDF Export
    if st.button("📄 Generate PDF Report"):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=10)

        pdf.cell(200, 10, txt="Filtered Data Summary Report", ln=True, align='C')
        pdf.ln(10)

        # Limit rows in PDF to 30
        for i, row in export_df.head(30).iterrows():
            line = ", ".join([f"{col}: {row[col]}" for col in export_df.columns])
            pdf.multi_cell(0, 8, line)

        pdf_output = BytesIO()
        pdf.output(pdf_output)
        pdf_output.seek(0)

        st.download_button(
            label="📥 Download PDF Report",
            data=pdf_output,
            file_name="summary_report.pdf",
            mime="application/pdf"
        )

    # Visualization Section
    st.markdown("---")
    st.subheader("📊 Data Visualization")
    with st.expander("📈 Create Chart"):
        numeric_cols = [col for col in export_df.columns if pd.api.types.is_numeric_dtype(export_df[col])]
        all_cols = export_df.columns.tolist()

        x_axis = st.selectbox("X-axis", options=all_cols)
        y_axis = st.selectbox("Y-axis", options=numeric_cols)
        chart_type = st.selectbox("Chart Type", ["Bar", "Line", "Scatter"])

        if chart_type == "Bar":
            fig = px.bar(export_df, x=x_axis, y=y_axis)
        elif chart_type == "Line":
            fig = px.line(export_df, x=x_axis, y=y_axis)
        else:
            fig = px.scatter(export_df, x=x_axis, y=y_axis)

        st.plotly_chart(fig, use_container_width=True)
