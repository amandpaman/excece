import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from fpdf import FPDF
import tempfile

# --- App Config ---
st.set_page_config(page_title="Excel Analyzer", layout="wide")
st.title("📊 Complex Excel Data Analyzer")

# --- Upload Excel ---
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])
if not uploaded_file:
    st.warning("Please upload an Excel file to continue.")
    st.stop()

excel = pd.ExcelFile(uploaded_file)
sheet_name = st.selectbox("Select sheet", excel.sheet_names)

# --- Multi-level Header Read ---
try:
    df = pd.read_excel(excel, sheet_name=sheet_name, header=[0, 1, 2])
except:
    df = pd.read_excel(excel, sheet_name=sheet_name, header=[0, 1])

st.subheader("🔎 Data Preview")
st.dataframe(df.head(), use_container_width=True)

# --- Section & Sub-Column Filters ---
st.sidebar.header("📂 Column Filters")
unique_sections = sorted(set(col[0] for col in df.columns))
selected_sections = st.sidebar.multiselect("Choose Sections", unique_sections, default=unique_sections)

filtered_cols = [col for col in df.columns if col[0] in selected_sections]
df_filtered = df[filtered_cols]

# Sub-column search
search = st.sidebar.text_input("🔍 Search sub-columns")
if search:
    filtered_cols = [col for col in filtered_cols if search.lower() in col[1].lower()]
    df_filtered = df[filtered_cols]

st.subheader("📄 Filtered Data")
st.dataframe(df_filtered, use_container_width=True)

# --- Excel Download ---
buffer = BytesIO()
df_filtered.to_excel(buffer, index=False)
st.download_button("📥 Download Filtered Excel", buffer.getvalue(), file_name="filtered_data.xlsx")

# --- Auto Chart ---
st.subheader("📈 Auto Chart Generator")
if not df_filtered.empty:
    chart_col = st.selectbox("Select column for chart", df_filtered.columns)
    chart_type = st.radio("Chart type", ["Bar", "Pie", "Line"])

    if chart_type == "Bar":
        chart_df = df_filtered[chart_col].value_counts().reset_index()
        fig = px.bar(chart_df, x='index', y=chart_col[2], title=f"Bar Chart of {chart_col[1]}")
    elif chart_type == "Pie":
        chart_df = df_filtered[chart_col].value_counts().reset_index()
        fig = px.pie(chart_df, names='index', values=chart_col[2], title=f"Pie Chart of {chart_col[1]}")
    else:
        try:
            fig = px.line(df_filtered, y=chart_col, title=f"Trend of {chart_col[1]}")
        except Exception:
            st.warning("Line chart failed. Try selecting a numeric column.")

    st.plotly_chart(fig, use_container_width=True)

# --- Pivot Table Generator ---
st.subheader("🧮 Pivot Table Generator")
pivot_col1 = st.selectbox("Row (Index)", df_filtered.columns, key="pivot_row")
pivot_col2 = st.selectbox("Column (Optional)", [None] + list(df_filtered.columns), key="pivot_col")
pivot_val = st.selectbox("Values", df_filtered.columns, key="pivot_val")
agg_func = st.selectbox("Aggregation", ["sum", "mean", "count"], key="agg_func")

try:
    pivot_df = pd.pivot_table(
        df_filtered,
        index=[pivot_col1],
        columns=[pivot_col2] if pivot_col2 else None,
        values=pivot_val,
        aggfunc=agg_func
    )
    st.dataframe(pivot_df, use_container_width=True)
except Exception as e:
    st.error(f"Pivot Error: {e}")

# --- PDF Summary Generator ---
st.subheader("📄 Generate Summary Report (PDF)")

if st.button("📤 Export Summary as PDF"):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmpfile:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt="Excel Data Summary Report", ln=True, align='C')

        pdf.ln(10)
        pdf.cell(200, 10, txt=f"Sheet: {sheet_name}", ln=True)

        # Add column stats
        pdf.ln(5)
        for col in df_filtered.columns[:10]:  # Limit to 10 columns for summary
            try:
                if pd.api.types.is_numeric_dtype(df_filtered[col]):
                    stats = df_filtered[col].describe()
                    pdf.multi_cell(0, 10, f"{col}:\nMean={stats['mean']:.2f}, Min={stats['min']}, Max={stats['max']}")
                    pdf.ln(1)
            except Exception:
                continue

        pdf.output(tmpfile.name)
        with open(tmpfile.name, "rb") as f:
            st.download_button("📥 Download PDF Summary", f.read(), file_name="summary_report.pdf")

# --- Footer ---
st.caption("Made with ❤️ using Streamlit | Advanced Analyzer Dashboard")
