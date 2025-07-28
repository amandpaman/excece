import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import openai  # Replace with Gemini API when available

# --- App Configuration ---
st.set_page_config(page_title="Excel Analyzer", layout="wide")
st.title("📊 Complex Excel Data Analyzer with Gemini AI")

# --- File Upload ---
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])
if not uploaded_file:
    st.warning("Please upload an Excel file to continue.")
    st.stop()

# --- Sheet Selector ---
excel = pd.ExcelFile(uploaded_file)
sheet_name = st.selectbox("Select sheet", excel.sheet_names)

# --- Read Data with Multi-Index Columns ---
try:
    df = pd.read_excel(excel, sheet_name=sheet_name, header=[0, 1, 2])
except:
    df = pd.read_excel(excel, sheet_name=sheet_name, header=[0, 1])

st.subheader("🔎 Preview of Uploaded Data")
st.dataframe(df.head(), use_container_width=True)

# --- Column & Subcolumn Filtering ---
st.sidebar.header("📂 Column & Sub-Column Filters")

unique_sections = sorted(set([col[0] for col in df.columns]))
selected_sections = st.sidebar.multiselect("Choose Sections", unique_sections, default=unique_sections)

filtered_cols = [col for col in df.columns if col[0] in selected_sections]
df_filtered = df[filtered_cols]

# Sub-column Search
all_subcols = sorted(set(col[1] for col in filtered_cols))
search = st.sidebar.text_input("🔍 Search sub-columns")
if search:
    filtered_cols = [col for col in filtered_cols if search.lower() in col[1].lower()]
    df_filtered = df[filtered_cols]

st.subheader("📄 Filtered Data")
st.dataframe(df_filtered, use_container_width=True)

# --- Download Filtered Data ---
buffer = BytesIO()
df_filtered.to_excel(buffer, index=False)
st.download_button("📥 Download Filtered Data", buffer.getvalue(), file_name="filtered_data.xlsx")

# --- Auto Chart Generator ---
st.subheader("📈 Auto Chart Generator")

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

# --- Gemini API (or Placeholder) ---
st.subheader("💡 Gemini Data Assistant")
query = st.text_input("Ask a question about this Excel sheet (e.g., 'What is the average delivery time?')")

if query:
    with st.spinner("Analyzing with Gemini..."):
        # Placeholder - replace with actual Gemini API call when ready
        answer = f"🤖 [Mock Response] You asked: '{query}' — Actual answer will appear here using Gemini API."
        st.success(answer)

# --- Footer ---
st.caption("Built with ❤️ using Streamlit | Gemini API integration coming soon.")
