import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide", page_title="📊 Excel Dashboard Viewer")

st.markdown("## 🧠 Complex Excel Data Viewer")
st.markdown("Upload an Excel file with multiple header rows (e.g., Main Columns & Sub-Columns).")

# --- File Upload ---
uploaded_file = st.file_uploader("📂 Upload Excel File", type=["xlsx"])

if uploaded_file:
    try:
        # --- Read Excel with multi-row header ---
        df = pd.read_excel(uploaded_file, header=[0, 1])  # Read two header rows

        # Flatten the headers to display
        st.success("✅ File uploaded and read successfully.")
        st.markdown("### 📌 Column Structure Preview")
        st.dataframe(df.head(), use_container_width=True)

        # Extract main and sub-columns
        main_sections = sorted(set([col[0] for col in df.columns if isinstance(col, tuple)]))

        selected_section = st.selectbox("🔍 Select Main Section", main_sections)

        # Get sub-columns for that section
        sub_columns = [col[1] for col in df.columns if col[0] == selected_section]
        selected_subcol = st.selectbox("📁 Select Sub-Column", sub_columns)

        # Final column tuple
        selected_col = (selected_section, selected_subcol)

        # Display table for selected column
        st.markdown(f"### 🧾 Data for `{selected_section}` → `{selected_subcol}`")
        st.dataframe(df[[selected_col]].dropna(), use_container_width=True)

        # Try visualizing if numeric or categorical
        chart_type = st.radio("📊 Choose Chart Type", ["Bar Chart", "Pie Chart", "Raw Table"])

        chart_data = df[[selected_col]].dropna()
        chart_data.columns = ["value"]  # Rename for easier charting

        if chart_type == "Pie Chart":
            fig = px.pie(chart_data, names="value", title=f"Pie Chart of {selected_subcol}")
            st.plotly_chart(fig, use_container_width=True)

        elif chart_type == "Bar Chart":
            top_counts = chart_data["value"].value_counts().head(20).reset_index()
            top_counts.columns = [selected_subcol, "Count"]
            fig = px.bar(top_counts, x=selected_subcol, y="Count", title=f"Top {selected_subcol} values")
            st.plotly_chart(fig, use_container_width=True)

        else:
            st.dataframe(chart_data)

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")

else:
    st.info("⬆️ Upload your Excel file to begin.")
