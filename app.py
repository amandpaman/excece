import streamlit as st
import pandas as pd
import plotly.express as px
import mermaid as mmd  # For flowcharts (custom component)
from io import BytesIO

# 1. Data Parser
def parse_excel(uploaded_file):
    df = pd.read_excel(uploaded_file, header=[2])  # Skip 2 header rows
    sections = {
        "Order Info": df.columns[:7].tolist(),
        "Commercial Info": df.columns[7:17].tolist(),
        "Last Mile Info": df.columns[17:32].tolist(),
        "Posting Info": df.columns[32:40].tolist(),
        "Order Aging": df.columns[40:50].tolist(),
        "Delivery Info": df.columns[50:].tolist()
    }
    return df, sections

# 2. Generate Mermaid Flowchart
def generate_flowchart(sections):
    chart = """
    graph TD
    """
    for section, cols in sections.items():
        chart += f"    {section.replace(' ', '_')}[{section}]\n"
        for col in cols[:3]:  # Show first 3 cols per section
            chart += f"    {section.replace(' ', '_')} --> {col.replace(' ', '_')}\n"
    return chart

# 3. Streamlit App
st.set_page_config(layout="wide")
st.title("📊 Telecom Data Visualizer")

# --- Upload & Parse ---
uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])
if uploaded_file:
    df, sections = parse_excel(uploaded_file)
    
    # --- Tabs ---
    tab1, tab2, tab3 = st.tabs(["Hierarchy", "Charts", "Data Table"])
    
    with tab1:
        st.subheader("Data Structure Flowchart")
        flowchart = generate_flowchart(sections)
        st.graphviz_chart(flowchart)
        
        # Alternative: Use Mermaid (requires custom component)
        # mmd.mermaid(flowchart, key="flowchart")
    
    with tab2:
        st.subheader("CAPEX Analysis")
        capex_cols = [col for col in df.columns if "CAPEX" in col]
        capex_data = df[capex_cols].sum().reset_index()
        capex_data.columns = ["Category", "Total"]
        
        fig = px.bar(capex_data, x="Category", y="Total", title="CAPEX Breakdown")
        st.plotly_chart(fig, use_container_width=True)
        
        st.subheader("Aging Status")
        fig2 = px.pie(df, names="RAG STATUS", title="Project Status")
        st.plotly_chart(fig2, use_container_width=True)
    
    with tab3:
        st.subheader("Interactive Data Table")
        st.dataframe(df, use_container_width=True)
        
        # Filter
        selected_circle = st.selectbox("Filter by Circle:", df["A END CIRCLE"].unique())
        filtered_df = df[df["A END CIRCLE"] == selected_circle]
        st.write(f"Showing {len(filtered_df)} records for {selected_circle}")

    # --- Export ---
    st.sidebar.download_button(
        label="📥 Download Processed Data",
        data=df.to_csv(index=False).encode(),
        file_name="processed_telecom_data.csv"
    )
else:
    st.info("Upload an Excel file to begin analysis")

# 4. Run with: streamlit run app.py
