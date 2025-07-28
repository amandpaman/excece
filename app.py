import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import graphviz  # For flowcharts (replaces mermaid)

# 1. Data Parser
def parse_excel(uploaded_file):
    """Parse Excel file and identify sections"""
    try:
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
    except Exception as e:
        st.error(f"Error parsing file: {str(e)}")
        return None, None

# 2. Generate Graphviz Flowchart
def generate_flowchart(sections):
    """Create interactive flowchart using Graphviz"""
    dot = graphviz.Digraph()
    for section, cols in sections.items():
        dot.node(section)
        for col in cols[:3]:  # Show first 3 columns per section
            dot.node(col)
            dot.edge(section, col)
    return dot

# 3. Streamlit App
def main():
    st.set_page_config(layout="wide", page_title="📊 Telecom Data Visualizer")
    st.title("📊 Telecom Data Visualizer")
    
    # --- File Upload ---
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    
    if uploaded_file:
        df, sections = parse_excel(uploaded_file)
        
        if df is not None:
            # --- Tabs ---
            tab1, tab2, tab3 = st.tabs(["Data Structure", "Charts", "Raw Data"])
            
            with tab1:
                st.subheader("Data Hierarchy Flowchart")
                dot = generate_flowchart(sections)
                st.graphviz_chart(dot)
                
                with st.expander("Section Details"):
                    for section, cols in sections.items():
                        st.write(f"**{section}**")
                        st.code(", ".join(cols), language="text")
            
            with tab2:
                st.subheader("Financial Analysis")
                
                # CAPEX Bar Chart
                capex_cols = [col for col in df.columns if "CAPEX" in col and pd.api.types.is_numeric_dtype(df[col])]
                if capex_cols:
                    capex_data = df[capex_cols].sum().reset_index()
                    capex_data.columns = ["Category", "Total"]
                    fig = px.bar(capex_data, x="Category", y="Total", title="CAPEX Breakdown")
                    st.plotly_chart(fig, use_container_width=True)
                
                # Status Pie Chart
                if "RAG STATUS" in df.columns:
                    fig2 = px.pie(df, names="RAG STATUS", title="Project Status")
                    st.plotly_chart(fig2, use_container_width=True)
            
            with tab3:
                st.subheader("Interactive Data Table")
                
                # Filters
                col1, col2 = st.columns(2)
                with col1:
                    circle_filter = st.selectbox(
                        "Filter by Circle:", 
                        options=["All"] + df["A END CIRCLE"].unique().tolist()
                    )
                with col2:
                    status_filter = st.selectbox(
                        "Filter by Status:", 
                        options=["All"] + df["RAG STATUS"].unique().tolist()
                    )
                
                # Apply filters
                filtered_df = df.copy()
                if circle_filter != "All":
                    filtered_df = filtered_df[filtered_df["A END CIRCLE"] == circle_filter]
                if status_filter != "All":
                    filtered_df = filtered_df[filtered_df["RAG STATUS"] == status_filter]
                
                st.dataframe(filtered_df, use_container_width=True, height=600)
                
                # Download button
                csv = filtered_df.to_csv(index=False).encode()
                st.download_button(
                    label="Download Filtered Data (CSV)",
                    data=csv,
                    file_name="filtered_telecom_data.csv",
                    mime="text/csv"
                )
        else:
            st.warning("Failed to process the uploaded file. Please check the format.")
    else:
        st.info("👆 Please upload an Excel file to begin analysis")

if __name__ == "__main__":
    main()
