import streamlit as st
import pandas as pd
import io
from datetime import datetime
import zipfile
from invoice_processor import InvoiceProcessor
from config.brand_configs import BRAND_CONFIGS

# Page configuration
st.set_page_config(
    page_title="Invoice Generator Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# iOS 26 "Liquid Glass" inspired custom CSS
st.markdown("""
<style>
    /* System font stack like San Francisco */
    html, body, [class*="css"]  {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Oxygen, Ubuntu, Cantarell, "Open Sans", "Helvetica Neue", sans-serif;
    }

    /* Subtle page background */
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #e4edf5 100%);
    }

    /* Main header - centered, larger, elegant */
    .main-header {
        font-size: 2.8rem;
        font-weight: 700;
        text-align: center;
        margin: 2rem 0 3rem 0;
        background: linear-gradient(to right, #1e3c72, #2a5298);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }

    /* Frosted glass sidebar */
    section[data-testid="stSidebar"] {
        background: rgba(255, 255, 255, 0.65);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border-right: 1px solid rgba(255, 255, 255, 0.4);
        border-radius: 20px;
        margin: 10px;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.08);
    }

    /* Glass cards for main content sections */
    .glass-card {
        background: rgba(255, 255, 255, 0.55);
        backdrop-filter: blur(20px);
        -webkit-backdrop-filter: blur(20px);
        border-radius: 20px;
        border: 1px solid rgba(255, 255, 255, 0.6);
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
    }

    /* Success box - iOS style */
    .success-box {
        background: rgba(212, 237, 218, 0.7);
        backdrop-filter: blur(10px);
        border: 1px solid rgba(195, 230, 203, 0.8);
        border-radius: 16px;
        padding: 1.5rem;
        margin: 1.5rem 0;
    }

    /* Metrics and stats */
    div[data-testid="metric-container"] {
        background: rgba(255, 255, 255, 0.7);
        backdrop-filter: blur(12px);
        border-radius: 16px;
        padding: 1rem;
        border: 1px solid rgba(255, 255, 255, 0.5);
    }

    /* Buttons - iOS primary style */
    .stButton > button {
        border-radius: 16px !important;
        height: 3em;
        font-weight: 600;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
    }
    .stButton > button[kind="primary"] {
        background: linear-gradient(to bottom, #007aff, #0056cc);
        border: none;
        color: white;
    }

    /* Dataframes and expanders */
    .stDataFrame, div[data-testid="stExpander"] {
        border-radius: 16px;
        overflow: hidden;
    }

    /* Footer */
    .app-footer {
        text-align: center;
        padding: 2rem 0;
        margin-top: 4rem;
        color: #666;
        font-size: 0.95rem;
    }

    /* Hide Streamlit menu and footer for cleaner look */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

def main():
    # Header
    st.markdown('<h1 class="main-header">Invoice Automation System</h1>', unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.image("assets/rv_solutions_logo.png", use_container_width=True)

        st.markdown("<h3 style='text-align: center; margin-top: 1rem;'>Settings</h3>", unsafe_allow_html=True)

        selected_brand = st.selectbox("Select Brand", list(BRAND_CONFIGS.keys()))

        st.markdown("### Instructions")
        st.markdown("""
        1. Select your brand  
        2. Upload the raw data Excel file  
        3. Click Generate Invoices  
        4. Download all invoices as ZIP
        """)

        st.markdown("### Required Columns")
        config = BRAND_CONFIGS[selected_brand]
        for col in config['required_columns']:
            st.markdown(f"• {col}")

    # Main content with glass cards
    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
        st.subheader("📂 Upload Raw Data")
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload the raw billing data Excel file"
        )

        if uploaded_file:
            try:
                df = pd.read_excel(uploaded_file)
                st.success(f"File loaded successfully! ({len(df)} rows, {len(df.columns)} columns)")

                with st.expander("📄 Data Preview", expanded=False):
                    st.dataframe(df.head(), use_container_width=True)

                    missing_cols = [col for col in config['required_columns'] if col not in df.columns]
                    if missing_cols:
                        st.warning(f"⚠️ Missing columns: {', '.join(missing_cols)}")
                    else:
                        st.success("✓ All required columns present")
            except Exception as e:
                st.error(f"Error reading file: {str(e)}")
        st.markdown("</div>", unsafe_allow_html=True)

    with col2:
        st.markdown("<div class='glass-card'>", unsafe_allow_html=True)
        st.subheader("📊 Statistics")

        if uploaded_file:
            try:
                df = pd.read_excel(uploaded_file)
                asc_column = config['asc_column']

                if asc_column in df.columns:
                    total_ascs = df[asc_column].nunique()
                    total_records = len(df)
                    total_earning = df['Earning'].sum() if 'Earning' in df.columns else 0

                    st.metric("Total ASCs", total_ascs)
                    st.metric("Total Records", total_records)
                    st.metric("Total Earning", f"₹{total_earning:,.2f}")
                else:
                    st.warning(f"ASC column '{asc_column}' not found")
            except:
                st.info("Upload file to see statistics")
        else:
            st.info("Upload file to see statistics")
        st.markdown("</div>", unsafe_allow_html=True)

    # Generate section
    st.markdown("---")

    if uploaded_file:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            generate_btn = st.button(
                "🚀 Generate All Invoices",
                type="primary",
                use_container_width=True
            )

        if generate_btn:
            with st.spinner("Processing invoices..."):
                try:
                    processor = InvoiceProcessor(selected_brand, config)
                    uploaded_file.seek(0)
                    results = processor.process_invoices(uploaded_file)

                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    total_ascs = len(results)
                    for processed, (asc_name, _) in enumerate(results.items(), 1):
                        progress_bar.progress(processed / total_ascs)
                        status_text.text(f"Processing {asc_name}... ({processed}/{total_ascs})")

                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for asc_name, data in results.items():
                            safe_name = "".join(c for c in asc_name if c.isalnum() or c in (' ', '-', '_')).strip()
                            invoice_filename = f"Invoice_{safe_name}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                            zip_file.writestr(invoice_filename, data['invoice'])
                            raw_filename = f"RawData_{safe_name}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                            zip_file.writestr(raw_filename, data['raw_data'])

                    zip_buffer.seek(0)

                    st.markdown("""
                    <div class="success-box">
                        <h3>✅ Invoice Generation Complete!</h3>
                        <p>Successfully generated invoices for all ASCs.</p>
                    </div>
                    """, unsafe_allow_html=True)

                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    zip_filename = f"{selected_brand}_Invoices_{timestamp}.zip"

                    st.download_button(
                        label="📥 Download All Invoices (ZIP)",
                        data=zip_buffer,
                        file_name=zip_filename,
                        mime="application/zip",
                        use_container_width=True
                    )

                    st.subheader("📋 Generation Summary")
                    summary_data = []
                    total_records = 0
                    total_earning = 0

                    for asc_name, data in results.items():
                        summary_data.append({
                            'ASC Name': asc_name,
                            'Records': data['records'],
                            'Total Amount': f"₹{data['total_amount']:,.2f}"
                        })
                        total_records += data['records']
                        total_earning += data['total_amount']

                    summary_df = pd.DataFrame(summary_data)
                    st.dataframe(summary_df, use_container_width=True)

                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total ASCs", len(results))
                    with col2:
                        st.metric("Total Records", total_records)
                    with col3:
                        st.metric("Total Amount", f"₹{total_earning:,.2f}")

                except Exception as e:
                    st.error(f"Error generating invoices: {str(e)}")
    else:
        st.info("Please upload an Excel file to begin")

    # Footer
    st.markdown("""
        <div class="app-footer">
            <strong>Crafted with precision by Harpinder Singh</strong><br>
            <span>© RV Solutions</span>
        </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()