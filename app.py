# app.py
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
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=SF+Pro+Display:wght@300;400;500;600;700&display=swap');
    
    * {
        font-family: -apple-system, BlinkMacSystemFont, 'SF Pro Display', 'Segoe UI', sans-serif;
    }
    
    /* Main container styling */
    .main {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 2rem;
    }
    
    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background: rgba(255, 255, 255, 0.7);
        backdrop-filter: blur(20px) saturate(180%);
        -webkit-backdrop-filter: blur(20px) saturate(180%);
        border-right: 1px solid rgba(255, 255, 255, 0.3);
    }
    
    [data-testid="stSidebar"] > div:first-child {
        background: transparent;
    }
    
    /* Header styling */
    .main-header {
        font-size: 2.8rem;
        font-weight: 700;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        margin-bottom: 2rem;
        letter-spacing: -0.5px;
    }
    
    /* Glass card effect */
    .glass-card {
        background: rgba(255, 255, 255, 0.8);
        backdrop-filter: blur(10px) saturate(180%);
        -webkit-backdrop-filter: blur(10px) saturate(180%);
        border-radius: 20px;
        border: 1px solid rgba(255, 255, 255, 0.3);
        padding: 24px;
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
        transition: all 0.3s ease;
    }
    
    .glass-card:hover {
        box-shadow: 0 12px 48px rgba(0, 0, 0, 0.15);
        transform: translateY(-2px);
    }
    
    /* Success box */
    .success-box {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 20px;
        padding: 28px;
        margin: 24px 0;
        box-shadow: 0 8px 32px rgba(102, 126, 234, 0.3);
        color: white;
        border: none;
    }
    
    .success-box h3 {
        color: white;
        font-weight: 600;
        font-size: 1.5rem;
        margin-bottom: 8px;
    }
    
    .success-box p {
        color: rgba(255, 255, 255, 0.9);
        font-size: 1rem;
    }
    
    /* Stat box */
    .stat-box {
        background: rgba(255, 255, 255, 0.9);
        backdrop-filter: blur(10px);
        border-radius: 16px;
        padding: 20px;
        text-align: center;
        border: 1px solid rgba(255, 255, 255, 0.3);
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.08);
        transition: all 0.3s ease;
    }
    
    .stat-box:hover {
        transform: translateY(-4px);
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.12);
    }
    
    /* Buttons */
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 12px;
        border: none;
        padding: 12px 32px;
        font-weight: 600;
        font-size: 1rem;
        box-shadow: 0 4px 16px rgba(102, 126, 234, 0.4);
        transition: all 0.3s ease;
        letter-spacing: 0.3px;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 24px rgba(102, 126, 234, 0.5);
    }
    
    .stButton > button:active {
        transform: translateY(0);
    }
    
    /* Download button */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        color: white;
        border-radius: 12px;
        border: none;
        padding: 14px 32px;
        font-weight: 600;
        font-size: 1.05rem;
        box-shadow: 0 4px 16px rgba(245, 87, 108, 0.4);
        transition: all 0.3s ease;
    }
    
    .stDownloadButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 24px rgba(245, 87, 108, 0.5);
    }
    
    /* File uploader */
    [data-testid="stFileUploader"] {
        background: rgba(255, 255, 255, 0.9);
        backdrop-filter: blur(10px);
        border-radius: 16px;
        padding: 24px;
        border: 2px dashed rgba(102, 126, 234, 0.3);
        transition: all 0.3s ease;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: rgba(102, 126, 234, 0.6);
        background: rgba(255, 255, 255, 1);
    }
    
    /* Selectbox */
    .stSelectbox > div > div {
        background: rgba(255, 255, 255, 0.95);
        border-radius: 12px;
        border: 1px solid rgba(102, 126, 234, 0.2);
    }
    
    /* Metrics */
    [data-testid="stMetricValue"] {
        font-size: 2rem;
        font-weight: 700;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    [data-testid="stMetricLabel"] {
        font-weight: 600;
        color: #64748b;
        font-size: 0.9rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    /* Dataframe */
    .stDataFrame {
        background: rgba(255, 255, 255, 0.95);
        border-radius: 16px;
        overflow: hidden;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.08);
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background: rgba(255, 255, 255, 0.9);
        border-radius: 12px;
        font-weight: 600;
        color: #334155;
    }
    
    /* Progress bar */
    .stProgress > div > div {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        border-radius: 10px;
    }
    
    /* Info/Warning/Success boxes */
    .stAlert {
        background: rgba(255, 255, 255, 0.9);
        backdrop-filter: blur(10px);
        border-radius: 12px;
        border-left: 4px solid;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.08);
    }
    
    /* Sidebar content */
    .css-1d391kg, .css-1v0mbdj {
        padding: 1.5rem;
    }
    
    /* Sidebar headers */
    [data-testid="stSidebar"] h2 {
        color: #1e293b;
        font-weight: 700;
        font-size: 1.5rem;
    }
    
    [data-testid="stSidebar"] h3 {
        color: #475569;
        font-weight: 600;
        font-size: 1.1rem;
    }
    
    /* Sidebar markdown */
    [data-testid="stSidebar"] .stMarkdown {
        color: #64748b;
    }
    
    /* Footer */
    .app-footer {
        text-align: center;
        padding: 32px 0;
        margin-top: 48px;
        color: #64748b;
        font-size: 0.95rem;
        border-top: 1px solid rgba(255, 255, 255, 0.3);
        background: rgba(255, 255, 255, 0.5);
        backdrop-filter: blur(10px);
        border-radius: 20px;
        margin: 48px -1rem 0;
        padding: 24px;
    }
    
    .app-footer strong {
        display: block;
        font-weight: 600;
        font-size: 1.1rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 8px;
    }
    
    .app-footer span {
        display: block;
        margin-top: 8px;
        font-size: 0.9rem;
        color: #94a3b8;
    }
    
    /* Subheaders */
    h2, h3 {
        color: #1e293b;
        font-weight: 600;
    }
    
    /* Dividers */
    hr {
        margin: 2rem 0;
        border: none;
        height: 1px;
        background: linear-gradient(90deg, transparent, rgba(102, 126, 234, 0.3), transparent);
    }
    
    /* Remove default Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Animations */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    .main > div {
        animation: fadeIn 0.5s ease-out;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Header
    st.markdown('<h1 class="main-header">Invoice Automation System</h1>', unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.image(
            "assets/rv_solutions_logo.png",
            use_container_width=True
        )

        st.markdown(
            "<h2 style='text-align: center; margin-top: 1.5rem;'>Settings</h2>",
            unsafe_allow_html=True
        )
        
        # Brand selection
        selected_brand = st.selectbox(
            "Select Brand",
            list(BRAND_CONFIGS.keys()),
            index=0
        )
        
        st.markdown("---")
        st.markdown("### Instructions")
        st.markdown("""
        1. Select your brand
        2. Upload the raw data Excel file
        3. Click Generate Invoices
        4. Download all invoices as ZIP
        """)
        
        st.markdown("---")
        st.markdown("### Required Columns")
        config = BRAND_CONFIGS[selected_brand]
        for col in config['required_columns']:
            st.markdown(f"• {col}")
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # File upload
        st.subheader("📂 Upload Raw Data")
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload the raw billing data Excel file"
        )
        
        if uploaded_file:
            try:
                # Preview data
                df = pd.read_excel(uploaded_file)
                st.success(f"File loaded successfully! ({len(df)} rows, {len(df.columns)} columns)")
                
                with st.expander("Data Preview", expanded=False):
                    st.dataframe(df.head(), use_container_width=True)
                    
                    # Show missing columns
                    missing_cols = [col for col in config['required_columns'] if col not in df.columns]
                    if missing_cols:
                        st.warning(f"⚠️ Missing columns: {', '.join(missing_cols)}")
                    else:
                        st.success("All required columns present")
            
            except Exception as e:
                st.error(f"X Error reading file: {str(e)}")
    
    with col2:
        # Statistics Panel
        st.subheader("Statistics")
        
        if uploaded_file:
            try:
                df = pd.read_excel(uploaded_file)
                asc_column = config['asc_column']
                
                if asc_column in df.columns:
                    total_ascs = df[asc_column].nunique()
                    total_records = len(df)
                    
                    # Determine amount column based on brand
                    if selected_brand == 'Harman':
                        amount_column = 'Call Charge'
                    elif selected_brand == 'Philips':
                        amount_column = 'Final Amount'
                    elif selected_brand == 'LifeLong':
                        amount_column = 'Final Amount'
                    else:
                        amount_column = 'Earning'
                    
                    # Check if amount column exists
                    if amount_column in df.columns:
                        total_amount = df[amount_column].sum()
                    else:
                        # Try alternative column names
                        amount_column_found = False
                        for col in df.columns:
                            if 'earning' in str(col).lower() or 'amount' in str(col).lower() or 'charge' in str(col).lower():
                                total_amount = df[col].sum()
                                amount_column_found = True
                                break
                        if not amount_column_found:
                            total_amount = 0
                            st.warning(f"⚠️ Amount column not found. Tried: Earning, Call Charge, Final Amount")
                    
                    st.metric("Total ASCs", total_ascs)
                    st.metric("Total Records", total_records)
                    st.metric("Total Amount", f"₹{total_amount:,.2f}")
                else:
                    st.warning(f"⚠️ ASC column '{asc_column}' not found")
            except Exception as e:
                st.error(f"Error calculating statistics: {str(e)}")
                st.info("Upload file to see statistics")
        else:
            st.info("Upload file to see statistics")
    
    # Generate button
    st.markdown("---")
    
    if uploaded_file:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            generate_btn = st.button(
                "Generate All Invoices",
                type="primary",
                use_container_width=True,
                help="Click to generate invoices for all ASCs"
            )
        
        if generate_btn:
            with st.spinner("Processing invoices..."):
                try:
                    # Process invoices
                    processor = InvoiceProcessor(selected_brand, config)
                    uploaded_file.seek(0)  # Reset file pointer
                    results = processor.process_invoices(uploaded_file)
                    
                    # Progress bar
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    total_ascs = len(results)
                    processed = 0
                    
                    for asc_name, data in results.items():
                        processed += 1
                        progress = int((processed / total_ascs) * 100)
                        progress_bar.progress(progress)
                        status_text.text(f"Processing {asc_name}... ({processed}/{total_ascs})")
                    
                    # Create ZIP file
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for asc_name, data in results.items():
                            # Add invoice
                            safe_name = "".join(c for c in asc_name if c.isalnum() or c in (' ', '-', '_')).strip()
                            invoice_filename = f"Invoice_{safe_name}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                            zip_file.writestr(invoice_filename, data['invoice'])
                            
                            # Add raw data
                            raw_filename = f"RawData_{safe_name}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                            zip_file.writestr(raw_filename, data['raw_data'])
                    
                    zip_buffer.seek(0)
                    
                    # Success message
                    st.markdown("""
                    <div class="success-box">
                        <h3>Invoice Generation Complete!</h3>
                        <p>Successfully generated invoices for all ASCs. Ready to download!</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Download button
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    zip_filename = f"{selected_brand}_Invoices_{timestamp}.zip"
                    
                    st.download_button(
                        label="Download All Invoices (ZIP)",
                        data=zip_buffer,
                        file_name=zip_filename,
                        mime="application/zip",
                        use_container_width=True,
                        help="Click to download all generated invoices"
                    )
                    
                    # Summary table
                    st.subheader("Generation Summary")
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

                    # Totals
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total ASCs", len(results))
                    with col2:
                        st.metric("Total Records", total_records)
                    with col3:
                        st.metric("Total Amount", f"₹{total_earning:,.2f}")
                    
                except Exception as e:
                    st.error(f"X Error generating invoices: {str(e)}")
    
    else:
        st.info("Please upload an Excel file to begin")

    # Footer
    st.markdown("""
        <div class="app-footer">
            <strong>Crafted with precision by Harpinder Singh</strong>
            <span>© 2025 RV Solutions • All rights reserved</span>
        </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()