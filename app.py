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
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #2c3e50;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 20px;
        margin: 20px 0;
    }
    .stat-box {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 5px;
        padding: 15px;
        text-align: center;
    }
    .app-footer {
        text-align: center;
        padding: 20px 0;
        margin-top: 40px;
        color: #6c757d;
        font-size: 0.9rem;
        border-top: 1px solid #e9ecef;
    }
    .app-footer span {
        display: block;
        margin-top: 5px;
        font-size: 0.85rem;
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
            "<h2 style='text-align: center;'>Settings</h2>",
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
            st.markdown(f"- {col}")
    
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
                st.error(f"Error reading file: {str(e)}")
    
    with col2:
        # Statistics panel
        st.subheader("Statistics")
        
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
                        <h3>✅ Invoice Generation Complete!</h3>
                        <p>Successfully generated invoices for all ASCs.</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Download button
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    zip_filename = f"{selected_brand}_Invoices_{timestamp}.zip"
                    
                    st.download_button(
                        label="📥 Download All Invoices (ZIP)",
                        data=zip_buffer,
                        file_name=zip_filename,
                        mime="application/zip",
                        use_container_width=True,
                        help="Click to download all generated invoices"
                    )
                    
                    # Summary table
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
                    
                    # Totals
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
            <strong>Crafted with precision by Harpinder Singh</strong>
            <span>© RV Solutions</span>
        </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()