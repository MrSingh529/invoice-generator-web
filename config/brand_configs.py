BRAND_CONFIGS = {
    'Amazon': {
        'asc_column': 'ASC Name',
        'required_columns': [
            'ASC Name', 'Earning', 'COD', 'quantity', 'category',
            'Owner Name', 'Contact No.', 'PAN No.', 'GST No.', 'Address'
        ],
        'invoice_template': {
            'company_name': 'RV Solutions Private Limited',
            'company_address': 'D-59, Sector-2, Gautam Buddh Nagar, Noida, Uttar Pradesh Noida-201301.',
            'gst_template': 'IGST',  # 'IGST' for inter-state, 'CGST/SGST' for intra-state
            'sac_code': '998715',
            'declaration': "We declare that this invoice shows the actual price of the goods/services described..."
        }
    },
    # Add more brands here as needed
    # 'Samsung': { ... },
    # 'LG': { ... },
}