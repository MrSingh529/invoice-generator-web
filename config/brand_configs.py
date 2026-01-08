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
            'gst_template': 'IGST',
            'sac_code': '998715',
            'declaration': "Declaration:- We declare that this invoice shows the actual price of the goods/services described and that all particulars are true and correct.\n\n* In case of non reflection of the GST amount in GSTR-2B of RV Solutions Pvt. Ltd. within 30th-June of Next Financial year, we agree to pay RV Solutions Pvt. Ltd. the GST amount along with interest @18% p.a. on delayed payment."
        }
    },
    'Harman': {
        'asc_column': 'ASC Name',
        'required_columns': [
            'ASC Name', 'Description', 'Call Charge',
            'Owner Name', 'Contact No.', 'PAN No.', 'GST No.', 'Address'
        ],
        'invoice_template': {
            'company_name': 'RV Solutions Private Limited',
            'company_address': 'D-59, Sector-2, Gautam Buddh Nagar, Noida, Uttar Pradesh Noida-201301.',
            'gst_template': 'IGST',
            'sac_code': '998715',
            'declaration': "Declaration:- We declare that this invoice shows the actual price of the goods/services described and that all particulars are true and correct.\n\nTerms: * In case of non reflection of the GST amount in GSTR-2B of RV Solutions Pvt. Ltd. within 30th-June of Next Financial year, we agree to pay RV Solutions Pvt. Ltd. the GST amount along with interest @24% p.a. on delayed payment."
        },
        'harman_specific': {
            'invoice_title': 'Bill of Supply',
            'amount_column': 'Call Charge',
            'description_column': 'Description',
            'exclude_gst_for_freelancer': True,
            'exclude_cod_section': True
        }
    },
    'Philips': {
        'asc_column': 'ASC Name',
        'required_columns': [
            'ASC Name', 'Category', 'Final Amount',
            'Owner Name', 'Contact No.', 'PAN No.', 'GST No.', 'Address'
        ],
        'invoice_template': {
            'company_name': 'RV Solutions Private Limited',
            'company_address': 'D-59, Sector-2, Gautam Buddh Nagar, Noida, Uttar Pradesh Noida-201301.',
            'gst_template': 'IGST',
            'sac_code': '998715',
            'declaration': "Declaration:- We declare that this invoice shows the actual price of the goods/services described and that all particulars are true and correct.\n\n* In case of non reflection of the GST amount in GSTR-2B of RV Solutions Pvt. Ltd. within 30th-June of Next Financial year, we agree to pay RV Solutions Pvt. Ltd. the GST amount along with interest @18% p.a. on delayed payment."
        },
        'philips_specific': {
            'invoice_title': 'Tax Invoice',
            'amount_column': 'Final Amount',
            'description_column': 'Category',
            'exclude_cod_section': True
        }
    },
    'LifeLong': {
        'asc_column': 'ASC Name',
        'required_columns': [
            'ASC Name', 'Description', 'Final Amount',
            'Owner Name', 'Contact No.', 'PAN No.', 'GST No.', 'Address'
        ],
        'invoice_template': {
            'company_name': 'RV Solutions Private Limited',
            'company_address': 'D-59, Sector-2, Gautam Buddh Nagar, Noida, Uttar Pradesh Noida-201301.',
            'gst_template': 'IGST',
            'sac_code': '998715',
            'declaration': "Declaration:- We declare that this invoice shows the actual price of the goods/services described and that all particulars are true and correct.\n\nTerms: * In case of non reflection of the GST amount in GSTR-2B of RV Solutions Pvt. Ltd. within 30th-June of Next Financial year, we agree to pay RV Solutions Pvt. Ltd. the GST amount along with interest @24% p.a. on delayed payment."
        },
        'lifelong_specific': {
            'invoice_title': 'Bill of Supply',
            'amount_column': 'Final Amount',
            'description_column': 'Description',
            'exclude_gst_for_freelancer': True,
            'exclude_cod_section': True,
            'add_spacing_rows': True
        }
    },
    'CandorCRM': {
        'asc_column': 'ASC Name',
        'required_columns': [
            'ASC Name', 'Claim Status', 'Amount',
            'PAN No.', 'GST No.', 'Address'
        ],
        'invoice_template': {
            'company_name': 'RV Solutions Private Limited',
            'company_address': 'D-59, Sector-2, District-Gautam Buddh Nagar, Noida, Uttar Pradesh - 201301.',
            'gst_template': 'IGST',
            'sac_code': '998729',
            'declaration': "Declaration:- We declare that this invoice shows the actual price of the goods/services described and that all particulars are true and correct.\n\n* In case of non reflection of the GST amount in GSTR-2B of RV Solutions Pvt. Ltd. within 30th-June of Next Financial year, we agree to pay RV Solutions Pvt. Ltd. the GST amount along with interest @18% p.a. on delayed payment."
        },
        'candor_specific': {
            'invoice_title': 'Tax Invoice',
            'amount_column': 'Amount',
            'description_column': 'Claim Status',
            'exclude_cod_section': True,
            'show_sac_code_row': True,
            'buyer_label': True,
            'show_contact_no': True,
            'headers': ['Description', 'Quantity', 'Rate', 'Amount'],
            'grand_total_label': 'Grand Total'
        }
    }
}