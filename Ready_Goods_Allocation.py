import streamlit as st
import pandas as pd
from io import BytesIO

st.title("🛒 Goods Ready Stock Allocation")

# Upload Sales Report and Ready Goods files
sales_file = st.file_uploader("Upload Sales Report (.xlsx)", type="xlsx")
goods_file = st.file_uploader("Upload Ready Goods (.xlsx)", type="xlsx")

if sales_file and goods_file:
    try:
        sales_df = pd.read_excel(sales_file, engine="openpyxl")
        goods_df = pd.read_excel(goods_file, engine="openpyxl")

        # Clean headers
        sales_df.columns = sales_df.columns.str.strip().str.lower()
        goods_df.columns = goods_df.columns.str.strip().str.lower()

        # Merge files
        merged = goods_df.merge(sales_df, on="sku", how="left")

        # Add 100% Conant Msoh
        if 'conant soh' in merged.columns and 'mthly max avg sales (a,b & c)' in merged.columns:
            merged['100% conant msoh'] = (merged['conant soh'] / merged['mthly max avg sales (a,b & c)']).round(2)
        else:
            merged['100% conant msoh'] = None

        # Add calculation columns
        merged['conant qty'] = 0
        merged['ocean qty'] = 0
        merged['conant msoh'] = (
            (merged['conant soh'] + merged['conant qty']) /
            (merged['mthly max avg sales (a,b & c)'] * 0.6)
        ).round(2)
        merged['ocean msoh'] = (
            (merged['ocean soh'] + merged['ocean qty']) /
            (merged['mthly max avg sales (a,b & c)'] * 0.4)
        ).round(2)

        # Show result
        st.success("Merge complete. Preview below:")
        st.dataframe(merged)

        # Download as Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            merged.to_excel(writer, index=False)
        st.download_button(
            label="📥 Download Merged Excel",
            data=output.getvalue(),
            file_name="Merged_Goods_Allocation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error: {e}")
