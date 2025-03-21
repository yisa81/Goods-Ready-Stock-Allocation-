import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.title("🛒 Goods Ready Stock Allocation (with Excel formulas)")

# Upload files
sales_file = st.file_uploader("Upload Sales Report (.xlsx)", type="xlsx")
goods_file = st.file_uploader("Upload Ready Goods (.xlsx)", type="xlsx")

if sales_file and goods_file:
    try:
        # Read into DataFrames
        sales_df = pd.read_excel(sales_file, engine="openpyxl")
        goods_df = pd.read_excel(goods_file, engine="openpyxl")

        # Normalize headers
        sales_df.columns = sales_df.columns.str.strip().str.lower()
        goods_df.columns = goods_df.columns.str.strip().str.lower()

        # Merge on SKU
        merged_df = goods_df.merge(sales_df, on='sku', how='left')

        # Add static column
        if 'conant soh' in merged_df.columns and 'mthly max avg sales (a,b & c)' in merged_df.columns:
            merged_df['100% conant msoh'] = (merged_df['conant soh'] / merged_df['mthly max avg sales (a,b & c)']).round(2)
        else:
            merged_df['100% conant msoh'] = None

        # Add blank input columns
        merged_df['conant qty'] = 0
        merged_df['ocean qty'] = 0
        merged_df['conant msoh'] = ""  # Will add formula later
        merged_df['ocean msoh'] = ""

        # Save to temporary Excel (without formulas)
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            merged_df.to_excel(writer, index=False, sheet_name="Data")

        # Re-open and inject formulas
        output.seek(0)
        wb = load_workbook(output)
        ws = wb["Data"]

        headers = list(merged_df.columns)
        row_count = ws.max_row

        # Get column indexes (1-based)
        conant_soh_col = headers.index("conant soh") + 1
        conant_qty_col = headers.index("conant qty") + 1
        mthly_max_col = headers.index("mthly max avg sales (a,b & c)") + 1
        conant_msoh_col = headers.index("conant msoh") + 1

        ocean_soh_col = headers.index("ocean soh") + 1
        ocean_qty_col = headers.index("ocean qty") + 1
        ocean_msoh_col = headers.index("ocean msoh") + 1

        # Find "Ready to Ship" column and apply yellow highlight
        ready_to_ship_col = headers.index("ready to ship") + 1 if "ready to ship" in headers else None
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        
        for row in range(2, row_count + 1):
            # Excel formula using RC
            ws.cell(row=row, column=conant_msoh_col).value = f"=ROUND(({ws.cell(row, conant_soh_col).coordinate}+{ws.cell(row, conant_qty_col).coordinate})/({ws.cell(row, mthly_max_col).coordinate}*0.6), 2)"
            ws.cell(row=row, column=ocean_msoh_col).value = f"=ROUND(({ws.cell(row, ocean_soh_col).coordinate}+{ws.cell(row, ocean_qty_col).coordinate})/({ws.cell(row, mthly_max_col).coordinate}*0.4), 2)"

            # Highlight "Ready to Ship" column
            if ready_to_ship_col:
                ws.cell(row=row, column=ready_to_ship_col).fill = yellow_fill
                
        # Group and hide specified columns
        columns_to_hide = [(10, 27), (31, 36), (44, 60), (63, 64)]  # (J to AA), (AE to AJ), (AR to BH), (BK to BL)
        for col_start, col_end in columns_to_hide:
            ws.column_dimensions[ws.cell(row=1, column=col_start).column_letter].hidden = True
            for col in range(col_start, col_end + 1):
                ws.column_dimensions[ws.cell(row=1, column=col).column_letter].hidden = True        
                
        # Save final output
        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        st.success("✅ File processed successfully!")
        st.download_button(
            label="📥 Download Final Excel with Formulas",
            data=final_output,
            file_name="Merged_Goods_Allocation_with_Formulas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Something went wrong: {e}")

