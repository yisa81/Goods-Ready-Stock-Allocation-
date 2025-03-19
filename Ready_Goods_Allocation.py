
from fastapi import FastAPI, UploadFile, File
import pandas as pd
from io import BytesIO
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

# Allow frontend requests
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/upload")
async def upload_files(sales_report: UploadFile = File(...), ready_goods: UploadFile = File(...)):
    # Read sales report
    sales_df = pd.read_excel(BytesIO(await sales_report.read()))
    ready_goods_df = pd.read_excel(BytesIO(await ready_goods.read()))
    
    # Ensure SKU column names match
    sales_df.rename(columns=lambda x: x.strip().lower(), inplace=True)
    ready_goods_df.rename(columns=lambda x: x.strip().lower(), inplace=True)
Store, Store, Store, Store, Store, Store, Store, Store, Store, Store, Store, Store, Store, Store, me, Store

Hi All Please see today's CSV attached. If you have any ques    
    # Select required columns from sales report
    selected_columns = [
        'sku', 'life cycle', 'b - last 3 mths avg', 'avg pc/store (3mths)', 'count',
        'conant soh', 'ocean soh', 'status', 'mthly max avg sales (a,b & c)'
    ]
    sales_df = sales_df[selected_columns]
    
    # Merge based on SKU
    merged_df = ready_goods_df.merge(sales_df, on='sku', how='left')
    
    # Add new columns
    merged_df['conant qty'] = 0
    merged_df['ocean qty'] = 0
    merged_df['conant msoh'] = (merged_df['conant soh'] + merged_df['conant qty']) / (merged_df['mthly max avg sales (a,b & c)'] * 0.6)
    merged_df['ocean msoh'] = (merged_df['ocean soh'] + merged_df['ocean qty']) / (merged_df['mthly max avg sales (a,b & c)'] * 0.4)
    
    # Convert to JSON for frontend
    return merged_df.to_json(orient='records')
