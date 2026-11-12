import os
import time
from typing import Optional
import pandas as pd
from dotenv import load_dotenv
from google import genai
from pydantic import BaseModel, Field
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

load_dotenv()
api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    raise ValueError("❌ ไม่พบ GEMINI_API_KEY ในไฟล์ .env!")

client = genai.Client(api_key=api_key)
MODEL_NAME = "gemini-3.8-flash"

class ProductSpec(BaseModel):
    row_id: int = Field(description="ลำดับ row_id ที่ส่งเข้าไป")
    brand: str = Field(description="ยี่ห้อสินค้า (ตัวพิมพ์ใหญ่ทั้งหมด เช่น ACER, ASUS, HP, DELL, APPLE, LENOVO, MSI, GIGABYTE)")
    device_type: str = Field(description="ประเภทอุปกรณ์ เช่น Notebook, MacBook, Mac mini")
    ram_gb: Optional[int] = Field(default=None, description="ความจุ RAM เป็นตัวเลข GB (ถ้าไม่ระบุให้ใส่ null)")
    storage_gb: Optional[int] = Field(default=None, description="ความจุ SSD/Storage เป็น GB (เช่น 1TB = 1024, ถ้าไม่ระบุให้ใส่ null)")
    screen_inch: Optional[float] = Field(default=None, description="ขนาดหน้าจอเป็นนิ้ว (ถ้าไม่ระบุให้ใส่ null)")
    color: Optional[str] = Field(default=None, description="สีของตัวเครื่อง (ถ้าไม่ระบุให้ใส่ null)")

def extract_batch(batch_df: pd.DataFrame) -> list[dict]:
    """ส่งข้อมูลเป็นชุด (Batch) ให้ Gemini สกัดสเปคเป็น Structured JSON"""
    items_text = "\n".join([f"row_id={idx}: {row['Product_Name']}" for idx, row in batch_df.iterrows()])
    prompt = f"""
คุณคือผู้เชี่ยวชาญด้าน Data Engineering
จงสกัดสเปคสินค้าจากชื่อสินค้าด้านล่างให้ตรงตาม Schema อย่างเคร่งครัด
กฎสำคัญ:
1. ห้ามเดาข้อมูลที่ไม่มีในชื่อสินค้า ถ้าไม่มีให้คืนค่า null
2. brand ให้ใช้ตัวพิมพ์ใหญ่ทั้งหมด (เช่น APPLE, ACER, ASUS)
3. สำหรับ screen_inch: หากมีเลขบอกขนาดจอในชื่อรุ่นหรือสเปค เช่น "13-inch", "14", "15", "16", "17", "S14", "LITE 16", "AL15", "X1504" หรือ "15-FC" ให้สกัดตัวเลขขนาดจอนั้นออกมา (เช่น 14.0, 15.0, 16.0)

รายการสินค้า:
{items_text}
"""
    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model=MODEL_NAME,
                contents=prompt,
                config={
                    "response_mime_type": "application/json",
                    "response_schema": list[ProductSpec],
                    "temperature": 0.0,
                },
            )
            return [item.model_dump() for item in response.parsed]
        except Exception as e:
            print(f"⚠️ รอบที่ {attempt + 1} มีปัญหา: {e}")
            time.sleep(3)
    return []

def format_worksheet(ws, currency_cols=None):
    currency_cols = currency_cols or set()
    header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    data_font = Font(name="Calibri", size=11, color="1E293B")
    thin_side = Side(style="thin", color="CBD5E1")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 28
    col_names = [cell.value for cell in ws[1]]

    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = cell_border

    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 22
        is_even = (row_idx % 2 == 0)
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = cell_border
            if is_even:
                cell.fill = zebra_fill

            col_name = col_names[col_idx - 1]
            if col_name in currency_cols:
                cell.number_format = "#,##0.00" if isinstance(cell.value, float) and not cell.value.is_integer() else "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_name in {"RAM_GB", "Storage_GB", "Screen_Inch", "Total_Models"}:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_len = max((len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(1, ws.max_row + 1)), default=10)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 14), 55)

def run_pipeline():
    df = pd.read_excel("input_products.xlsx", sheet_name="Products_Data")
    df["row_id"] = df.index
    total_rows = len(df)
    batch_size = 20

    print(f"🚀 เริ่มรัน AI Enrichment Pipeline ทั้งหมด {total_rows} รายการ (Batch ละ {batch_size})...")
    extracted_records = []

    for start in range(0, total_rows, batch_size):
        end = min(start + batch_size, total_rows)
        print(f"🤖 กำลังส่งแถวที่ {start + 1} - {end} ให้ {MODEL_NAME}...")
        batch_df = df.iloc[start:end]
        results = extract_batch(batch_df)
        extracted_records.extend(results)
        time.sleep(1.5)  # หน่วงเวลากัน Rate Limit

    # รวมข้อมูลสเปคจาก AI กลับเข้าตารางเดิมด้วย row_id
    spec_df = pd.DataFrame(extracted_records).rename(columns={
        "brand": "Brand",
        "device_type": "Device_Type",
        "ram_gb": "RAM_GB",
        "storage_gb": "Storage_GB",
        "screen_inch": "Screen_Inch",
        "color": "Color",
    })

    merged_df = pd.merge(df, spec_df, on="row_id", how="left").drop(columns=["row_id"])
    # จัดเรียงคอลัมน์ให้อ่านง่าย
    merged_df["Color"] = merged_df["Color"].str.strip().str.title()
    cols_order = ["Brand", "Device_Type", "Screen_Inch", "RAM_GB", "Storage_GB", "Color", "Price_THB", "Product_Name", "Product_URL"]
    merged_df = merged_df[cols_order]

    # สร้างตารางวิเคราะห์ราคาแยกตาม Brand (Business Intelligence!)
    brand_summary = merged_df.groupby("Brand", as_index=False).agg(
        Total_Models=("Product_Name", "count"),
        Avg_Price_THB=("Price_THB", "mean"),
        Min_Price_THB=("Price_THB", "min"),
        Max_Price_THB=("Price_THB", "max"),
    ).sort_values(by="Total_Models", ascending=False)
    brand_summary["Avg_Price_THB"] = brand_summary["Avg_Price_THB"].round(2)

    output_file = "enriched_products_report.xlsx"
    price_cols = {"Price_THB", "Avg_Price_THB", "Min_Price_THB", "Max_Price_THB"}
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        merged_df.to_excel(writer, sheet_name="Enriched_Products", index=False)
        brand_summary.to_excel(writer, sheet_name="Brand_Analytics", index=False)
        format_worksheet(writer.sheets["Enriched_Products"], currency_cols=price_cols)
        format_worksheet(writer.sheets["Brand_Analytics"], currency_cols=price_cols)

    print(f"\n✅ เสร็จสมบูรณ์! สร้างไฟล์ '{output_file}' เรียบร้อยแล้ว")
    print("\n📊 สรุปข้อมูลแยกตามแบรนด์:")
    print(brand_summary.to_string(index=False))

if __name__ == "__main__":
    run_pipeline()