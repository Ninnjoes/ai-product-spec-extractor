import os
import pandas as pd
import plotly.express as px
import streamlit as st

# 1. ตั้งค่าหน้าเว็บให้กว้างเต็มจอ
st.set_page_config(
    page_title="AI Product Market Analytics",
    page_icon="📊",
    layout="wide"
)

@st.cache_data
def load_data(file_source) -> pd.DataFrame:
    if isinstance(file_source, str) and file_source.endswith(".csv"):
        df = pd.read_csv(file_source)
    else:
        df = pd.read_excel(file_source, sheet_name="Enriched_Products")
    if "Color" in df.columns:
        df["Color"] = df["Color"].astype("string").str.strip().str.title()
    return df

LOCAL_EXCEL = "enriched_products_report.xlsx"
SAMPLE_CSV = "sample_data/enriched_products_sample.csv"

st.title("📊 AI-Enriched Product Market Analytics")
st.caption("End-to-End Pipeline: Web Scraping → Gemini AI Spec Extraction (Pydantic) → Interactive BI Dashboard")

uploaded_file = st.sidebar.file_uploader("📂 อัปโหลดรายงาน Excel / CSV ของคุณเอง", type=["xlsx", "csv"])

if uploaded_file is not None:
    df = load_data(uploaded_file)
elif os.path.exists(LOCAL_EXCEL):
    df = load_data(LOCAL_EXCEL)
elif os.path.exists(SAMPLE_CSV):
    df = load_data(SAMPLE_CSV)
    st.info("ℹ️ กำลังแสดงผลจาก Sample Dataset (100 รายการ) — สามารถอัปโหลดไฟล์ของตัวเองได้ที่แถบซ้าย")
else:
    st.error("ไม่พบไฟล์ข้อมูล กรุณารัน `python enrich_pipeline.py` หรืออัปโหลดไฟล์")
    st.stop()
    
search_text = st.sidebar.text_input("ค้นหาชื่อรุ่น / สเปค (เช่น OLED, M5, TUF)", "")

all_brands = sorted(df["Brand"].dropna().unique().tolist())
selected_brands = st.sidebar.multiselect("เลือกยี่ห้อ (Brand)", all_brands, default=all_brands)

all_types = sorted(df["Device_Type"].dropna().unique().tolist())
selected_types = st.sidebar.multiselect("ประเภทอุปกรณ์ (Device Type)", all_types, default=all_types)

min_p, max_p = int(df["Price_THB"].min()), int(df["Price_THB"].max())
price_range = st.sidebar.slider("ช่วงราคา (THB)", min_value=min_p, max_value=max_p, value=(min_p, max_p), step=500)

# กรอง DataFrame ตามค่าที่ผู้ใช้เลือก
filtered_df = df[
    (df["Brand"].isin(selected_brands)) &
    (df["Device_Type"].isin(selected_types)) &
    (df["Price_THB"] >= price_range[0]) &
    (df["Price_THB"] <= price_range[1])
]

if search_text.strip():
    filtered_df = filtered_df[
        filtered_df["Product_Name"].str.contains(search_text.strip(), case=False, na=False)
    ]

# 4. แสดง KPI Cards 4 ช่องด้านบน
col1, col2, col3, col4 = st.columns(4)
col1.metric("จำนวนสินค้าที่พบ", f"{len(filtered_df):,} รุ่น")
if len(filtered_df) > 0:
    col2.metric("ราคาเฉลี่ย", f"{filtered_df['Price_THB'].mean():,.0f} ฿")
    col3.metric("ราคาต่ำสุด", f"{filtered_df['Price_THB'].min():,.0f} ฿")
    col4.metric("ราคาสูงสุด", f"{filtered_df['Price_THB'].max():,.0f} ฿")
else:
    col2.metric("ราคาเฉลี่ย", "-")
    col3.metric("ราคาต่ำสุด", "-")
    col4.metric("ราคาสูงสุด", "-")

st.divider()

if len(filtered_df) == 0:
    st.warning("ไม่พบสินค้าที่ตรงกับเงื่อนไข กรุณาปรับ Filter ด้านซ้ายครับ")
    st.stop()

# 5. กราฟวิเคราะห์ 2 ฝั่ง (Plotly Interactive Charts)
chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    st.subheader("💰 ราคาเฉลี่ยแยกตามแบรนด์ (Avg Price by Brand)")
    brand_stats = (
        filtered_df.groupby("Brand", as_index=False)
        .agg(Avg_Price=("Price_THB", "mean"), Count=("Product_Name", "count"))
        .sort_values("Avg_Price", ascending=True)
    )
    fig_bar = px.bar(
        brand_stats,
        x="Avg_Price",
        y="Brand",
        orientation="h",
        text_auto=",.0f",
        color="Avg_Price",
        color_continuous_scale="Teal",
        hover_data=["Count"],
        labels={"Avg_Price": "ราคาเฉลี่ย (บาท)", "Brand": "แบรนด์", "Count": "จำนวนรุ่น"}
    )
    fig_bar.update_layout(showlegend=False, coloraxis_showscale=False, height=400)
    st.plotly_chart(fig_bar, use_container_width=True)

with chart_col2:
    st.subheader("🖥️ การกระจายตัวของราคาตามขนาดจอ (Price vs Screen Size)")
    screen_df = filtered_df.dropna(subset=["Screen_Inch"]).copy()
    if len(screen_df) > 0:
        screen_df["Screen_Label"] = screen_df["Screen_Inch"].astype(int).astype(str) + '"'
        fig_box = px.box(
            screen_df.sort_values("Screen_Inch"),
            x="Screen_Label",
            y="Price_THB",
            color="Brand",
            points="all",
            hover_data=["Product_Name"],
            labels={"Screen_Label": "ขนาดหน้าจอ", "Price_THB": "ราคา (บาท)"}
        )
        fig_box.update_layout(height=400)
        st.plotly_chart(fig_box, use_container_width=True)
    else:
        st.info("รายการที่เลือกไม่มีข้อมูลขนาดหน้าจอ (เช่น Mac mini)")

# 6. ตารางข้อมูล + ปุ่มดาวน์โหลด CSV
st.subheader("📋 รายการสินค้าที่ผ่านการกรอง")
st.dataframe(
    filtered_df[["Brand", "Device_Type", "Screen_Inch", "RAM_GB", "Storage_GB", "Color", "Price_THB", "Product_Name"]],
    use_container_width=True,
    hide_index=True
)

csv_bytes = filtered_df.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    label="⬇️ ดาวน์โหลดข้อมูลที่กรองแล้ว (.csv)",
    data=csv_bytes,
    file_name="filtered_market_data.csv",
    mime="text/csv"
)