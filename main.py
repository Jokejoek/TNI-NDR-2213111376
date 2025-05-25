import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from sklearn.linear_model import LinearRegression
import plotly.express as px

# โหลดข้อมูล
df = pd.read_excel("PTT-SET-23May2025-6M.xlsx", sheet_name="PTT", skiprows=1)

# ตั้งชื่อคอลัมน์ใหม่
df.columns = [
    "วันที่", "ราคาเปิด", "ราคาสูงสุด", "ราคาต่ำสุด", "ราคาเฉลี่ย", "ราคาปิด",
    "เปลี่ยนแปลง", "เปลี่ยนแปลง(%)", "ปริมาณ(พันหุ้น)", "มูลค่า(ล้านบาท)",
    "SET Index", "SET เปลี่ยนแปลง(%)"
]

# แปลงวันที่จาก Buddhist Era (2568) เป็น Gregorian Era (2025)
def parse_thai_date(date_str):
    day, month, year = date_str.split('-')
    year = str(int(year) + 543)  # Convert Buddhist year to Gregorian year
    month = month.replace('May', '05')  # Adjust for the specific month format in the data
    return f"{day}-{month}-{year}"


# ตั้ง index ให้เริ่มที่ 1
df.index = range(1, len(df) + 1)

# ลบแถวที่มี NaN
df = df.dropna()

# เพิ่ม CSS เพื่อปรับขนาดตาราง
st.markdown(
    """
    <style>
    .dataframe {
        width: 864px !important;  # ปรับให้เท่ากับความกว้างของกราฟ (12 นิ้ว * 72 พิกเซล)
        overflow-x: auto;
        margin: 0 auto;  # จัดกึ่งกลาง
    }
    </style>
    """,
    unsafe_allow_html=True
)


# แสดงตาราง
st.title("📊 วิเคราะห์ข้อมูลหุ้น PTT")
st.subheader("ข้อมูลย้อนหลัง 5 แถวแรก")
st.dataframe(df.head())