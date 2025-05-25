import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from sklearn.linear_model import LinearRegression
import plotly.express as px
from datetime import datetime
import plotly.graph_objects as go

# Mapping Thai month abbreviations to numbers
thai_month_map = {
    'ม.ค.': '01', 'ก.พ.': '02', 'มี.ค.': '03', 'เม.ย.': '04', 'พ.ค.': '05', 
    'มิ.ย.': '06', 'ก.ค.': '07', 'ส.ค.': '08', 'ก.ย.': '09', 'ต.ค.': '10', 
    'พ.ย.': '11', 'ธ.ค.': '12'
}

# Function to parse Thai date and convert to Gregorian
def parse_thai_date(date_str):
    try:
        day, month, year = date_str.split()
        month_num = thai_month_map.get(month.strip(), '01')  # Default to '01' if month not found
        year_gregorian = int(year) - 543  # Convert Buddhist year to Gregorian
        return pd.to_datetime(f"{day}-{month_num}-{year_gregorian}", format="%d-%m-%Y")
    except Exception as e:
        st.warning(f"Error parsing date {date_str}: {e}")
        return pd.NaT

# Load data
try:
    df = pd.read_excel("PTT-SET-23May2025-6M.xlsx", sheet_name="PTT", skiprows=1)
except Exception as e:
    st.error(f"Error loading Excel file: {e}")
    st.stop()


# Set column names
df.columns = [
    "วันที่", "ราคาเปิด", "ราคาสูงสุด", "ราคาต่ำสุด", "ราคาเฉลี่ย", "ราคาปิด",
    "เปลี่ยนแปลง", "เปลี่ยนแปลง(%)", "ปริมาณ(พันหุ้น)", "มูลค่า(ล้านบาท)",
    "SET Index", "SET เปลี่ยนแปลง(%)"
]

# Parse and convert date column
df["วันที่"] = df["วันที่"].apply(parse_thai_date)

# Drop rows with invalid dates or NaN values
df = df.dropna(subset=["วันที่", "ราคาปิด"])

# Sort data by date
df_sorted = df.sort_values("วันที่")

# Set index starting from 1
df.index = range(1, len(df) + 1)

# Add CSS for table styling
st.markdown(
    """
    <style>
    .dataframe {
        width: 864px !important;  # Match graph width (12 inches * 72 pixels)
        overflow-x: auto;
        margin: 0 auto;  # Center the table
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Display title and table
st.title("📊 วิเคราะห์ข้อมูลหุ้น PTT")
st.subheader("ข้อมูลย้อนหลัง")
st.dataframe(df.head())




# Prepare data for Linear Regression
X = df_sorted["วันที่"].map(lambda x: x.toordinal()).values.reshape(-1, 1)
y = df_sorted["ราคาปิด"].values

# Create and fit Linear Regression model
model = LinearRegression()
model.fit(X, y)
trend = model.predict(X)


#Create Funtions
def calculate_macd(df, col='ราคาปิด'):
    ema12 = df[col].ewm(span=12).mean()
    ema26 = df[col].ewm(span=26).mean()
    macd = ema12 - ema26
    signal = macd.ewm(span=9).mean()
    return macd, signal, macd - signal
 

st.title("PTT Stock Chart")

chart_type = st.selectbox("Select Indicators Chart", ["Linear Regression", "Interactive", "MACD"])
if chart_type == "Linear Regression":
    X = df_sorted["วันที่"].map(pd.Timestamp.toordinal).values.reshape(-1, 1)
    y = df_sorted["ราคาปิด"].values
    model = LinearRegression().fit(X, y)
    trend = model.predict(X)

    plt.figure(figsize=(10, 5))
    plt.plot(df_sorted["วันที่"], y, label="Actual Closing Price")
    plt.plot(df_sorted["วันที่"], trend, label="Trend (Linear Regression)", linestyle="--", color="red")
    plt.legend()
    plt.xlabel("Date")
    plt.ylabel("Price (THB)")
    plt.grid(True)
    st.pyplot(plt)
elif chart_type == "Interactive":
    fig = px.line(df, x='วันที่', y='ราคาปิด', title='META Stock Price')
    fig.update_layout(xaxis_title='Date', yaxis_title='Price')
    st.plotly_chart(fig, use_container_width=True)
elif chart_type == "MACD":
    macd, signal, hist = calculate_macd(df_sorted)
    fig = go.Figure([
        go.Bar(x=df_sorted['วันที่'], y=hist, name='Histogram', marker_color='red'),
        go.Scatter(x=df_sorted['วันที่'], y=macd, name='MACD', line=dict(color='blue')),
        go.Scatter(x=df_sorted['วันที่'], y=signal, name='Signal', line=dict(color='orange'))
    ])
    fig.update_layout(title='MACD Chart', hovermode='x unified')
    st.plotly_chart(fig, use_container_width=True)    