import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from sklearn.linear_model import LinearRegression
import plotly.express as px
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
        return f"{year_gregorian}-{month_num}-{day}"  # Return as string in yyyy-mm-dd format
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
    "Date", "Opening Price", "Highest Price", "Lowest Price", "Average Price", "Closing Price",
    "Change", "Change (%)", "Volume (Thousand Shares)", "Value (Million Baht)",
    "SET Index", "SET Change (%)"
]

# Parse and convert date column to yyyy-mm-dd string format
df["Date"] = df["Date"].apply(parse_thai_date)

# Drop rows with invalid dates or NaN values
df = df.dropna(subset=["Date", "Closing Price"])

# Convert Date column to datetime for sorting and plotting
df["Date"] = pd.to_datetime(df["Date"], format="%Y-%m-%d")

# Sort data by date
df_sorted = df.sort_values("Date")

# Convert Date back to string format yyyy-mm-dd for display
df["Date"] = df["Date"].dt.strftime("%Y-%m-%d")
df_sorted["Date"] = df_sorted["Date"].dt.strftime("%Y-%m-%d")

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

# Display title
st.title("📊 PTT stock data analysis")
st.subheader("Display rows from Excel file ")

# Select how many rows to display
howmanyrow = st.selectbox("Select how many rows to Display", ["5", "20", "30", "All"])
if howmanyrow == "5":
    st.dataframe(df.head(5))
elif howmanyrow == "20":
    st.dataframe(df.head(20))
elif howmanyrow == "30":
    st.dataframe(df.head(30))
elif howmanyrow == "All":
    st.dataframe(df)

# Prepare data for Linear Regression
X = pd.to_datetime(df_sorted["Date"]).map(lambda x: x.toordinal()).values.reshape(-1, 1)
y = df_sorted["Closing Price"].values

# Create and fit Linear Regression model
model = LinearRegression()
model.fit(X, y)
trend = model.predict(X)

# Create Functions
def calculate_macd(df, col='Closing Price'):
    ema12 = df[col].ewm(span=12).mean()
    ema26 = df[col].ewm(span=26).mean()
    macd = ema12 - ema26
    signal = macd.ewm(span=9).mean()
    return macd, signal, macd - signal

st.title("PTT Stock Chart")

chart_type = st.selectbox("Select Indicators Chart", ["Linear Regression", "Interactive", "MACD"])
if chart_type == "Linear Regression":
    X = pd.to_datetime(df_sorted["Date"]).map(pd.Timestamp.toordinal).values.reshape(-1, 1)
    y = df_sorted["Closing Price"].values
    model = LinearRegression().fit(X, y)
    trend = model.predict(X)

    plt.figure(figsize=(10, 5))
    plt.plot(pd.to_datetime(df_sorted["Date"]), y, label="Actual Closing Price")
    plt.plot(pd.to_datetime(df_sorted["Date"]), trend, label="Trend (Linear Regression)", linestyle="--", color="red")
    plt.legend()
    plt.xlabel("Date")
    plt.ylabel("Price (THB)")
    plt.grid(True)
    st.pyplot(plt)
elif chart_type == "Interactive":
    fig = px.line(df, x='Date', y='Closing Price', title='PTT Stock Price')
    fig.update_layout(xaxis_title='Date', yaxis_title='Price')
    st.plotly_chart(fig, use_container_width=True)
elif chart_type == "MACD":
    macd, signal, hist = calculate_macd(df_sorted)
    fig = go.Figure([
        go.Bar(x=df_sorted['Date'], y=hist, name='Histogram', marker_color='red'),
        go.Scatter(x=df_sorted['Date'], y=macd, name='MACD', line=dict(color='blue')),
        go.Scatter(x=df_sorted['Date'], y=signal, name='Signal', line=dict(color='orange'))
    ])
    fig.update_layout(title='MACD Chart', hovermode='x unified')
    st.plotly_chart(fig, use_container_width=True)