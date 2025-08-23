import streamlit as st
import pandas as pd
import plotly.express as px

# OneDrive public link (direct download)
onedrive_url = "https://1drv.ms/x/c/139ee0f07655cc68/EZfwgI3AedZOgrWNhd1vEdABZVGlqsN0fwhhL88jamgIOQ?e=oVyVTo"

# Đọc dữ liệu Excel trực tiếp từ OneDrive
@st.cache_data
def load_data():
    return pd.read_excel(onedrive_url)

df = load_data()

st.title("📊 Dashboard cập nhật từ OneDrive")

# Hiển thị bảng dữ liệu
st.dataframe(df)

# Ví dụ vẽ biểu đồ
if "Category" in df.columns and "Value" in df.columns:
    fig = px.bar(df, x="Category", y="Value", title="Biểu đồ mẫu")
    st.plotly_chart(fig)
else:
    st.warning("Hãy đảm bảo file Excel có cột 'Category' và 'Value'.")

