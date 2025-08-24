import streamlit as st
import pandas as pd
import plotly.express as px
import requests, io, base64

# Hàm tổng hợp: nhập link share OneDrive -> trả về DataFrame
def load_excel_from_onedrive(share_link: str) -> pd.DataFrame:
    try:
        # Tạo direct download link từ link share
        b64 = base64.urlsafe_b64encode(share_link.encode()).decode().rstrip("=")
        direct_url = f"https://api.onedrive.com/v1.0/shares/u!{b64}/root/content"

        # Tải file
        resp = requests.get(direct_url, timeout=30)
        resp.raise_for_status()
        
        # Đọc Excel
        return pd.read_excel(io.BytesIO(resp.content))
    except Exception as e:
        st.error(f"Lỗi khi tải/đọc file từ OneDrive: {e}")
        st.stop()

# ----------------------------
# Ví dụ sử dụng
onedrive_link = "https://1drv.ms/x/c/139ee0f07655cc68/EZfwgI3AedZOgrWNhd1vEdABZVGlqsN0fwhhL88jamgIOQ?e=bzzH8G"
df = load_excel_from_onedrive(onedrive_link)

st.title("📊 Dashboard cập nhật từ OneDrive")
st.dataframe(df)

# Vẽ thử biểu đồ nếu có cột phù hợp
if "Category" in df.columns and "Value" in df.columns:
    fig = px.bar(df, x="Category", y="Value", title="Biểu đồ mẫu")
    st.plotly_chart(fig)
else:
    st.warning("Hãy đảm bảo file Excel có cột 'Category' và 'Value'.")
