import streamlit as st
import pandas as pd
import plotly.express as px
import requests          # ← THÊM
import io                # ← THÊM

# Link share OneDrive (copy từ nút Share)
onedrive_link = "https://1drv.ms/x/c/139ee0f07655cc68/EZfwgI3AedZOgrWNhd1vEdABZVGlqsN0fwhhL88jamgIOQ?e=oVyVTo"

# Hàm chuyển đổi link share sang link download trực tiếp
def get_direct_download_link(share_link):
    if "1drv.ms" in share_link:
        base_url = "https://api.onedrive.com/v1.0/shares/u!"
        import base64
        link_bytes = base64.urlsafe_b64encode(share_link.encode("utf-8"))
        return f"{base_url}{link_bytes.decode('utf-8').rstrip('=')}/root/content"
    return share_link

direct_link = get_direct_download_link(onedrive_link)

# Tải file Excel từ OneDrive
try:
    response = requests.get(direct_link, timeout=20)
    response.raise_for_status()
except requests.exceptions.RequestException as e:
    st.error(f"Không tải được file từ OneDrive: {e}")
    st.stop()

df = pd.read_excel(io.BytesIO(response.content))

st.title("📊 Dashboard cập nhật từ OneDrive")

# Hiển thị bảng dữ liệu
st.dataframe(df)

# Ví dụ vẽ biểu đồ
if "Category" in df.columns and "Value" in df.columns:
    fig = px.bar(df, x="Category", y="Value", title="Biểu đồ mẫu")
    st.plotly_chart(fig)
else:
    st.warning("Hãy đảm bảo file Excel có cột 'Category' và 'Value'.")


