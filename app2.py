import streamlit as st
import pandas as pd
import plotly.express as px
import requests
import io

onedrive_link = "https://1drv.ms/x/c/139ee0f07655cc68/EZfwgI3AedZOgrWNhd1vEdABZVGlqsN0fwhhL88jamgIOQ?e=oVyVTo"

def resolve_onedrive_download_link(share_link: str) -> str:
    """
    Trả về link download cuối cùng bằng cách FOLLOW REDIRECT từ link share.
    Áp dụng cho OneDrive personal & nhiều loại 1drv.ms (x, s, u,...).
    """
    try:
        # chỉ lấy URL cuối cùng, không cần tải toàn bộ nội dung
        with requests.get(share_link, allow_redirects=True, timeout=20, stream=True) as r:
            final_url = r.url  # URL sau chuỗi redirect
        # Nếu là trang xem file trên onedrive.live.com, ép tải về
        if "onedrive.live.com" in final_url and "download" not in final_url:
            sep = "&" if "?" in final_url else "?"
            final_url = f"{final_url}{sep}download=1"
        return final_url
    except requests.exceptions.RequestException:
        return share_link  # fallback

download_url = resolve_onedrive_download_link(onedrive_link)

try:
    resp = requests.get(download_url, timeout=30)
    resp.raise_for_status()
except requests.exceptions.RequestException as e:
    st.error(f"Không tải được file từ OneDrive: {e}")
    st.write("Debug URL:", download_url)  # gỡ lỗi
    st.stop()

# Đọc Excel
try:
    df = pd.read_excel(io.BytesIO(resp.content))
except Exception as e:
    st.error(f"Không đọc được Excel: {e}")
    st.stop()

st.title("📊 Dashboard cập nhật từ OneDrive")
st.dataframe(df)

if "Category" in df.columns and "Value" in df.columns:
    fig = px.bar(df, x="Category", y="Value", title="Biểu đồ mẫu")
    st.plotly_chart(fig)
else:
    st.warning("Hãy đảm bảo file Excel có cột 'Category' và 'Value'.")


