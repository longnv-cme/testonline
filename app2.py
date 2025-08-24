import streamlit as st
import pandas as pd
import plotly.express as px
import requests, io
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

st.set_page_config(page_title="OneDrive → DataFrame", page_icon="📊", layout="wide")

# -------------------- Helper: nhận diện response là file --------------------
def _looks_like_file(resp: requests.Response) -> bool:
    ct = (resp.headers.get("Content-Type") or "").lower()
    cd = (resp.headers.get("Content-Disposition") or "").lower()
    cl = resp.headers.get("Content-Length")
    excel_mimes = (
        "application/vnd.openxmlformats-officedocument",  # xlsx
        "application/vnd.ms-excel",                      # xls
    )
    if any(m in ct for m in excel_mimes) or "application/octet-stream" in ct or "attachment" in cd:
        return True
    try:
        return cl is not None and int(cl) > 1024  # >1KB
    except Exception:
        return False

# -------------------- Helper: chuyển 1drv.ms → direct download --------------------
def onedrive_share_to_direct_url(share_link: str, timeout: int = 30) -> str:
    """
    Trả về link download ẩn danh dùng được với requests.get(...)
    Thử lần lượt:
      1) Trích resid/authkey/cid → onedrive.live.com/download?cid=...&resid=...&authkey=...
      2) Doc.aspx + sourcedoc={GUID} → Download.aspx?UniqueId={GUID}&download=1
      3) Ép download=1 vào URL cuối
      4) Nếu chính URL cuối trả file thì dùng URL đó
    """
    s = requests.Session()
    r = s.get(share_link, allow_redirects=True, timeout=timeout, stream=True)
    urls = [h.url for h in r.history] + [r.url]

    # 1) Dựng download?cid=&resid=&authkey=
    for u in urls:
        pu = urlparse(u)
        q = parse_qs(pu.query)
        resid = (q.get("resid") or [None])[0]
        authkey = (q.get("authkey") or [None])[0]
        cid = (q.get("cid") or [None])[0]
        if resid and authkey:
            if not cid and "!" in resid:
                cid = resid.split("!", 1)[0]
            direct = f"https://onedrive.live.com/download?cid={cid}&resid={resid}&authkey={authkey}"
            g = s.get(direct, allow_redirects=True, timeout=timeout, stream=True)
            if g.status_code == 200 and _looks_like_file(g):
                return direct
            if 300 <= g.status_code < 400:
                g2 = s.get(g.headers.get("Location", direct), allow_redirects=True, timeout=timeout, stream=True)
                if g2.status_code == 200 and _looks_like_file(g2):
                    return direct

    # 2) Doc.aspx → Download.aspx?UniqueId={GUID}
    final_url = urls[-1]
    p = urlparse(final_url)
    if p.path.lower().endswith("/_layouts/15/doc.aspx"):
        q = parse_qs(p.query)
        sourcedoc = (q.get("sourcedoc") or [None])[0]  # {GUID}
        if sourcedoc:
            download_path = p.path.replace("/Doc.aspx", "/Download.aspx")
            download_qs = {"UniqueId": sourcedoc, "Translate": "false", "download": "1"}
            direct = urlunparse((p.scheme, p.netloc, download_path, "", urlencode(download_qs), ""))
            g = s.get(direct, allow_redirects=True, timeout=timeout, stream=True)
            if g.status_code == 200 and _looks_like_file(g):
                return direct

    # 3) Thêm download=1 vào URL cuối (nếu là onedrive.live.com)
    if "onedrive.live.com" in p.netloc:
        q = parse_qs(p.query)
        if "download" not in q:
            q["download"] = ["1"]
            direct = urlunparse((p.scheme, p.netloc, p.path, "", urlencode({k: v[0] for k, v in q.items()}), ""))
            g = s.get(direct, allow_redirects=True, timeout=timeout, stream=True)
            if g.status_code == 200 and _looks_like_file(g):
                return direct

    # 4) Nếu URL cuối trả file thì dùng luôn
    if r.status_code == 200 and _looks_like_file(r):
        return urls[-1]

    raise ValueError(
        "Không thể tạo link download ẩn danh từ link share này. "
        "Hãy đảm bảo 'Anyone with the link can view' và thử lại."
    )

# -------------------- UI --------------------
st.title("📦 OneDrive Share → 📊 Streamlit Dashboard")

with st.expander("Nhập link share OneDrive (1drv.ms hoặc onedrive.live.com)"):
    share_link = st.text_input(
        "Dán link share tại đây",
        placeholder="https://1drv.ms/x/...",
        value=""  # có thể để trống hoặc dán sẵn link mẫu của bạn
    )
    colA, colB = st.columns([1,1])
    with colA:
        run_btn = st.button("📥 Tải dữ liệu", type="primary")
    with colB:
        show_debug = st.toggle("Hiển thị thông tin gỡ lỗi", value=False)

if run_btn:
    if not share_link.strip():
        st.warning("Vui lòng dán link share OneDrive.")
        st.stop()

    try:
        direct_url = onedrive_share_to_direct_url(share_link.strip())
        if show_debug:
            st.code(f"Direct URL: {direct_url}", language="text")
    except Exception as e:
        st.error(f"Không tạo được direct URL: {e}")
        st.stop()

    # Tải nội dung
    try:
        resp = requests.get(direct_url, timeout=60)
        resp.raise_for_status()
        if show_debug:
            st.write("Response headers:", dict(resp.headers))
    except requests.exceptions.RequestException as e:
        st.error(f"Không tải được file từ OneDrive: {e}")
        if show_debug:
            st.write("Tried URL:", direct_url)
        st.stop()

    # Quyết định đọc CSV hay Excel
    content_type = (resp.headers.get("Content-Type") or "").lower()
    dispo = (resp.headers.get("Content-Disposition") or "").lower()
    filename = ""
    if "filename=" in dispo:
        # lấy tên file từ Content-Disposition nếu có
        filename = dispo.split("filename=", 1)[1].strip('"; ')

    try:
        if filename.endswith(".csv") or "text/csv" in content_type:
            df = pd.read_csv(io.BytesIO(resp.content))
        else:
            # Mặc định đọc Excel
            df = pd.read_excel(io.BytesIO(resp.content))
    except Exception as e:
        st.error(f"Không đọc được dữ liệu: {e}")
        st.stop()

    st.success("Đã tải và đọc dữ liệu thành công!")
    st.dataframe(df, use_container_width=True)

    # Ví dụ vẽ biểu đồ nếu có cột phù hợp
    if {"Category", "Value"}.issubset(df.columns):
        fig = px.bar(df, x="Category", y="Value", title="Biểu đồ mẫu")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Thêm 2 cột 'Category' và 'Value' để xem ví dụ biểu đồ, hoặc sửa code để phù hợp dữ liệu của bạn.")
