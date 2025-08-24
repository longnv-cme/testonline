import requests
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

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
        return cl is not None and int(cl) > 1024
    except Exception:
        return False

def _try_get(url: str, session: requests.Session, timeout: int = 30) -> requests.Response | None:
    try:
        r = session.get(url, allow_redirects=True, timeout=timeout, stream=True)
        if r.status_code == 200 and _looks_like_file(r):
            return r
        return None
    except requests.exceptions.RequestException:
        return None

def onedrive_to_direct_url(share_or_embed_link: str, timeout: int = 30) -> str:
    """
    Nhận link 1drv.ms / onedrive.live.com (share hoặc embed) -> trả direct download URL.
    Chiến lược:
      A) Nếu là EMBED: embed?resid=...&authkey=...  -> download?resid=...&authkey=...
      B) Theo redirect từ link share để lấy resid/authkey -> download?...
      C) Doc.aspx + sourcedoc -> Download.aspx?UniqueId=...
      D) Nếu Download.aspx ở my.microsoftpersonalcontent.com bị 401 -> đổi sang onedrive.live.com
    """
    s = requests.Session()
    u = urlparse(share_or_embed_link)

    # A) Link EMBED → DOWNLOAD
    if "onedrive.live.com" in u.netloc and "/embed" in u.path:
        q = parse_qs(u.query)
        resid = (q.get("resid") or [None])[0]
        authkey = (q.get("authkey") or [None])[0]
        if resid and authkey:
            return f"https://onedrive.live.com/download?resid={resid}&authkey={authkey}"

    # B) Theo redirect từ link SHARE
    r = s.get(share_or_embed_link, allow_redirects=True, timeout=timeout, stream=True)
    chain = [h.url for h in r.history] + [r.url]

    # B1) bắt resid/authkey/cid ở bất kỳ bước nào
    for url in chain:
        pu = urlparse(url)
        if "onedrive.live.com" not in pu.netloc and "my.microsoftpersonalcontent.com" not in pu.netloc:
            continue
        q = parse_qs(pu.query)
        resid = (q.get("resid") or [None])[0]
        authkey = (q.get("authkey") or [None])[0]
        cid = (q.get("cid") or [None])[0]
        if resid and authkey:
            if not cid and "!" in resid:
                cid = resid.split("!", 1)[0]
            return f"https://onedrive.live.com/download?cid={cid}&resid={resid}&authkey={authkey}"

    # C) Dạng Doc.aspx + sourcedoc -> Download.aspx
    final_url = chain[-1]
    p = urlparse(final_url)
    if p.path.lower().endswith("/_layouts/15/doc.aspx"):
        q = parse_qs(p.query)
        sourcedoc = (q.get("sourcedoc") or [None])[0]  # {GUID}
        if sourcedoc:
            dl_path = p.path.replace("/Doc.aspx", "/Download.aspx")
            dl_qs = {"UniqueId": sourcedoc, "Translate": "false", "download": "1"}
            # ưu tiên onedrive.live.com
            return urlunparse(("https", "onedrive.live.com", dl_path, "", urlencode(dl_qs), ""))

    # D) nếu URL cuối cùng đã trả file thì dùng luôn
    if r.status_code == 200 and _looks_like_file(r):
        return chain[-1]

    # nếu không tìm được
    raise ValueError("Không thể tạo direct URL từ link đã cung cấp.")

def fetch_onedrive_file(share_or_embed_link: str, timeout: int = 45) -> bytes:
    """
    Tạo direct URL rồi tải file. Có fallback chuyển domain my.microsoftpersonalcontent.com -> onedrive.live.com
    """
    s = requests.Session()

    # Tạo direct URL
    direct_url = onedrive_to_direct_url(share_or_embed_link, timeout=timeout)

    # Thử GET lần 1
    r = _try_get(direct_url, s, timeout=timeout)
    if r:
        return r.content

    # Fallback: nếu domain là my.microsoftpersonalcontent.com -> đổi sang onedrive.live.com
    pu = urlparse(direct_url)
    if "my.microsoftpersonalcontent.com" in pu.netloc:
        alt = urlunparse((pu.scheme, "onedrive.live.com", pu.path, pu.params, pu.query, pu.fragment))
        r2 = _try_get(alt, s, timeout=timeout)
        if r2:
            return r2.content

    # Thử thêm: nếu là download?resid=.. thiếu cid, thêm cid từ resid
    if "onedrive.live.com" in pu.netloc and "/download" in pu.path:
        q = parse_qs(pu.query)
        resid = (q.get("resid") or [None])[0]
        authkey = (q.get("authkey") or [None])[0]
        cid = (q.get("cid") or [None])[0]
        if resid and authkey and not cid and "!" in resid:
            cid = resid.split("!", 1)[0]
            q["cid"] = [cid]
            alt2 = urlunparse((pu.scheme, pu.netloc, pu.path, "", urlencode({k:v[0] for k,v in q.items()}), ""))
            r3 = _try_get(alt2, s, timeout=timeout)
            if r3:
                return r3.content

    # Hết cách
    raise requests.HTTPError(f"Không tải được file từ OneDrive sau các bước fallback. URL thử: {direct_url}")
