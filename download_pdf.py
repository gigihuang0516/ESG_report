import requests
from query_reports import has_valid_uuid, normalize_cell

FILE_URL = "https://esggenplus.twse.com.tw/api/api/MopsSustainReport/data/FileStream"
TIMEOUT_SECONDS = 120
MAX_PDF_BYTES = 60 * 1024 * 1024

def new_session():
    session = requests.Session()
    session.headers.update({"User-Agent":"Mozilla/5.0", "Referer":"https://esggenplus.twse.com.tw/inquiry/report?lang=zh-TW"})
    return session

def choose_version(row, prefer_edit=True):
    first, edit = row.get('twFirstReportDownloadId'), row.get('twEditReportDownloadId')
    if prefer_edit and has_valid_uuid(edit): return normalize_cell(edit), '修正版'
    if has_valid_uuid(first): return normalize_cell(first), '初次申報版'
    if has_valid_uuid(edit): return normalize_cell(edit), '修正版（初次版無效）'
    raise ValueError('沒有有效 UUID')

def download_pdf(session, download_id, log):
    # Original non-streaming request; size cap is checked after reception.
    with session.get(FILE_URL, params={"id":download_id}, timeout=TIMEOUT_SECONDS) as response:
        log['HTTP狀態碼'] = response.status_code
        log['Content-Type'] = response.headers.get('Content-Type','')
        response.raise_for_status()
        content = response.content
        log['檔案大小(bytes)'] = len(content)
        if not content.startswith(b'%PDF-'): raise ValueError('回傳內容不是 PDF（缺少 %PDF- 標頭）')
        if len(content) < 1000: raise ValueError('檔案大小異常，低於 1,000 bytes')
        if len(content) > MAX_PDF_BYTES: raise ValueError('單份 PDF 超過 60 MB 上限')
        return content
