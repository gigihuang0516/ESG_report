import json
import re
import shutil
import tempfile
import time
import zipfile
from http.client import HTTPException
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from uuid import UUID

PAGE = Path(__file__).with_name("index.html").read_bytes()
COMPANY_DATA = re.search(
    rb'<script id="company-data" type="application/json">(.*?)</script>',
    PAGE,
    re.DOTALL,
)
COMPANIES = {
    company["code"]: company["name"]
    for company in json.loads(COMPANY_DATA[1])["companies"]
}
BASE_URL = "https://esggenplus.twse.com.tw"
REPORT_API = BASE_URL + "/api/api/MopsSustainReport/data"
HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Referer": BASE_URL + "/inquiry/report?lang=zh-TW",
}
MIN_YEAR = 2015
MAX_YEAR = 2025
MAX_COMPANIES = 50
MAX_PDF_BYTES = 60 * 1024 * 1024
MAX_ZIP_BYTES = 250 * 1024 * 1024
TIMEOUT = 120


def download_id(value):
    try:
        identifier = UUID(str(value).strip())
    except ValueError:
        return None
    return str(identifier) if identifier.int else None


def query_report(year, code):
    payload = {
        "marketType": 0,
        "year": year,
        "industryNameList": [],
        "companyCodeList": [code],
        "industryName": "all",
        "companyCode": code,
    }
    request = Request(
        REPORT_API,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            **HEADERS,
            "Content-Type": "application/json",
            "Origin": BASE_URL,
        },
    )
    with urlopen(request, timeout=TIMEOUT) as response:
        limit = 4 * 1024 * 1024
        body = response.read(limit + 1)
        if len(body) > limit:
            raise ValueError("查詢回應過大")
        result = json.loads(body)
    if not isinstance(result, dict) or result.get("success") is not True:
        raise ValueError("查詢 API 回傳失敗")
    rows = result.get("data")
    if rows is None:
        rows = []
    if not isinstance(rows, list) or any(not isinstance(row, dict) for row in rows):
        raise ValueError("查詢 API 資料格式錯誤")
    # 同一公司有多筆結果時，優先選有中文修正版的報告。
    for field in ("twEditReportDownloadId", "twFirstReportDownloadId"):
        for row in rows:
            identifier = download_id(row.get(field))
            if identifier:
                return identifier
    return None


def fetch_pdf(identifier, destination, byte_limit):
    request = Request(
        REPORT_API + "/FileStream?" + urlencode({"id": identifier}),
        headers=HEADERS,
    )
    with urlopen(request, timeout=TIMEOUT) as response, destination.open("wb") as file:
        header = response.read(5)
        if not header.startswith(b"%PDF-"):
            raise ValueError("來源未回傳 PDF")
        size = len(header)
        file.write(header)
        while chunk := response.read(64 * 1024):
            size += len(chunk)
            if size > byte_limit:
                raise ValueError("PDF 超過批次檔案大小限制")
            file.write(chunk)
        if size < 1000:
            raise ValueError("PDF 大小異常")
        return size


def create_batch(year, codes, folder):
    results = {}
    total_bytes = 0
    zip_path = folder / f"ESG_{year}.zip"
    pdf_path = folder / "report.pdf"
    with zipfile.ZipFile(zip_path, "w") as archive:
        for index, code in enumerate(codes):
            if index:
                time.sleep(1)
            try:
                identifier = query_report(year, code)
            except (ValueError, OSError, HTTPException) as error:
                print(f"{code} 查詢失敗：{error}", flush=True)
                results[code] = "query_failed"
                continue
            if identifier is None:
                results[code] = "not_found"
                continue
            try:
                size = fetch_pdf(
                    identifier, pdf_path, min(MAX_PDF_BYTES, MAX_ZIP_BYTES - total_bytes)
                )
            except (ValueError, OSError, HTTPException) as error:
                print(f"{code} 下載失敗：{error}", flush=True)
                results[code] = "download_failed"
                continue
            filename = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", f"{code}_{COMPANIES[code]}_{year}")
            archive.write(pdf_path, f"{filename}.pdf")
            total_bytes += size
            results[code] = "downloaded"
    return zip_path, results


def validate_selection(data):
    if not isinstance(data, dict):
        raise ValueError("請提供報告年度與公司代碼。")
    year = data.get("year")
    codes = data.get("company_codes")
    if type(year) is not int or not MIN_YEAR <= year <= MAX_YEAR:
        raise ValueError("請選擇有效的報告年度。")
    if not isinstance(codes, list) or not 1 <= len(codes) <= MAX_COMPANIES:
        raise ValueError(f"請選擇 1～{MAX_COMPANIES} 家公司。")
    if any(not isinstance(code, str) or code not in COMPANIES for code in codes):
        raise ValueError("公司代碼不在下拉選單的清單中。")
    return year, list(dict.fromkeys(codes))


class Handler(BaseHTTPRequestHandler):
    def send_bytes(self, status, content_type, body):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_json(self, status, data):
        self.send_bytes(
            status, "application/json; charset=utf-8",
            json.dumps(data, ensure_ascii=False).encode("utf-8"),
        )

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self.send_bytes(200, "text/html; charset=utf-8", PAGE)
        else:
            self.send_error(404)

    def do_POST(self):
        if self.path != "/api/download":
            self.send_error(404)
            return
        if self.headers.get_content_type() != "application/json":
            self.send_json(415, {"error": "請使用 JSON 傳送資料。"})
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if not 0 < length <= 8192:
                raise ValueError("請求長度不正確。")
            year, codes = validate_selection(json.loads(self.rfile.read(length)))
        except (json.JSONDecodeError, UnicodeError):
            self.send_json(400, {"error": "請求資料格式錯誤。"})
            return
        except ValueError as error:
            self.send_json(400, {"error": str(error)})
            return

        with tempfile.TemporaryDirectory() as directory:
            zip_path, results = create_batch(year, codes, Path(directory))
            if "downloaded" not in results.values():
                self.send_json(200, {"results": results})
                return
            self.send_response(200)
            self.send_header("Content-Type", "application/zip")
            self.send_header("Content-Length", str(zip_path.stat().st_size))
            self.send_header("Content-Disposition", f'attachment; filename="{zip_path.name}"')
            self.send_header("X-ESG-Results", json.dumps(results, separators=(",", ":")))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            with zip_path.open("rb") as file:
                shutil.copyfileobj(file, self.wfile)


if __name__ == "__main__":
    address = ("127.0.0.1", 8000)
    print("請開啟 http://127.0.0.1:8000", flush=True)
    with ThreadingHTTPServer(address, Handler) as server:
        server.serve_forever()
