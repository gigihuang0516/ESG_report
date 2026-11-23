# app.py
# 與 index.html 放在同一個資料夾。
# 安裝 Excel 套件：python -m pip install openpyxl
# 啟動：python app.py
# 開啟：http://127.0.0.1:8000
# 此改版的 Excel 上傳及實際證交所下載仍待執行驗證。

import base64
import binascii
import io
import json
import os
import re
import shutil
import tempfile
import time
from http import HTTPStatus
from http.client import HTTPException
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from unicodedata import normalize
from urllib.parse import urlencode, urlsplit
from urllib.request import Request, urlopen
from uuid import UUID
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile, ZipFile

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException


PAGE = Path(__file__).with_name("index.html")
BASE_URL = "https://esggenplus.twse.com.tw"
REPORT_API = f"{BASE_URL}/api/api/MopsSustainReport/data"
PAGES_ORIGIN = "https://gigihuang0516.github.io"

MIN_YEAR = 2015
MAX_YEAR = 2025
MAX_COMPANIES = 50
TIMEOUT = 120
MAX_EXCEL_BYTES = 2 * 1024 * 1024
MAX_EXCEL_EXPANDED_BYTES = 16 * 1024 * 1024
MAX_REQUEST_BYTES = 4 * ((MAX_EXCEL_BYTES + 2) // 3) + 8192
MAX_QUERY_BYTES = 4 * 1024 * 1024
MAX_PDF_BYTES = 60 * 1024 * 1024
MAX_BATCH_BYTES = 250 * 1024 * 1024

COMPANY_CODE = re.compile(r"[0-9]{4}")
INVALID_FILENAME = re.compile(r'[\\/:*?"<>|\x00-\x1f]')

UPSTREAM_HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Referer": f"{BASE_URL}/inquiry/report?lang=zh-TW",
}


def normalize_code(value):
    if value is None:
        return ""
    return normalize("NFKC", str(value)).strip().removesuffix(".0")


def read_excel_companies(encoded):
    if not isinstance(encoded, str) or not encoded:
        raise ValueError("請上傳 .xlsx 公司清單。")

    try:
        content = base64.b64decode(encoded, validate=True)
    except (ValueError, binascii.Error) as error:
        raise ValueError("Excel 檔案資料格式不正確。") from error

    if not 0 < len(content) <= MAX_EXCEL_BYTES:
        raise ValueError("請上傳 2 MB 以內的 .xlsx 公司清單。")

    try:
        with ZipFile(io.BytesIO(content)) as archive:
            expanded_size = sum(item.file_size for item in archive.infolist())
    except BadZipFile as error:
        raise ValueError("檔案不是有效的 .xlsx 公司清單。") from error

    if expanded_size > MAX_EXCEL_EXPANDED_BYTES:
        raise ValueError("Excel 解壓縮後的資料過大。")

    workbook = None
    try:
        workbook = load_workbook(
            io.BytesIO(content), read_only=True, data_only=True, keep_links=False
        )
        if "公司清單" not in workbook.sheetnames:
            raise ValueError("Excel 找不到「公司清單」工作表。")

        sheet = workbook["公司清單"]
        # 串流讀取實際儲存格，避免 Excel 的格式設定擴大讀取範圍。
        sheet.reset_dimensions()
        rows = sheet.iter_rows(values_only=True)
        headers = [str(value or "").strip() for value in next(rows, ())]
        if "證券代碼" not in headers or "公司名稱" not in headers:
            raise ValueError("「公司清單」第一列須包含「證券代碼」及「公司名稱」。")

        code_index = headers.index("證券代碼")
        name_index = headers.index("公司名稱")
        companies = {}

        for row_number, row in enumerate(rows, start=2):
            code = normalize_code(row[code_index]) if code_index < len(row) else ""
            if not code:
                continue
            if not COMPANY_CODE.fullmatch(code):
                raise ValueError(f"Excel 第 {row_number} 列的證券代碼須為四位數字。")

            value = row[name_index] if name_index < len(row) else None
            name = str(value or "").strip()[:80]
            companies[code] = companies.get(code) or name
            if len(companies) > MAX_COMPANIES:
                raise ValueError(f"每次最多處理 {MAX_COMPANIES} 家公司。")

        if not companies:
            raise ValueError("「公司清單」沒有有效的證券代碼。")
        return companies

    except (BadZipFile, InvalidFileException, KeyError, OSError, ParseError) as error:
        raise ValueError("無法讀取這份 Excel 公司清單。") from error
    finally:
        if workbook is not None:
            workbook.close()


def validate_selection(data):
    if not isinstance(data, dict):
        raise ValueError("請提供報告年度與公司清單。")

    year = data.get("year")
    if type(year) is not int or not MIN_YEAR <= year <= MAX_YEAR:
        raise ValueError("請選擇有效的報告年度。")

    if "company_file" in data:
        if "company_codes" in data:
            raise ValueError("請選擇一種公司清單來源。")
        return year, read_excel_companies(data["company_file"])

    codes = data.get("company_codes")
    if not isinstance(codes, list) or not codes:
        raise ValueError("請輸入至少一個上市公司代碼。")

    companies = {}
    for raw_code in codes:
        if not isinstance(raw_code, str):
            raise ValueError("公司代碼須為四位數字。")
        code = normalize("NFKC", raw_code).strip()
        if not COMPANY_CODE.fullmatch(code):
            raise ValueError("公司代碼須為四位數字。")
        companies[code] = ""

    if len(companies) > MAX_COMPANIES:
        raise ValueError(f"每次最多處理 {MAX_COMPANIES} 家公司。")
    return year, companies


def valid_download_id(value):
    try:
        identifier = UUID(str(value).strip())
    except ValueError:
        return None
    return str(identifier) if identifier.int else None


def query_report(year, code):
    # 沿用原版的單一公司查詢方式，marketType=0 表示上市。
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
            **UPSTREAM_HEADERS,
            "Accept": "application/json",
            "Content-Type": "application/json",
            "Origin": BASE_URL,
        },
        method="POST",
    )

    with urlopen(request, timeout=TIMEOUT) as response:
        content = response.read(MAX_QUERY_BYTES + 1)

    if len(content) > MAX_QUERY_BYTES:
        raise ValueError("證交所查詢回應過大。")
    try:
        result = json.loads(content)
    except (ValueError, UnicodeError) as error:
        raise ValueError("證交所未回傳有效的 JSON 查詢結果。") from error

    if not isinstance(result, dict) or result.get("success") is not True:
        raise ValueError("證交所查詢未成功。")

    rows = result.get("data")
    if rows is None:
        return None
    if not isinstance(rows, list) or any(not isinstance(row, dict) for row in rows):
        raise ValueError("證交所的報告清單格式有誤。")

    matching_rows = [
        row for row in rows
        if code in (normalize_code(row.get("code")), normalize_code(row.get("companyCode")))
    ]
    if rows and not matching_rows:
        raise ValueError("證交所的查詢結果無法核對公司代碼。")

    # 只採用中文 PDF，有修正版時優先下載。
    for field in ("twEditReportDownloadId", "twFirstReportDownloadId"):
        for row in matching_rows:
            identifier = valid_download_id(row.get(field))
            if identifier:
                name = row.get("shortName") or row.get("name") or ""
                return identifier, str(name).strip()[:80]
    return None


def download_pdf(identifier, destination, byte_limit):
    if byte_limit < 1000:
        raise ValueError("已達批次檔案大小限制。")

    query = urlencode({"id": identifier})
    request = Request(f"{REPORT_API}/FileStream?{query}", headers=UPSTREAM_HEADERS)

    with urlopen(request, timeout=TIMEOUT) as response:
        signature = response.read(5)
        if signature != b"%PDF-":
            raise ValueError("證交所未回傳 PDF。")

        size = len(signature)
        with destination.open("wb") as file:
            file.write(signature)
            while chunk := response.read(64 * 1024):
                size += len(chunk)
                if size > byte_limit:
                    raise ValueError("超過 PDF 或批次檔案大小限制。")
                file.write(chunk)

    if size < 1000:
        raise ValueError("PDF 大小異常。")
    return size


def create_batch(year, companies, directory):
    zip_path = directory / f"ESG_{year}.zip"
    pdf_path = directory / "report.pdf"
    results = []
    total_bytes = 0

    with ZipFile(zip_path, "w") as archive:
        for index, (code, excel_name) in enumerate(companies.items()):
            if index:
                time.sleep(1)

            result = {"code": code, "name": excel_name, "status": "query_failed"}
            results.append(result)

            try:
                report = query_report(year, code)
            except (OSError, ValueError, HTTPException) as error:
                print(f"{code} 查詢失敗：{error}", flush=True)
                continue

            if report is None:
                result["status"] = "not_found"
                continue

            identifier, api_name = report
            result["name"] = excel_name or api_name
            try:
                limit = min(MAX_PDF_BYTES, MAX_BATCH_BYTES - total_bytes)
                size = download_pdf(identifier, pdf_path, limit)
            except (OSError, ValueError, HTTPException) as error:
                print(f"{code} 下載失敗：{error}", flush=True)
                result["status"] = "download_failed"
                continue

            parts = (code, result["name"], str(year))
            stem = "_".join(part for part in parts if part)
            filename = INVALID_FILENAME.sub("_", stem)
            archive.write(pdf_path, f"{filename}.pdf")
            total_bytes += size
            result["status"] = "downloaded"

    return zip_path, results


class ReportHandler(BaseHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Vary", "Origin")
        if self.headers.get("Origin") == PAGES_ORIGIN:
            self.send_header("Access-Control-Allow-Origin", PAGES_ORIGIN)
            self.send_header("Access-Control-Expose-Headers", "X-ESG-Results")
        super().end_headers()

    def send_bytes(self, status, content_type, body):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def send_json(self, status, data):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_bytes(status, "application/json; charset=utf-8", body)

    def do_GET(self):
        if urlsplit(self.path).path in ("/", "/index.html"):
            self.send_bytes(HTTPStatus.OK, "text/html; charset=utf-8", PAGE.read_bytes())
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_OPTIONS(self):
        if self.path != "/api/download" or self.headers.get("Origin") != PAGES_ORIGIN:
            self.send_error(HTTPStatus.FORBIDDEN)
            return

        self.send_response(HTTPStatus.NO_CONTENT)
        self.send_header("Access-Control-Allow-Methods", "POST")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_POST(self):
        if self.path != "/api/download":
            self.send_json(HTTPStatus.NOT_FOUND, {"error": "找不到下載功能。"})
            return

        if self.headers.get_content_type() != "application/json":
            self.send_json(
                HTTPStatus.UNSUPPORTED_MEDIA_TYPE,
                {"error": "請使用 JSON 傳送下載條件。"},
            )
            return

        try:
            length = int(self.headers.get("Content-Length", "0"))
            if not 0 < length <= MAX_REQUEST_BYTES:
                raise ValueError("下載條件的資料大小超出限制。")
            data = json.loads(self.rfile.read(length))
            year, companies = validate_selection(data)
        except (json.JSONDecodeError, UnicodeError):
            self.send_json(HTTPStatus.BAD_REQUEST, {"error": "下載條件的格式不正確。"})
            return
        except ValueError as error:
            self.send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
            return

        with tempfile.TemporaryDirectory() as directory:
            try:
                zip_path, results = create_batch(year, companies, Path(directory))
            except OSError as error:
                print(f"ZIP 打包失敗：{error}", flush=True)
                self.send_json(
                    HTTPStatus.INTERNAL_SERVER_ERROR,
                    {"error": "暫時無法產生 ZIP，請稍後重試。"},
                )
                return

            if not any(item["status"] == "downloaded" for item in results):
                self.send_json(HTTPStatus.OK, {"results": results})
                return

            # 結果放在標頭中。公司名稱以 ASCII 安全的 JSON 編碼傳送。
            summary = json.dumps(results, ensure_ascii=True, separators=(",", ":"))
            # 供代理伺服器轉送 ZIP 時控制標頭大小，完整公司名稱已保留在 PDF 檔名。
            if len(summary) > 6000:
                summary = json.dumps(
                    [{"code": item["code"], "status": item["status"]} for item in results],
                    separators=(",", ":"),
                )

            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/zip")
            self.send_header("Content-Length", str(zip_path.stat().st_size))
            self.send_header("Content-Disposition", f'attachment; filename="{zip_path.name}"')
            self.send_header("X-ESG-Results", summary)
            self.send_header("Cache-Control", "no-store")
            self.end_headers()

            with zip_path.open("rb") as file:
                shutil.copyfileobj(file, self.wfile)


if __name__ == "__main__":
    # 部署時可由平台設定 HOST=0.0.0.0 與 PORT。
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))

    with ThreadingHTTPServer((host, port), ReportHandler) as server:
        print(f"永續報告書下載服務已啟動：http://127.0.0.1:{port}", flush=True)
        server.serve_forever()
