# app.py
# 和 index.html 放在同一個資料夾，執行 python app.py。
# 本機網址 http://127.0.0.1:8000
# 批次查詢和 ZIP 已用模擬資料驗證。實際證交所下載仍待驗證。

import json
import os
import re
import shutil
import tempfile
import time
from http.client import HTTPException
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from uuid import UUID
from zipfile import ZipFile

PAGE = Path(__file__).with_name("index.html")
BASE_URL = "https://esggenplus.twse.com.tw"
REPORT_API = f"{BASE_URL}/api/api/MopsSustainReport/data"
PAGES_ORIGIN = "https://gigihuang0516.github.io"

MIN_YEAR = 2015
MAX_YEAR = 2025
MAX_COMPANIES = 50
TIMEOUT = 120
MAX_PDF_BYTES = 60 * 1024 * 1024
MAX_BATCH_BYTES = 250 * 1024 * 1024

UPSTREAM_HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Referer": f"{BASE_URL}/inquiry/report?lang=zh-TW",
}


def valid_download_id(value):
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
            **UPSTREAM_HEADERS,
            "Accept": "application/json",
            "Content-Type": "application/json",
            "Origin": BASE_URL,
        },
        method="POST",
    )

    with urlopen(request, timeout=TIMEOUT) as response:
        limit = 4 * 1024 * 1024
        body = response.read(limit + 1)
        if len(body) > limit:
            raise ValueError("查詢回應過大。")
        result = json.loads(body)

    if not isinstance(result, dict) or result.get("success") is not True:
        raise ValueError("證交所查詢未成功。")

    rows = result.get("data")
    if rows is None:
        rows = []
    if not isinstance(rows, list) or any(not isinstance(row, dict) for row in rows):
        raise ValueError("證交所回傳的報告清單格式有誤。")

    matching_rows = []
    for row in rows:
        raw_code = row.get("code") or row.get("companyCode") or ""
        returned_code = str(raw_code).strip().removesuffix(".0")
        if returned_code == code:
            matching_rows.append(row)

    for field in ("twEditReportDownloadId", "twFirstReportDownloadId"):
        for row in matching_rows:
            identifier = valid_download_id(row.get(field))
            if identifier:
                company_name = row.get("shortName") or row.get("name") or ""
                return identifier, str(company_name).strip()

    return None


def download_pdf(identifier, destination, max_bytes):
    request = Request(
        f"{REPORT_API}/FileStream?{urlencode({'id': identifier})}",
        headers=UPSTREAM_HEADERS,
    )

    with urlopen(request, timeout=TIMEOUT) as response:
        header = response.read(5)
        if header != b"%PDF-":
            raise ValueError("來源未回傳 PDF。")
        if len(header) > max_bytes:
            raise ValueError("超過批次檔案大小限制。")

        with destination.open("wb") as file:
            file.write(header)
            size = len(header)

            while chunk := response.read(64 * 1024):
                size += len(chunk)
                if size > max_bytes:
                    raise ValueError("超過 PDF 或批次檔案大小限制。")
                file.write(chunk)

    if size < 1000:
        raise ValueError("PDF 大小異常。")

    return size


def create_batch(year, codes, directory):
    zip_path = directory / f"ESG_{year}.zip"
    pdf_path = directory / "report.pdf"
    results = {}
    total_bytes = 0

    with ZipFile(zip_path, "w") as archive:
        for index, code in enumerate(codes):
            if index:
                time.sleep(1)

            try:
                report = query_report(year, code)
            except (OSError, ValueError, HTTPException) as error:
                print(f"{code} 查詢失敗：{error}", flush=True)
                results[code] = "query_failed"
                continue

            if report is None:
                results[code] = "not_found"
                continue

            identifier, name = report
            try:
                byte_limit = min(MAX_PDF_BYTES, MAX_BATCH_BYTES - total_bytes)
                size = download_pdf(identifier, pdf_path, byte_limit)

                parts = (code, name[:80], str(year))
                stem = "_".join(part for part in parts if part)
                filename = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", stem)
                archive.write(pdf_path, f"{filename}.pdf")
            except (OSError, ValueError, HTTPException) as error:
                print(f"{code} 下載失敗：{error}", flush=True)
                results[code] = "download_failed"
                continue

            total_bytes += size
            results[code] = "downloaded"

    return zip_path, results


def validate_selection(data):
    if not isinstance(data, dict):
        raise ValueError("請提供年度與公司代碼。")

    year = data.get("year")
    codes = data.get("company_codes")

    if type(year) is not int or not MIN_YEAR <= year <= MAX_YEAR:
        raise ValueError("請選擇有效的報告年度。")

    if not isinstance(codes, list) or not codes:
        raise ValueError("請輸入至少一個上市公司代碼。")

    if any(not isinstance(code, str) or not re.fullmatch(r"[0-9]{4}", code) for code in codes):
        raise ValueError("公司代碼須為四位數字。")

    unique_codes = list(dict.fromkeys(codes))
    if len(unique_codes) > MAX_COMPANIES:
        raise ValueError(f"每次最多處理 {MAX_COMPANIES} 家公司。")

    return year, unique_codes


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
        self.end_headers()
        self.wfile.write(body)

    def send_json(self, status, data):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_bytes(status, "application/json; charset=utf-8", body)

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self.send_bytes(200, "text/html; charset=utf-8", PAGE.read_bytes())
        else:
            self.send_error(404)

    def do_OPTIONS(self):
        if self.path != "/api/download" or self.headers.get("Origin") != PAGES_ORIGIN:
            self.send_error(403)
            return

        self.send_response(204)
        self.send_header("Access-Control-Allow-Methods", "POST")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_POST(self):
        if self.path != "/api/download":
            self.send_json(404, {"error": "找不到下載功能。"})
            return

        if self.headers.get_content_type() != "application/json":
            self.send_json(415, {"error": "請使用 JSON 傳送下載條件。"})
            return

        try:
            length = int(self.headers.get("Content-Length", "0"))
            if not 0 < length <= 8192:
                raise ValueError("下載條件的資料長度不正確。")

            data = json.loads(self.rfile.read(length))
            year, codes = validate_selection(data)
        except (json.JSONDecodeError, UnicodeError):
            self.send_json(400, {"error": "下載條件的格式不正確。"})
            return
        except ValueError as error:
            self.send_json(400, {"error": str(error)})
            return

        with tempfile.TemporaryDirectory() as folder:
            zip_path, results = create_batch(year, codes, Path(folder))

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
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))

    with ThreadingHTTPServer((host, port), ReportHandler) as server:
        print(f"請開啟 http://127.0.0.1:{port}", flush=True)
        server.serve_forever()
