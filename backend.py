"""ESG report query/download logic adapted from the user's existing scripts."""
import csv
import io
import re
import time
import zipfile
from datetime import datetime
from pathlib import Path
from uuid import UUID

import requests
from openpyxl import load_workbook

API_URL = 'https://esggenplus.twse.com.tw/api/api/MopsSustainReport/data'
FILE_URL = API_URL + '/FileStream'
HEADERS = {'User-Agent': 'Mozilla/5.0', 'Origin': 'https://esggenplus.twse.com.tw', 'Referer': 'https://esggenplus.twse.com.tw/inquiry/report?lang=zh-TW'}
LOG_FIELDS = ['證券代碼', 'Excel公司名稱', 'HTTP狀態碼', 'API成功', 'API count', '實際回傳筆數', '查詢結果', '訊息']
DL_FIELDS = ['公司代號', '公司名稱', '年度', '版本', 'HTTP狀態碼', 'Content-Type', '檔案大小(bytes)', '結果', '檔案路徑']


def clean(value):
    if value is None or isinstance(value, bool):
        return ''
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip().replace('\u3000', '')
    return text[:-2] if text.endswith('.0') and text[:-2].isdigit() else text


def valid_uuid(value):
    try:
        parsed = UUID(clean(value))
        return parsed.int != 0 and str(parsed) == clean(value).lower()
    except (ValueError, AttributeError, TypeError):
        return False


def load_companies(excel_bytes):
    workbook = load_workbook(io.BytesIO(excel_bytes), read_only=True, data_only=True)
    try:
        if '公司清單' not in workbook.sheetnames:
            raise ValueError('Excel 必須有「公司清單」工作表。')
        sheet = workbook['公司清單']
        headers = [clean(cell.value) for cell in sheet[1]]
        if '證券代碼' not in headers or '公司名稱' not in headers:
            raise ValueError('第一列必須包含「證券代碼」及「公司名稱」。')
        ci, ni = headers.index('證券代碼'), headers.index('公司名稱')
        companies, seen = [], set()
        for values in sheet.iter_rows(min_row=2, values_only=True):
            code = clean(values[ci] if ci < len(values) else None)
            name = clean(values[ni] if ni < len(values) else None)
            if code and code not in seen:
                seen.add(code)
                companies.append({'code': code, 'name': name})
        if not companies:
            raise ValueError('公司清單沒有有效的證券代碼。')
        return companies
    finally:
        workbook.close()


def query_reports(companies, year, progress=None, delay=0.5):
    session = requests.Session()
    session.headers.update({**HEADERS, 'Content-Type': 'application/json'})
    all_rows, logs = [], []
    try:
        for index, company in enumerate(companies, 1):
            code, name = company['code'], company['name']
            payload = {'marketType': 0, 'year': int(year), 'industryNameList': [], 'companyCodeList': [code], 'industryName': 'all', 'companyCode': code}
            status, count, returned, message, ok, state = '', '', 0, '', False, ''
            try:
                response = session.post(API_URL, json=payload, timeout=(10, 30))
                status = response.status_code
                response.raise_for_status()
                result = response.json()
                if not isinstance(result, dict):
                    raise ValueError('API 回傳格式錯誤。')
                ok = result.get('success') is True
                count = result.get('count', '')
                message = str(result.get('message') or '')
                if ok:
                    data = result.get('data') or []
                    if not isinstance(data, list):
                        raise ValueError('API data 不是清單。')
                    returned = len(data)
                    for item in data:
                        if isinstance(item, dict):
                            all_rows.append({**item, 'code': code, 'inputCompanyName': name, 'reportYear': str(year)})
                    state = '成功（有資料）' if returned else '成功（無資料）'
                else:
                    state = 'API 回傳失敗'
                    message = message or 'API success 欄位不是 True'
            except Exception as exc:
                state, message = '查詢錯誤', str(exc)
            logs.append({'證券代碼': code, 'Excel公司名稱': name, 'HTTP狀態碼': status, 'API成功': '是' if ok else '否', 'API count': count, '實際回傳筆數': returned, '查詢結果': state, '訊息': message})
            if progress:
                progress(index, len(companies), f'查詢 {code} {name}：{state}')
            if index < len(companies) and delay:
                time.sleep(delay)
    finally:
        session.close()
    selected = [row for row in all_rows if valid_uuid(row.get('twEditReportDownloadId')) or valid_uuid(row.get('twFirstReportDownloadId'))]
    return selected, logs


def safe_filename(value):
    return re.sub(r'[\\/:*?"<>|\x00-\x1f]', '_', str(value)).strip().rstrip('. ') or '未命名公司'


def csv_bytes(rows, fields):
    output = io.StringIO(newline='')
    writer = csv.DictWriter(output, fieldnames=fields, extrasaction='ignore')
    writer.writeheader()
    writer.writerows(rows)
    return ('\ufeff' + output.getvalue()).encode('utf-8')


def make_zip(companies, rows, query_logs, year, progress=None, delay=1):
    company_map = {c['code']: c['name'] for c in companies}
    logs = []
    output = io.BytesIO()
    session = requests.Session()
    session.headers.update(HEADERS)
    success = 0
    try:
        with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_DEFLATED) as archive:
            for index, row in enumerate(rows, 1):
                code = clean(row.get('code'))
                name = company_map.get(code) or clean(row.get('shortName') or row.get('name') or row.get('inputCompanyName'))
                report_year = clean(row.get('reportYear')) or str(year)
                edit_id, first_id = row.get('twEditReportDownloadId'), row.get('twFirstReportDownloadId')
                if valid_uuid(edit_id):
                    download_id, version = clean(edit_id), '修正版'
                elif valid_uuid(first_id):
                    download_id, version = clean(first_id), '初次申報版'
                else:
                    continue
                filename = safe_filename(f'{code}_{name}_{report_year}.pdf')
                status, content_type, size, result_text, filepath = '', '', 0, '', ''
                try:
                    # stream=True avoids buffering a full response before validating size.
                    with session.get(FILE_URL, params={'id': download_id}, timeout=(10, 30), stream=True) as response:
                        status = response.status_code
                        content_type = response.headers.get('Content-Type', '')
                        response.raise_for_status()
                        max_bytes = 60 * 1024 * 1024  # per-PDF safety cap for public hosting
                        chunks = []
                        for chunk in response.iter_content(chunk_size=1024 * 256):
                            if chunk:
                                size += len(chunk)
                                if size > max_bytes:
                                    raise ValueError('單份 PDF 超過 60 MB 上限')
                                chunks.append(chunk)
                        content = b''.join(chunks)
                    if not content.startswith(b'%PDF-') or len(content) < 1000:
                        raise ValueError('回傳內容不是有效 PDF')
                    archive.writestr(f'PDF/{filename}', content)
                    success += 1
                    filepath = f'PDF/{filename}'
                    result_text = '成功'
                except Exception as exc:
                    result_text = f'失敗：{exc}'
                logs.append({'公司代號': code, '公司名稱': name, '年度': report_year, '版本': version, 'HTTP狀態碼': status, 'Content-Type': content_type, '檔案大小(bytes)': size, '結果': result_text, '檔案路徑': filepath})
                if progress:
                    progress(index, len(rows), f'下載 {code} {name}：{result_text}')
                if index < len(rows) and delay:
                    time.sleep(delay)
            fields = list(dict.fromkeys(k for row in rows for k in row))
            fields += [key for key in ['code', 'inputCompanyName', 'reportYear', 'twFirstReportDownloadId', 'twEditReportDownloadId'] if key not in fields]
            archive.writestr(f'{year}_篩選結果.csv', csv_bytes(rows, fields))
            archive.writestr(f'{year}_查詢紀錄.csv', csv_bytes(query_logs, LOG_FIELDS))
            archive.writestr('download_log.csv', csv_bytes(logs, DL_FIELDS))
    finally:
        session.close()
    return output.getvalue(), logs, success
