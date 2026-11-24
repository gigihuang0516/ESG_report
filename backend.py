
"""ESG 永續報告書查詢與下載模組"""

import csv
import io
import re
import time
import zipfile
from uuid import UUID

import requests
from openpyxl import load_workbook


# =========================
# 基本設定
# =========================

API_URL = 'https://esggenplus.twse.com.tw/api/api/MopsSustainReport/data'
FILE_URL = API_URL + '/FileStream'

HEADERS = {
    'User-Agent': 'Mozilla/5.0',
    'Origin': 'https://esggenplus.twse.com.tw',
    'Referer': 'https://esggenplus.twse.com.tw/inquiry/report?lang=zh-TW'
}

TIMEOUT = (10, 30)
MAX_PDF_BYTES = 60 * 1024 * 1024

LOG_FIELDS = [
    '證券代碼',
    'Excel公司名稱',
    'HTTP狀態碼',
    'API成功',
    'API count',
    '實際回傳筆數',
    '查詢結果',
    '訊息',
    '耗時(秒)'
]

DL_FIELDS = [
    '公司代號',
    '公司名稱',
    '年度',
    '版本',
    'HTTP狀態碼',
    'Content-Type',
    '檔案大小(bytes)',
    '結果',
    '檔案路徑',
    '耗時(秒)'
]


# =========================
# 共用工具
# =========================

def clean(value):
    if value is None or isinstance(value, bool):
        return ''

    if isinstance(value, int):
        return str(value)

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    text = str(value).strip().replace('\u3000', '')

    return (
        text[:-2]
        if text.endswith('.0') and text[:-2].isdigit()
        else text
    )


def valid_uuid(value):
    try:
        parsed = UUID(clean(value))

        return (
            parsed.int != 0
            and str(parsed) == clean(value).lower()
        )

    except (ValueError, AttributeError, TypeError):
        return False


def notify(progress, index, total, message):
    print(message, flush=True)

    if progress:
        progress(index, total, message)


def safe_filename(value):
    return (
        re.sub(
            r'[\\/:*?"<>|\x00-\x1f]',
            '_',
            str(value)
        ).strip().rstrip('. ')
        or '未命名公司'
    )


def csv_bytes(rows, fields):
    output = io.StringIO(newline='')

    writer = csv.DictWriter(
        output,
        fieldnames=fields,
        extrasaction='ignore'
    )

    writer.writeheader()
    writer.writerows(rows)

    return ('\ufeff' + output.getvalue()).encode('utf-8')


# =========================
# 讀取 Excel 公司清單
# =========================

def load_companies(excel_bytes):

    workbook = load_workbook(
        io.BytesIO(excel_bytes),
        read_only=True,
        data_only=True
    )

    try:
        if '公司清單' not in workbook.sheetnames:
            raise ValueError(
                'Excel 必須有「公司清單」工作表。'
            )

        sheet = workbook['公司清單']

        headers = [
            clean(cell.value)
            for cell in sheet[1]
        ]

        if '證券代碼' not in headers or '公司名稱' not in headers:
            raise ValueError(
                '第一列必須包含「證券代碼」及「公司名稱」。'
            )

        ci = headers.index('證券代碼')
        ni = headers.index('公司名稱')

        companies = []
        seen = set()

        for values in sheet.iter_rows(
            min_row=2,
            values_only=True
        ):

            code = clean(
                values[ci] if ci < len(values) else None
            )

            name = clean(
                values[ni] if ni < len(values) else None
            )

            if code and code not in seen:
                seen.add(code)

                companies.append({
                    'code': code,
                    'name': name
                })

        if not companies:
            raise ValueError(
                '公司清單沒有有效的證券代碼。'
            )

        return companies

    finally:
        workbook.close()


# =========================
# API 查詢及報告書篩選
# =========================

def query_reports(
    companies,
    year,
    progress=None,
    delay=0.5
):

    session = requests.Session()

    session.headers.update({
        **HEADERS,
        'Content-Type': 'application/json'
    })

    all_rows = []
    logs = []

    try:

        for index, company in enumerate(companies, 1):

            code = company['code']
            name = company['name']

            start = time.perf_counter()

            payload = {
                'marketType': 0,
                'year': int(year),
                'industryNameList': [],
                'companyCodeList': [code],
                'industryName': 'all',
                'companyCode': code
            }

            status = ''
            count = ''
            returned = 0
            message = ''
            ok = False
            state = ''

            notify(
                progress,
                index - 1,
                len(companies),
                f'正在連線 TWSE API：{code} {name}'
            )

            try:

                response = session.post(
                    API_URL,
                    json=payload,
                    timeout=TIMEOUT
                )

                status = response.status_code

                response.raise_for_status()

                result = response.json()

                if not isinstance(result, dict):
                    raise ValueError(
                        'API 回傳格式錯誤。'
                    )

                count = result.get('count', '')
                message = str(
                    result.get('message') or ''
                )

                if result.get('success') is True:

                    data = result.get('data') or []

                    if not isinstance(data, list):
                        raise ValueError(
                            'API data 不是清單。'
                        )

                    returned = len(data)

                    for item in data:

                        if isinstance(item, dict):

                            all_rows.append({
                                **item,
                                'code': code,
                                'inputCompanyName': name,
                                'reportYear': str(year)
                            })

                    ok = True

                    state = (
                        '成功（有資料）'
                        if returned
                        else '成功（無資料）'
                    )

                else:

                    state = 'API 回傳失敗'

                    message = (
                        message
                        or 'API success 欄位不是 True'
                    )

            except requests.exceptions.Timeout as exc:

                state = 'API 連線或讀取逾時'
                message = str(exc)

            except Exception as exc:

                state = '查詢錯誤'
                message = str(exc)

            elapsed = round(
                time.perf_counter() - start, 2
            )

            logs.append({
                '證券代碼': code,
                'Excel公司名稱': name,
                'HTTP狀態碼': status,
                'API成功': '是' if ok else '否',
                'API count': count,
                '實際回傳筆數': returned,
                '查詢結果': state,
                '訊息': message,
                '耗時(秒)': elapsed
            })

            notify(
                progress,
                index,
                len(companies),
                f'查詢 {code} {name}：{state}，'
                f'耗時 {elapsed} 秒'
            )

            if message and not ok:
                print(message, flush=True)

            if index < len(companies) and delay:
                time.sleep(delay)

    finally:
        session.close()

    selected = [
        row for row in all_rows
        if valid_uuid(
            row.get('twEditReportDownloadId')
        )
        or valid_uuid(
            row.get('twFirstReportDownloadId')
        )
    ]

    print(
        f'查詢完成，找到 {len(selected)} 份可下載報告。',
        flush=True
    )

    return selected, logs


# =========================
# PDF 下載及 ZIP 打包
# =========================

def make_zip(
    companies,
    rows,
    query_logs,
    year,
    progress=None,
    delay=1
):

    company_map = {
        c['code']: c['name']
        for c in companies
    }

    logs = []
    output = io.BytesIO()
    success = 0

    session = requests.Session()
    session.headers.update(HEADERS)

    try:

        with zipfile.ZipFile(
            output,
            'w',
            compression=zipfile.ZIP_DEFLATED
        ) as archive:

            for index, row in enumerate(rows, 1):

                start = time.perf_counter()

                code = clean(row.get('code'))

                name = (
                    company_map.get(code)
                    or clean(
                        row.get('shortName')
                        or row.get('name')
                        or row.get('inputCompanyName')
                    )
                )

                report_year = (
                    clean(row.get('reportYear'))
                    or str(year)
                )

                edit_id = row.get(
                    'twEditReportDownloadId'
                )

                first_id = row.get(
                    'twFirstReportDownloadId'
                )

                if valid_uuid(edit_id):

                    download_id = clean(edit_id)
                    version = '修正版'

                elif valid_uuid(first_id):

                    download_id = clean(first_id)
                    version = '初次申報版'

                else:
                    continue

                filename = safe_filename(
                    f'{code}_{name}_{report_year}.pdf'
                )

                status = ''
                content_type = ''
                size = 0
                result_text = ''
                filepath = ''

                notify(
                    progress,
                    index - 1,
                    len(rows),
                    f'正在連線 PDF 下載端點：{code} {name}'
                )

                try:

                    with session.get(
                        FILE_URL,
                        params={'id': download_id},
                        timeout=TIMEOUT,
                        stream=True
                    ) as response:

                        status = response.status_code

                        content_type = response.headers.get(
                            'Content-Type', ''
                        )

                        response.raise_for_status()

                        try:
                            total_bytes = int(
                                response.headers.get(
                                    'Content-Length', 0
                                )
                            )
                        except (ValueError, TypeError):
                            total_bytes = 0

                        if total_bytes > MAX_PDF_BYTES:
                            raise ValueError(
                                '單份 PDF 超過 60 MB 上限'
                            )

                        buffer = io.BytesIO()
                        last_update = 0

                        for chunk in response.iter_content(
                            chunk_size=256 * 1024
                        ):

                            if not chunk:
                                continue

                            size += len(chunk)

                            if size > MAX_PDF_BYTES:
                                raise ValueError(
                                    '單份 PDF 超過 60 MB 上限'
                                )

                            buffer.write(chunk)

                            now = time.monotonic()

                            if now - last_update >= 0.5:

                                downloaded_mb = (
                                    size / 1024 / 1024
                                )

                                if total_bytes > 0:

                                    total_mb = (
                                        total_bytes
                                        / 1024
                                        / 1024
                                    )

                                    fraction = min(
                                        size / total_bytes,
                                        1.0
                                    )

                                    msg = (
                                        f'正在下載 {code} {name}：'
                                        f'{downloaded_mb:.2f} / '
                                        f'{total_mb:.2f} MB'
                                    )

                                else:

                                    fraction = 0

                                    msg = (
                                        f'正在下載 {code} {name}：'
                                        f'已下載 {downloaded_mb:.2f} MB'
                                    )

                                if progress:
                                    progress(
                                        index - 1 + fraction,
                                        len(rows),
                                        msg
                                    )

                                last_update = now

                        content = buffer.getvalue()
                        buffer.close()

                    if (
                        not content.startswith(b'%PDF-')
                        or len(content) < 1000
                    ):
                        raise ValueError(
                            '回傳內容不是有效 PDF'
                        )

                    notify(
                        progress,
                        index - 1,
                        len(rows),
                        f'正在將 {code} {name} 加入 ZIP'
                    )

                    archive.writestr(
                        f'PDF/{filename}',
                        content,
                        compress_type=zipfile.ZIP_STORED
                    )

                    success += 1

                    filepath = f'PDF/{filename}'
                    result_text = '成功'

                    del content

                except requests.exceptions.Timeout as exc:

                    result_text = f'下載逾時：{exc}'

                except Exception as exc:

                    result_text = f'失敗：{exc}'

                elapsed = round(
                    time.perf_counter() - start, 2
                )

                logs.append({
                    '公司代號': code,
                    '公司名稱': name,
                    '年度': report_year,
                    '版本': version,
                    'HTTP狀態碼': status,
                    'Content-Type': content_type,
                    '檔案大小(bytes)': size,
                    '結果': result_text,
                    '檔案路徑': filepath,
                    '耗時(秒)': elapsed
                })

                notify(
                    progress,
                    index,
                    len(rows),
                    f'下載 {code} {name}：{result_text}，'
                    f'耗時 {elapsed} 秒'
                )

                if index < len(rows) and delay:
                    time.sleep(delay)

            notify(
                progress,
                len(rows),
                len(rows),
                '正在整理 CSV 執行紀錄及完成 ZIP 打包…'
            )

            fields = list(
                dict.fromkeys(
                    k for row in rows for k in row
                )
            )

            required_fields = [
                'code',
                'inputCompanyName',
                'reportYear',
                'twFirstReportDownloadId',
                'twEditReportDownloadId'
            ]

            fields += [
                key for key in required_fields
                if key not in fields
            ]

            archive.writestr(
                f'{year}_篩選結果.csv',
                csv_bytes(rows, fields)
            )

            archive.writestr(
                f'{year}_查詢紀錄.csv',
                csv_bytes(query_logs, LOG_FIELDS)
            )

            archive.writestr(
                'download_log.csv',
                csv_bytes(logs, DL_FIELDS)
            )

    finally:
        session.close()

    print(
        f'ZIP 打包完成，成功下載 {success} 份 PDF。',
        flush=True
    )

    return output.getvalue(), logs, success
