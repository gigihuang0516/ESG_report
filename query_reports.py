import re
import time
from pathlib import Path
import requests
from openpyxl import load_workbook

API_URL = "https://esggenplus.twse.com.tw/api/api/MopsSustainReport/data"
TIMEOUT_SECONDS = 120

ZERO_UUID = "00000000-0000-0000-0000-000000000000"
UUID_PATTERN = re.compile(
    r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$"
)


def normalize_cell(value):
    """將 Excel/API 欄位轉成乾淨字串，避免出現 1101.0。"""
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    value = str(value).strip().replace("\u3000", "")
    if value.endswith(".0") and value[:-2].isdigit():
        value = value[:-2]
    return value


def has_valid_uuid(value):
    value = normalize_cell(value)
    return bool(UUID_PATTERN.fullmatch(value)) and value.lower() != ZERO_UUID


def load_company_list(file_path):
    """讀取 Excel 公司清單，依證券代碼去除重複資料。"""
    if isinstance(file_path, Path) and not file_path.exists():
        raise FileNotFoundError(
            f"找不到公司清單：{file_path}\n"
            "請將「公司清單.xlsx」放在本 Python 程式所在資料夾。"
        )

    workbook = load_workbook(file_path, read_only=True, data_only=True)
    try:
        if "公司清單" not in workbook.sheetnames:
            raise ValueError("Excel 找不到「公司清單」工作表。")

        worksheet = workbook["公司清單"]
        headers = [normalize_cell(cell.value) for cell in worksheet[1]]
        required = ["證券代碼", "公司名稱"]
        missing = [column for column in required if column not in headers]
        if missing:
            raise ValueError(
                f"Excel 缺少必要欄位：{', '.join(missing)}。\n"
                "第一列必須包含「證券代碼」及「公司名稱」。"
            )

        code_index = headers.index("證券代碼")
        name_index = headers.index("公司名稱")
        companies = []
        seen_codes = set()

        for row_number, values in enumerate(
            worksheet.iter_rows(min_row=2, values_only=True), start=2
        ):
            code = normalize_cell(values[code_index] if code_index < len(values) else None)
            name = normalize_cell(values[name_index] if name_index < len(values) else None)
            if not code:
                continue
            if code in seen_codes:
                print(f"提醒：Excel 第 {row_number} 列的證券代碼 {code} 重複，已略過。")
                continue
            seen_codes.add(code)
            companies.append({"code": code, "name": name})

        if not companies:
            raise ValueError(
                "公司清單沒有有效資料。請在「公司清單」工作表第 2 列開始填入證券代碼。"
            )
        return companies
    finally:
        workbook.close()


def row_company_name(row, excel_name):
    return (
        normalize_cell(row.get("shortName"))
        or normalize_cell(row.get("name"))
        or normalize_cell(excel_name)
    )


def query_reports(companies, year, progress=None, delay=0.5, target_industries=(), name_keywords=()):
    selected, logs = [], []
    with requests.Session() as session:
        session.headers.update({"Content-Type": "application/json", "User-Agent": "Mozilla/5.0", "Origin": "https://esggenplus.twse.com.tw", "Referer": "https://esggenplus.twse.com.tw/inquiry/report?lang=zh-TW"})
        for index, company in enumerate(companies, 1):
            code, name = company['code'], company['name']
            start = time.perf_counter()
            msg = f'正在查詢 {code} {name}'
            print(msg, flush=True)
            if progress: progress(index-1, len(companies), msg)
            status, count, returned, ok, message = '', '', 0, False, ''
            matches = []
            try:
                payload = {"marketType": 0, "year": int(year), "industryNameList": [], "companyCodeList": [code], "industryName": "all", "companyCode": code}
                response = session.post(API_URL, json=payload, timeout=TIMEOUT_SECONDS)
                status = response.status_code
                response.raise_for_status()
                result = response.json()
                count = result.get('count', '')
                if not result.get('success', False): raise ValueError(result.get('message') or 'API 回傳失敗')
                data = result.get('data') or []
                if not isinstance(data, list): raise ValueError('API data 不是清單')
                returned = len(data)
                for item in data:
                    if not isinstance(item, dict): continue
                    # Do not overwrite an unverified company identifier.
                    identifiers = [normalize_cell(item.get(k)) for k in ('code', 'companyCode') if normalize_cell(item.get(k))]
                    if identifiers:
                        if any(v != code for v in identifiers): continue
                    elif not name or name not in (normalize_cell(item.get('shortName')), normalize_cell(item.get('name'))):
                        continue
                    row = dict(item, code=code, inputCompanyName=name, reportYear=str(year))
                    if target_industries and normalize_cell(row.get('sector')) not in target_industries: continue
                    if name_keywords and not any(k in row_company_name(row, name) for k in name_keywords): continue
                    if not any(has_valid_uuid(row.get(k)) for k in ('twEditReportDownloadId','twFirstReportDownloadId')): continue
                    matches.append(row)
                # At most one distinct report per company/year; ambiguity is logged.
                unique = {}
                for row in matches:
                    uid = row.get('twEditReportDownloadId') if has_valid_uuid(row.get('twEditReportDownloadId')) else row.get('twFirstReportDownloadId')
                    unique[normalize_cell(uid).lower()] = row
                if len(unique) > 1:
                    message = '同公司年度有多個不同下載 UUID，暫不下載，需人工確認'
                elif unique:
                    selected.extend(unique.values())
                elif returned:
                    message = '無符合條件且可核對公司身分的報告；請檢查 API 欄位'
                ok = True
                state = '查詢成功'
            except Exception as exc:
                state, message = '查詢錯誤', str(exc)
            elapsed = round(time.perf_counter()-start, 2)
            logs.append({'證券代碼':code,'Excel公司名稱':name,'HTTP狀態碼':status,'API成功':'是' if ok else '否','API count':count,'實際回傳筆數':returned,'查詢結果':state,'訊息':message,'耗時(秒)':elapsed})
            msg = f'{code} {name}：{state}，{elapsed} 秒；{message}'
            print(msg, flush=True)
            if progress: progress(index,len(companies),msg)
            if index < len(companies): time.sleep(delay)
    return selected, logs
