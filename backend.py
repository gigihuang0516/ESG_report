import csv
import io
import re
import time
import zipfile
from query_reports import load_company_list, query_reports
from download_pdf import new_session, choose_version, download_pdf

def load_companies(excel_bytes):
    return load_company_list(io.BytesIO(excel_bytes))

def csv_bytes(rows):
    out = io.StringIO(newline='')
    fields = list(dict.fromkeys(k for row in rows for k in row))
    writer = csv.DictWriter(out, fieldnames=fields)
    writer.writeheader(); writer.writerows(rows)
    return ('\ufeff'+out.getvalue()).encode('utf-8')

def make_zip(companies, rows, query_logs, year, progress=None, delay=1, prefer_edit=True):
    company_map = {c['code']:c['name'] for c in companies}
    output, logs, success = io.BytesIO(), [], 0
    seen = set()
    def notify(i,msg):
        print(msg,flush=True)
        if progress: progress(i,len(rows),msg)
    with new_session() as session, zipfile.ZipFile(output,'w',compression=zipfile.ZIP_DEFLATED) as archive:
        for index,row in enumerate(rows,1):
            code = row.get('code',''); name = company_map.get(code,'')
            log = {'公司代號':code,'公司名稱':name,'年度':year,'版本':'','HTTP狀態碼':'','Content-Type':'','檔案大小(bytes)':0,'結果':'','檔案路徑':''}
            start = time.perf_counter()
            notify(index-1,f'正在下載 {code} {name}（等待完整回應）')
            try:
                if code not in company_map: raise ValueError('不在 Excel 公司清單內')
                uid, version = choose_version(row,prefer_edit)
                log['版本'] = version
                filename = re.sub(r'[\\/:*?"<>|\x00-\x1f]','_',f'{code}_{name}_{year}.pdf')
                if filename in seen: raise ValueError('本次 ZIP 已有同名檔案')
                content = download_pdf(session,uid,log)
                notify(index-1,f'正在打包 {code} {name}')
                archive.writestr('PDF/'+filename,content,compress_type=zipfile.ZIP_STORED)
                seen.add(filename); success += 1
                log.update({'結果':'成功','檔案路徑':'PDF/'+filename})
                del content
            except Exception as exc:
                log['結果'] = f'失敗：{exc}'
            log['耗時(秒)'] = round(time.perf_counter()-start,2)
            logs.append(log)
            notify(index,f"下載 {code} {name}：{log['結果']}，耗時 {log['耗時(秒)']} 秒")
            if index < len(rows): time.sleep(delay)
        archive.writestr(f'{year}_篩選結果.csv',csv_bytes(rows))
        archive.writestr(f'{year}_查詢紀錄.csv',csv_bytes(query_logs))
        archive.writestr('download_log.csv',csv_bytes(logs))
    return output.getvalue(),logs,success
