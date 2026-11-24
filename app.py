import hashlib
import time
import traceback
import pandas as pd
import streamlit as st
from backend import load_companies, query_reports, make_zip

st.set_page_config(page_title='ESG 永續報告書下載系統',page_icon='📄')
st.title('ESG 永續報告書下載系統')
st.caption('資料來源：臺灣證券交易所 ESG 平台｜模組化測試版')
st.info('每次最多 5 家；單份 PDF 60 MB 上限於接收後檢查。請勿上傳敏感資料。')
year = st.number_input('報告年度',2015,2026,2025,1)
upload = st.file_uploader('上傳公司清單.xlsx',type=['xlsx'],help='工作表：公司清單；欄位：證券代碼、公司名稱')
with st.expander('進階篩選設定'):
    industries = st.text_input('產業（逗號分隔，空白不限）')
    keywords = st.text_input('公司名稱關鍵字（逗號分隔，空白不限）')
    prefer = st.checkbox('優先下載修正版',value=True)
    split = lambda s: tuple(x.strip() for x in s.replace('，',',').split(',') if x.strip())
key = None
if upload is not None:
    try:
        data = upload.getvalue()
        key = (hashlib.sha256(data).hexdigest(),year,industries,keywords,prefer)
        companies = load_companies(data)
        st.success(f'成功讀取 {len(companies)} 家公司')
        st.dataframe(pd.DataFrame(companies).rename(columns={'code':'證券代碼','name':'公司名稱'}),hide_index=True,width='stretch')
        if len(companies)>5:
            st.warning('請縮減至 5 家以內。')
        elif st.button('開始查詢並下載',type='primary',width='stretch'):
            st.session_state.pop('result',None)
            start = time.perf_counter(); bar = st.progress(0.0); status = st.empty()
            def update(i,total,msg,base,weight):
                bar.progress(min(0.99,base+weight*i/max(total,1)))
                status.info(f'{msg}｜已執行 {time.perf_counter()-start:.1f} 秒')
            rows, qlogs = query_reports(companies,year,progress=lambda i,t,m:update(i,t,m,0,0.4),target_industries=split(industries),name_keywords=split(keywords))
            qt = time.perf_counter()-start
            result = {'key':key,'qlogs':qlogs,'qt':qt,'rows':len(rows)}
            if rows:
                z,logs,success = make_zip(companies,rows,qlogs,year,progress=lambda i,t,m:update(i,t,m,0.4,0.59),prefer_edit=prefer)
                result.update(z=z,logs=logs,success=success,dt=time.perf_counter()-start-qt)
                status.success('處理完成') if success==len(rows) else status.warning('下載有失敗項目，請查看紀錄')
            else:
                status.warning('沒有可下載報告，請查看查詢紀錄')
            bar.progress(1.0)
            result['total']=time.perf_counter()-start
            st.session_state['result']=result
    except Exception as exc:
        st.error(f'處理失敗：{exc}'); traceback.print_exc()
r=st.session_state.get('result')
if r and key is not None and r['key']==key:
    st.write(f"API 查詢：{r['qt']:.2f} 秒；總耗時：{r['total']:.2f} 秒")
    with st.expander('API 查詢紀錄',expanded=not r['rows']):
        st.dataframe(pd.DataFrame(r['qlogs']),hide_index=True,width='stretch')
    if 'z' in r:
        st.write(f"成功下載 {r['success']} / {r['rows']} 份；下載及打包 {r['dt']:.2f} 秒")
        st.dataframe(pd.DataFrame(r['logs']),hide_index=True,width='stretch')
        st.download_button('下載 ZIP（PDF 與執行紀錄）',r['z'],file_name=f'ESG_{year}_reports.zip',mime='application/zip',width='stretch',on_click='ignore')
st.caption('公開測試版：結果不會永久保存在伺服器。')
