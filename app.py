import streamlit as st
import pandas as pd
from backend import load_companies, query_reports, make_zip

st.set_page_config(page_title='ESG 永續報告書下載系統', page_icon='📄', layout='centered')
st.title('📄 ESG 永續報告書下載系統')
st.caption('上傳公司清單，查詢臺灣上市公司永續報告書並下載 PDF。資料來源：臺灣證券交易所 ESG 平台。')
st.info('第一版測試限制：每次最多 5 家公司、單份 PDF 最多 60 MB。請勿上傳個人敏感資料。')
year = st.number_input('報告年度', min_value=2015, max_value=2026, value=2025, step=1)
upload = st.file_uploader('上傳公司清單.xlsx', type=['xlsx'], help='工作表：公司清單；欄位：證券代碼、公司名稱')

if upload is not None:
    try:
        data = upload.getvalue()
        companies = load_companies(data)
        st.success(f'成功讀取 {len(companies)} 家公司')
        st.dataframe(pd.DataFrame(companies).rename(columns={'code': '證券代碼', 'name': '公司名稱'}), hide_index=True, use_container_width=True)
        if len(companies) > 5:
            st.warning('第一版一次最多處理 5 家公司，請縮減 Excel 清單。')
        else:
            if st.button('開始查詢並下載', type='primary', use_container_width=True):
                bar = st.progress(0)
                status = st.empty()
                def on_query(i, total, msg):
                    bar.progress(i / total * 0.4)
                    status.write(msg)
                with st.spinner('正在查詢 ESG 平台…'):
                    rows, query_logs = query_reports(companies, year, progress=on_query)
                if not any(log['API成功'] == '是' for log in query_logs):
                    st.error('所有公司查詢均失敗，請檢查 API 狀態或雲端存取限制。')
                    st.dataframe(pd.DataFrame(query_logs), hide_index=True)
                elif not rows:
                    st.warning('查詢成功，但沒有符合條件且可下載的中文報告書。')
                    st.dataframe(pd.DataFrame(query_logs), hide_index=True)
                else:
                    status.write(f'找到 {len(rows)} 份可下載報告，開始下載 PDF…')
                    def on_download(i, total, msg):
                        bar.progress(0.4 + i / total * 0.6)
                        status.write(msg)
                    with st.spinner('正在下載並打包 ZIP…'):
                        zip_data, logs, success = make_zip(companies, rows, query_logs, year, progress=on_download)
                    bar.progress(1.0)
                    st.write(f'成功下載 **{success}** 份／共 **{len(rows)}** 份')
                    st.dataframe(pd.DataFrame(logs), hide_index=True, use_container_width=True)
                    st.session_state['result_zip'] = zip_data
                    st.session_state['result_year'] = year
                    if success == 0:
                        st.error('沒有成功下載任何 PDF。可下載 ZIP 查看錯誤紀錄。')
    except Exception as exc:
        st.error(f'處理失敗：{exc}')

if 'result_zip' in st.session_state:
    st.download_button('⬇️ 下載 ZIP（PDF 與執行紀錄）', data=st.session_state['result_zip'], file_name=f"ESG_{st.session_state['result_year']}_reports.zip", mime='application/zip', use_container_width=True)

st.caption('公開測試版：網站執行資源有限，下載結果不會永久保存在伺服器。')
