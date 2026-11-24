
import time
import hashlib
import traceback

import streamlit as st
import pandas as pd

from backend import (
    load_companies,
    query_reports,
    make_zip
)


# =========================
# 網頁基本設定
# =========================

st.set_page_config(
    page_title='ESG 永續報告書下載系統',
    page_icon='📄',
    layout='centered'
)

st.title('📄 ESG 永續報告書下載系統')

st.caption(
    '上傳公司清單，查詢臺灣上市公司永續報告書'
    '並下載 PDF。資料來源：臺灣證券交易所 ESG 平台。'
)

st.info(
    '測試版限制：每次最多 5 家公司、'
    '單份 PDF 最多 60 MB。'
    '請勿上傳個人敏感資料。'
)


# =========================
# 使用者輸入
# =========================

year = st.number_input(
    '報告年度',
    min_value=2015,
    max_value=2026,
    value=2025,
    step=1
)

upload = st.file_uploader(
    '上傳公司清單.xlsx',
    type=['xlsx'],
    help='工作表：公司清單；欄位：證券代碼、公司名稱'
)

current_key = None


# =========================
# Excel 讀取
# =========================

if upload is not None:

    try:

        data = upload.getvalue()

        current_key = (
            hashlib.sha256(data).hexdigest(),
            int(year)
        )

        companies = load_companies(data)

        st.success(
            f'成功讀取 {len(companies)} 家公司'
        )

        st.dataframe(
            pd.DataFrame(companies).rename(
                columns={
                    'code': '證券代碼',
                    'name': '公司名稱'
                }
            ),
            hide_index=True,
            width='stretch'
        )

        if len(companies) > 5:

            st.warning(
                '測試版一次最多處理 5 家公司，'
                '請縮減 Excel 清單。'
            )

        else:

            # =========================
            # 開始執行
            # =========================

            if st.button(
                '開始查詢並下載',
                type='primary',
                width='stretch'
            ):

                st.session_state.pop('result_zip', None)
                st.session_state.pop('result_key', None)
                st.session_state.pop('result_logs', None)
                st.session_state.pop('result_query_logs', None)
                st.session_state.pop('result_summary', None)

                total_start = time.perf_counter()

                bar = st.progress(0.0)
                status = st.empty()
                timer = st.empty()

                def on_query(i, total, msg):

                    fraction = i / total if total else 0

                    bar.progress(
                        min(max(fraction * 0.4, 0.0), 0.4)
                    )

                    status.info(msg)

                    timer.caption(
                        f'已執行 '
                        f'{time.perf_counter() - total_start:.1f} 秒'
                    )

                def on_download(i, total, msg):

                    fraction = i / total if total else 0

                    bar.progress(
                        min(
                            max(0.4 + fraction * 0.59, 0.4),
                            0.99
                        )
                    )

                    status.info(msg)

                    timer.caption(
                        f'已執行 '
                        f'{time.perf_counter() - total_start:.1f} 秒'
                    )

                # =========================
                # API 查詢
                # =========================

                query_start = time.perf_counter()

                status.info('正在準備 API 查詢…')

                rows, query_logs = query_reports(
                    companies,
                    year,
                    progress=on_query
                )

                query_elapsed = round(
                    time.perf_counter() - query_start,
                    2
                )

                st.write(
                    f'API 查詢耗時：**{query_elapsed} 秒**'
                )

                st.session_state['result_query_logs'] = (
                    query_logs
                )

                st.session_state['result_key'] = current_key

                if not any(
                    log['API成功'] == '是'
                    for log in query_logs
                ):

                    status.error(
                        '所有公司 API 查詢均失敗。'
                    )

                    st.error(
                        '請查看下方查詢紀錄，'
                        '確認是否為連線逾時或 HTTP 錯誤。'
                    )

                elif not rows:

                    status.warning(
                        '查詢完成，但沒有符合條件的報告書。'
                    )

                else:

                    st.success(
                        f'找到 {len(rows)} 份可下載報告書'
                    )

                    # =========================
                    # PDF 下載及 ZIP 打包
                    # =========================

                    download_start = time.perf_counter()

                    status.info('正在準備下載 PDF…')

                    zip_data, logs, success = make_zip(
                        companies,
                        rows,
                        query_logs,
                        year,
                        progress=on_download
                    )

                    download_elapsed = round(
                        time.perf_counter() - download_start,
                        2
                    )

                    total_elapsed = round(
                        time.perf_counter() - total_start,
                        2
                    )

                    bar.progress(1.0)

                    timer.caption(
                        f'總執行時間：{total_elapsed} 秒'
                    )

                    if success == len(rows):

                        status.success(
                            '所有報告書下載及打包完成！'
                        )

                    elif success > 0:

                        status.warning(
                            '部分報告書下載成功，'
                            '請查看下載紀錄。'
                        )

                    else:

                        status.error(
                            '沒有成功下載任何 PDF，'
                            '可下載 ZIP 查看執行紀錄。'
                        )

                    st.session_state['result_zip'] = zip_data

                    st.session_state['result_year'] = year

                    st.session_state['result_logs'] = logs

                    st.session_state['result_summary'] = {
                        'success': success,
                        'total': len(rows),
                        'query_elapsed': query_elapsed,
                        'download_elapsed': download_elapsed,
                        'total_elapsed': total_elapsed
                    }

    except Exception as exc:

        st.error(f'處理失敗：{exc}')

        traceback.print_exc()


# =========================
# 顯示執行結果
# =========================

if (
    current_key is not None
    and st.session_state.get('result_key') == current_key
):

    if 'result_summary' in st.session_state:

        summary = st.session_state['result_summary']

        st.subheader('執行結果')

        col1, col2, col3 = st.columns(3)

        col1.metric(
            '成功下載',
            f"{summary['success']} 份"
        )

        col2.metric(
            'PDF 與 ZIP 耗時',
            f"{summary['download_elapsed']} 秒"
        )

        col3.metric(
            '總執行時間',
            f"{summary['total_elapsed']} 秒"
        )

    if 'result_query_logs' in st.session_state:

        with st.expander('查看 API 查詢紀錄'):

            st.dataframe(
                pd.DataFrame(
                    st.session_state['result_query_logs']
                ),
                hide_index=True,
                width='stretch'
            )

    if 'result_logs' in st.session_state:

        with st.expander(
            '查看 PDF 下載紀錄',
            expanded=True
        ):

            st.dataframe(
                pd.DataFrame(
                    st.session_state['result_logs']
                ),
                hide_index=True,
                width='stretch'
            )

    if 'result_zip' in st.session_state:

        zip_data = st.session_state['result_zip']

        zip_mb = len(zip_data) / 1024 / 1024

        st.download_button(
            f'⬇️ 下載 ZIP（{zip_mb:.2f} MB）',
            data=zip_data,
            file_name=(
                f"ESG_{st.session_state['result_year']}"
                f"_reports.zip"
            ),
            mime='application/zip',
            width='stretch',
            on_click='ignore'
        )


st.divider()

st.caption(
    '公開測試版：網站執行資源有限，'
    '下載結果不會永久保存在伺服器。'
)
