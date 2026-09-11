# -*- coding: utf-8 -*-
import io
import datetime
import streamlit as st
from core.feishu import load_all_data
from core.rebuild import rebuild_from_feishu

st.set_page_config(page_title='分析导出', layout='centered')

st.markdown("""
<style>
html, body, [data-testid="stAppViewContainer"], .stApp { background-color: #0d0d0d !important; }
[data-testid="stHeader"] { background-color: transparent !important; }
[data-testid="stToolbar"] { display: none !important; }
.stButton > button {
    background: #f5a623 !important; color: #0d0d0d !important;
    font-weight: 800 !important; border: none !important;
    border-radius: 0 !important; padding: 12px 28px !important;
}
[data-testid="stDownloadButton"] > button {
    background: transparent !important; color: #f5a623 !important;
    border: 1px solid rgba(245,166,35,.5) !important; border-radius: 0 !important;
    font-weight: 700 !important; padding: 12px 28px !important;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style="padding:40px 0 28px;text-align:center;">
    <div style="font-size:.6rem;letter-spacing:.28em;color:#2e2e2e;text-transform:uppercase;margin-bottom:12px;">
        PUBGM · Compare · Analysis Export
    </div>
    <div style="font-size:2.4rem;font-weight:900;color:#e8e8e8;letter-spacing:-.02em;">
        数据分析导出
    </div>
    <div style="width:28px;height:2px;background:#f5a623;margin:16px auto 0;"></div>
</div>
""", unsafe_allow_html=True)

st.divider()

if st.button('读取飞书 · 生成导出版 →', type='primary', use_container_width=True):
    with st.spinner('正在读取飞书区域数据…'):
        try:
            all_campaigns, country_total_ai = load_all_data()
        except Exception as e:
            st.error(f'飞书读取失败：{e}')
            st.stop()

    real_camp = sum(1 for c in all_campaigns if c['ai'] and c['ai'].strip() != '继续保持观察')
    real_ct   = sum(1 for v in country_total_ai.values() if v and v.strip() != '继续保持观察')
    st.success(
        f'飞书数据读取完成：{len(all_campaigns)} 条 campaign · '
        f'真实 AI 建议 {real_camp} 条（campaign级）+ {real_ct} 条（国家Total级）'
    )

    with st.spinner('正在重建数据分析 sheets…'):
        try:
            out_bytes = rebuild_from_feishu(all_campaigns, country_total_ai)
        except Exception as e:
            st.error(f'生成失败：{e}')
            import traceback
            with st.expander('详细错误'):
                st.code(traceback.format_exc())
            st.stop()

    now      = datetime.datetime.now().strftime('%Y%m%d_%H')
    out_name = f'Compare-{now}_数据分析_导出版.xlsx'

    st.download_button(
        label=f'DOWNLOAD ↓ {out_name}',
        data=out_bytes,
        file_name=out_name,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
    )
