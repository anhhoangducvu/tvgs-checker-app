# -*- coding: utf-8 -*-
"""
tvgs-checker-app — Kiểm tra báo cáo TVGS (Phụ lục IV NĐ 06/2021) trên web.
Phòng Kỹ thuật — Công ty Cổ phần TEXO Tư vấn và Đầu tư.

Chạy local:  streamlit run app.py
Deploy:      Streamlit Community Cloud (xem README.md)
"""

import os
import json
import tempfile
import datetime

import streamlit as st

from core import extract_sections as ex
from core import analyzer
from core.generate_review_docx import generate as gen_docx
from core.render_html import build_html

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CRIT_DIR = os.path.join(APP_DIR, 'criteria')

st.set_page_config(page_title='TVGS Checker — TEXO', page_icon='📋', layout='wide')

# ---------------- BẢO MẬT CƠ BẢN ----------------
PASSWORD = st.secrets.get('APP_PASSWORD', 'texo2026') if hasattr(st, 'secrets') else 'texo2026'
try:
    PASSWORD = st.secrets.get('APP_PASSWORD', 'texo2026')
except Exception:
    PASSWORD = 'texo2026'

if 'authed' not in st.session_state:
    st.session_state.authed = False

if not st.session_state.authed:
    st.markdown("<h2 style='color:#1F4E79'>📋 TVGS Checker — Phòng Kỹ thuật TEXO</h2>",
                unsafe_allow_html=True)
    pw = st.text_input('Nhập mật khẩu để sử dụng', type='password')
    if st.button('Đăng nhập', type='primary') or pw:
        if pw == PASSWORD:
            st.session_state.authed = True
            st.rerun()
        elif pw:
            st.error('Sai mật khẩu.')
    st.stop()

# ---------------- GIAO DIỆN CHÍNH ----------------
st.markdown(
    "<div style='border-bottom:3px solid #1F4E79;padding-bottom:6px;margin-bottom:14px'>"
    "<span style='color:#1F4E79;font-size:26px;font-weight:800'>📋 TVGS Checker</span> "
    "<span style='color:#E8731A;font-weight:700'> — Phiếu đánh giá báo cáo TVGS</span><br>"
    "<span style='color:#595959'>Phòng Kỹ thuật — Công ty Cổ phần TEXO Tư vấn và Đầu tư · "
    "Kiểm tra theo Phụ lục IV NĐ 06/2021 (định kỳ 8 mục / hoàn thành 12 mục) · "
    "công cụ quy tắc, không dùng AI</span></div>", unsafe_allow_html=True)

with st.sidebar:
    st.header('⚙️ Thiết lập')
    type_choice = st.radio('Loại báo cáo',
                           ['Tự nhận diện', 'Báo cáo HOÀN THÀNH (12 mục)', 'Báo cáo ĐỊNH KỲ (8 mục)'])
    so_phieu = st.text_input('Số phiếu đánh giá', value=f'Số: ...../{datetime.date.today().year}/PKT-ĐG')
    nguoi_dg = st.text_input('Người đánh giá (để trống nếu chưa rõ)', value='')
    truong_phong = st.text_input('Trưởng phòng Kỹ thuật', value='Hoàng Đức Vũ')
    st.divider()
    st.caption('Kết quả là kiểm tra tự động bằng quy tắc (regex/đối chiếu chéo), '
               'KHÔNG thay thế đánh giá chuyên môn. Bản dùng AI đầy đủ: skill tvgs-checker trên Claude.')

up = st.file_uploader('Tải lên báo cáo TVGS (.docx — khuyến nghị, hoặc .pdf)',
                      type=['docx', 'pdf'])

if not up:
    st.info('⬆️ Tải file báo cáo lên để bắt đầu kiểm tra.')
    st.stop()

# ---------------- XỬ LÝ ----------------
suffix = '.pdf' if up.name.lower().endswith('.pdf') else '.docx'
with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tf:
    tf.write(up.getvalue())
    tmp_path = tf.name

try:
    items = ex.walk_pdf(tmp_path) if suffix == '.pdf' else ex.walk_docx(tmp_path)
except Exception as err:
    st.error(f'Không đọc được file: {err}')
    st.stop()

forced = None
if 'HOÀN THÀNH' in type_choice:
    forced = 'hoan_thanh'
elif 'ĐỊNH KỲ' in type_choice:
    forced = 'dinh_ky'
rtype, conf, signals = ex.detect_type(items, forced)
if rtype == 'unknown':
    rtype = 'hoan_thanh'
    signals.append('Không chắc chắn — mặc định HOÀN THÀNH, hãy chọn thủ công ở sidebar nếu sai.')

sections, meta = ex.parse(items, rtype)
extracted = {'_meta': {'source_file': up.name, 'loai_bao_cao': rtype,
                       'do_tin_cay': conf, 'dau_hieu_nhan_dien': signals, **meta},
             **sections}

crit_file = 'dinh_ky_8muc.json' if rtype == 'dinh_ky' else 'hoan_thanh_12muc.json'
with open(os.path.join(CRIT_DIR, crit_file), encoding='utf-8') as f:
    criteria = json.load(f)

review = analyzer.analyze(extracted, criteria,
                          nguoi_danh_gia_ten=nguoi_dg,
                          truong_phong_ten=truong_phong,
                          so_phieu=so_phieu)
review.update({k: v for k, v in analyzer.guess_project_info(extracted).items() if v})

# ---------------- THÔNG TIN NHẬN DIỆN ----------------
c1, c2, c3 = st.columns(3)
c1.metric('Loại báo cáo', 'ĐỊNH KỲ (8 mục)' if rtype == 'dinh_ky' else 'HOÀN THÀNH (12 mục)',
          f'độ tin cậy {conf}')
c2.metric('Mục tìm thấy', f"{len(meta['sections_found'])}/{meta['so_muc_chuan']}",
          f"thiếu {len(meta['sections_missing'])}" if meta['sections_missing'] else 'đủ mục')
c3.metric('Xếp loại sơ bộ', review['ket_qua_chung']['xep_loai'])
with st.expander('Dấu hiệu nhận diện loại báo cáo'):
    for s in signals:
        st.write('•', s)

# ---------------- CHO PHÉP SỬA TRƯỚC KHI XUẤT ----------------
st.subheader('✏️ Hiệu chỉnh trước khi xuất phiếu')
cc1, cc2 = st.columns(2)
review['du_an'] = cc1.text_input('Dự án / Công trình', review.get('du_an', ''))
review['hang_muc_giai_doan'] = cc2.text_input('Hạng mục / Giai đoạn',
                                              review.get('hang_muc_giai_doan', ''))
review['chu_dau_tu'] = cc1.text_input('Chủ đầu tư', review.get('chu_dau_tu', ''))
review['don_vi_lap'] = cc2.text_input('Đơn vị lập báo cáo', review.get('don_vi_lap', ''))
if rtype == 'dinh_ky':
    review['ky_bao_cao'] = st.text_input('Kỳ báo cáo', review.get('ky_bao_cao', ''))

xl_opts = ['ĐẠT', 'ĐẠT CÓ ĐIỀU KIỆN', 'CẦN SỬA', 'KHÔNG ĐẠT']
xl_now = review['ket_qua_chung']['xep_loai']
review['ket_qua_chung']['xep_loai'] = st.selectbox(
    'Xếp loại', xl_opts, index=xl_opts.index(xl_now) if xl_now in xl_opts else 2)
review['ket_qua_chung']['tom_tat'] = st.text_area(
    'Tóm tắt kết quả', review['ket_qua_chung']['tom_tat'], height=90)
review['ket_luan'] = st.text_area('Kết luận', review['ket_luan'], height=120)

with st.expander('Sửa trạng thái / nhận xét từng mục', expanded=False):
    st_opts = ['DAT', 'CAN_SUA', 'LOI', 'THIEU']
    for i, m in enumerate(review['danh_gia_muc']):
        a, b, c = st.columns([1.2, 1, 3])
        a.markdown(f"**{m['muc']}** — {m['ten'][:45]}")
        m['trang_thai'] = b.selectbox('Trạng thái', st_opts,
                                      index=st_opts.index(m['trang_thai']),
                                      key=f'st{i}', label_visibility='collapsed')
        m['nhan_xet_ngan'] = c.text_input('Nhận xét', m['nhan_xet_ngan'],
                                          key=f'nx{i}', label_visibility='collapsed')

# ---------------- KẾT QUẢ ----------------
html_str = build_html(review)

tab1, tab2, tab3 = st.tabs(['📄 Phiếu đánh giá (HTML)', '🔬 Chi tiết trích xuất', '🧾 JSON'])
with tab1:
    st.components.v1.html(html_str, height=1400, scrolling=True)
with tab2:
    st.write('**Phụ lục trong file:**',
             ', '.join(a['ten'] for a in meta['phu_luc_found']) or 'không có')
    st.write('**Phụ lục được dẫn chiếu trong thân:**',
             ', '.join(meta['phu_luc_referenced_in_body']) or 'không có')
    ck = meta['chu_ky']
    st.write(f"**Chữ ký:** Giám sát trưởng: {'✅' if ck['giam_sat_truong'] else '❌'} · "
             f"Đại diện pháp luật: {'✅' if ck['dai_dien_phap_luat'] else '❌'}")
    for k in ['trang_bia', 'mo_dau', 'can_cu'] + \
             [f"muc_{i}" for i in range(1, meta['so_muc_chuan'] + 1)] + ['ket_luan']:
        if k in sections:
            with st.expander(f"{k} — {len(sections[k])} ký tự, "
                             f"{meta['tables_per_section'].get(k, 0)} bảng"):
                st.text(sections[k][:4000])
with tab3:
    st.json(review)

# ---------------- TẢI VỀ ----------------
st.subheader('⬇️ Tải kết quả')
base = (f"Danh gia bao cao TVGS {'dinh ky' if rtype == 'dinh_ky' else 'hoan thanh'} - "
        f"{(review.get('du_an') or up.name.rsplit('.', 1)[0])[:60]}").strip()
d1, d2 = st.columns(2)
with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tf:
    out_docx = tf.name
gen_docx(review, out_docx)
with open(out_docx, 'rb') as f:
    d1.download_button('📥 Tải phiếu đánh giá (.docx)', f.read(),
                       file_name=base + '.docx',
                       mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                       type='primary')
d2.download_button('📥 Tải phiếu đánh giá (.html)', html_str.encode('utf-8'),
                   file_name=base + '.html', mime='text/html')
