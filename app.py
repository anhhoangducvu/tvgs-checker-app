# -*- coding: utf-8 -*-
"""
tvgs-checker-app — Kiểm tra báo cáo TVGS (Phụ lục IV NĐ 06/2021) trên web.
Phòng Kỹ thuật — Công ty Cổ phần TEXO Tư vấn và Đầu tư.

HYBRID: lớp quy tắc (rule-based) luôn chạy, miễn phí, cục bộ.
AI là LỰA CHỌN bổ sung — người dùng tự nhập API key (Claude/Gemini/OpenAI/tương thích).

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
from core import llm
from core import ai_reviewer
from core.generate_review_docx import generate as gen_docx
from core.render_html import build_html

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CRIT_DIR = os.path.join(APP_DIR, 'criteria')

# Tên cố định (đặt làm hằng số — không cần UI): người đánh giá để TRỐNG khi chưa rõ
TRUONG_PHONG_TEN = 'Hoàng Đức Vũ'
NGUOI_DANH_GIA_TEN = ''

st.set_page_config(page_title='TVGS Checker — TEXO', page_icon='📋', layout='wide')

# ---------------- BẢO MẬT CƠ BẢN ----------------
try:
    PASSWORD = st.secrets.get('APP_PASSWORD', None) or os.environ.get('APP_PASSWORD', 'texo2026')
except Exception:
    PASSWORD = os.environ.get('APP_PASSWORD', 'texo2026')

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

# ---------------- SIDEBAR: CHẾ ĐỘ AI + BẢO MẬT + HƯỚNG DẪN ----------------
with st.sidebar:
    st.header('🧠 Chế độ phân tích')
    use_ai = st.radio('Chọn chế độ', ['Không dùng AI (mặc định, miễn phí)',
                                      'Dùng AI (tự nhập API key)'],
                      label_visibility='collapsed').startswith('Dùng AI')

    provider = model = base_url = api_key = None
    ai_ready = False
    if use_ai:
        prov_keys = list(llm.PROVIDERS.keys())
        provider = st.selectbox('Nhà cung cấp AI', prov_keys,
                                format_func=lambda k: llm.PROVIDERS[k]['label'])
        st.caption(llm.PROVIDERS[provider]['note'])
        api_key = st.text_input('API key', type='password',
                                help='Key chỉ giữ trong phiên làm việc này, KHÔNG lưu lại.')
        if llm.PROVIDERS[provider]['needs_base_url']:
            base_url = st.text_input('Base URL (bắt buộc)',
                                     placeholder='https://openrouter.ai/api/v1')

        mkey = f'models_{provider}'
        if st.button('🔄 Lấy danh sách model', use_container_width=True,
                     disabled=not api_key):
            ok, res = llm.list_models(provider, api_key, base_url)
            if ok:
                st.session_state[mkey] = res
                st.success(f'Tìm thấy {len(res)} model.')
            else:
                st.error(res)
        models = st.session_state.get(mkey, [])
        if models:
            model = st.selectbox('Chọn model', models)
        else:
            st.caption('Dán key rồi bấm "Lấy danh sách model" để chọn model thật của key '
                       '(không gõ tay tên model — tên model đổi liên tục).')

        if model and st.button('✅ Kiểm tra kết nối', use_container_width=True):
            with st.spinner('Đang gọi thử...'):
                ok, msg = llm.test_connection(provider, model, api_key, base_url)
            st.success(f'Kết nối OK: {msg}') if ok else st.error(msg)

        ai_ready = bool(api_key and model and
                        (base_url or not llm.PROVIDERS[provider]['needs_base_url']))

    st.divider()
    st.subheader('🔐 Bảo mật')
    st.markdown(
        '- API key **chỉ giữ trong phiên**, app **không lưu** key của bạn.\n'
        '- Bật AI ⇒ **toàn văn báo cáo được gửi** tới nhà cung cấp AI bạn chọn. '
        'Báo cáo mật/nhạy cảm → dùng chế độ Không AI (xử lý cục bộ).\n'
        '- Không dán key trên máy công cộng; không chia sẻ key; lộ key do cách dùng '
        'thuộc trách nhiệm người dùng.\n'
        '- Mật khẩu app đổi được qua **Settings → Secrets** (`APP_PASSWORD`).')

    with st.expander('📖 Hướng dẫn lấy & dùng API key'):
        st.markdown(
            '**Lấy key ở đâu?**\n'
            '- Claude: console.anthropic.com → API keys\n'
            '- Gemini: aistudio.google.com → Get API key (có bậc miễn phí)\n'
            '- OpenAI: platform.openai.com → API keys\n'
            '- OpenRouter: openrouter.ai → Keys (dùng loại "Tương thích OpenAI", '
            'Base URL `https://openrouter.ai/api/v1`)\n\n'
            '**Các bước:** dán key → bấm *Lấy danh sách model* → chọn model → '
            '*Kiểm tra kết nối* → quay lại trang chính bấm *Phân tích sâu bằng AI*.\n\n'
            '**Lỗi thường gặp:**\n'
            '- `401/403`: key sai, hoặc dùng key OpenRouter mà quên điền Base URL\n'
            '- `404`: model đã bị gỡ → bấm lại *Lấy danh sách model*\n'
            '- `429`: hết hạn mức → đổi model nhẹ hơn (flash/mini), chờ reset, hoặc bật billing\n'
            '- Kết quả AI bị cụt → app tự vá JSON, nhưng nên chọn model mạnh hơn')

    st.divider()
    st.caption('Chế độ Không AI: quét quy tắc TEXO, chạy cục bộ, miễn phí. '
               'Chế độ AI: đánh giá sâu từng mục theo chú thích kiểm tra gốc, '
               'trên nền kết quả quy tắc.')

# ---------------- HEADER ----------------
st.markdown(
    "<div style='border-bottom:3px solid #1F4E79;padding-bottom:6px;margin-bottom:8px'>"
    "<span style='color:#1F4E79;font-size:26px;font-weight:800'>📋 TVGS Checker</span> "
    "<span style='color:#E8731A;font-weight:700'> — Phiếu đánh giá báo cáo TVGS</span><br>"
    "<span style='color:#595959'>Phòng Kỹ thuật — Công ty Cổ phần TEXO Tư vấn và Đầu tư · "
    "Phụ lục IV NĐ 06/2021 (định kỳ 8 mục / hoàn thành 12 mục)</span></div>",
    unsafe_allow_html=True)

if not use_ai:
    st.info('🛡️ Chế độ **KHÔNG AI** — kiểm tra bằng bộ quy tắc TEXO, dữ liệu xử lý cục bộ.')
elif ai_ready:
    st.success(f'🧠 AI sẵn sàng: **{llm.PROVIDERS[provider]["label"]} / {model}** — '
               'kết quả quy tắc vẫn chạy trước, AI phân tích sâu khi bạn bấm nút.')
else:
    st.warning('🧠 Đã chọn dùng AI nhưng **chưa đủ thiết lập** (key / model / Base URL) — '
               'hoàn thiện ở thanh bên trái. Trong lúc đó app vẫn chạy chế độ quy tắc.')

type_choice = st.radio('Loại báo cáo', ['Tự nhận diện', 'HOÀN THÀNH (12 mục)', 'ĐỊNH KỲ (8 mục)'],
                       horizontal=True)
up = st.file_uploader('Tải lên báo cáo TVGS (.docx — khuyến nghị, hoặc .pdf)',
                      type=['docx', 'pdf'])
if not up:
    st.info('⬆️ Tải file báo cáo lên để bắt đầu kiểm tra.')
    st.stop()

# ---------------- TRÍCH XUẤT + QUY TẮC (luôn chạy) ----------------
suffix = '.pdf' if up.name.lower().endswith('.pdf') else '.docx'
with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tf:
    tf.write(up.getvalue())
    tmp_path = tf.name

try:
    items = ex.walk_pdf(tmp_path) if suffix == '.pdf' else ex.walk_docx(tmp_path)
except Exception as err:
    st.error(f'Không đọc được file: {err}')
    st.stop()

forced = ('hoan_thanh' if 'HOÀN THÀNH' in type_choice
          else 'dinh_ky' if 'ĐỊNH KỲ' in type_choice else None)
rtype, conf, signals = ex.detect_type(items, forced)
if rtype == 'unknown':
    rtype = 'hoan_thanh'
    signals.append('Không chắc chắn — mặc định HOÀN THÀNH, chọn thủ công nếu sai.')

sections, meta = ex.parse(items, rtype)
extracted = {'_meta': {'source_file': up.name, 'loai_bao_cao': rtype,
                       'do_tin_cay': conf, 'dau_hieu_nhan_dien': signals, **meta},
             **sections}
crit_file = 'dinh_ky_8muc.json' if rtype == 'dinh_ky' else 'hoan_thanh_12muc.json'
with open(os.path.join(CRIT_DIR, crit_file), encoding='utf-8') as f:
    criteria = json.load(f)

base_review = analyzer.analyze(extracted, criteria,
                               nguoi_danh_gia_ten=NGUOI_DANH_GIA_TEN,
                               truong_phong_ten=TRUONG_PHONG_TEN)
base_review.update({k: v for k, v in analyzer.guess_project_info(extracted).items() if v})

# ---------------- LỚP AI (tùy chọn) ----------------
file_sig = f'{up.name}_{up.size}_{rtype}'
if st.session_state.get('ai_sig') != file_sig:
    st.session_state.pop('ai_review', None)   # file mới -> bỏ kết quả AI cũ
    st.session_state['ai_sig'] = file_sig

if use_ai and ai_ready:
    cA, cB = st.columns([1.2, 3])
    if cA.button('🧠 Phân tích sâu bằng AI', type='primary', use_container_width=True):
        with st.spinner('AI đang đọc và đánh giá từng mục (20–90 giây)...'):
            merged, err, truncated = ai_reviewer.ai_review(
                provider, model, api_key, base_url, extracted, criteria, base_review)
        if err:
            st.error(f'Gọi AI thất bại: {err} — app tiếp tục dùng kết quả quy tắc.')
        else:
            st.session_state['ai_review'] = merged
            if truncated:
                st.warning('Phản hồi AI có thể bị cắt cụt (đã tự vá JSON) — '
                           'nếu thiếu mục, thử model mạnh hơn.')
    if st.session_state.get('ai_review') is not None:
        cB.success('Đang hiển thị kết quả **AI + quy tắc**. '
                   'Bấm lại nút để phân tích lại nếu cần.')

ai_result = st.session_state.get('ai_review')
using_ai_result = use_ai and ai_ready and ai_result is not None
review = json.loads(json.dumps(ai_result if using_ai_result else base_review))
mode_tag = 'ai' if using_ai_result else 'rule'

# ---------------- THÔNG TIN NHẬN DIỆN ----------------
c1, c2, c3, c4 = st.columns(4)
c1.metric('Loại báo cáo', 'ĐỊNH KỲ (8 mục)' if rtype == 'dinh_ky' else 'HOÀN THÀNH (12 mục)',
          f'độ tin cậy {conf}')
c2.metric('Mục tìm thấy', f"{len(meta['sections_found'])}/{meta['so_muc_chuan']}",
          f"thiếu {len(meta['sections_missing'])}" if meta['sections_missing'] else 'đủ mục')
c3.metric('Xếp loại sơ bộ', review['ket_qua_chung']['xep_loai'])
c4.metric('Nguồn đánh giá', 'AI + quy tắc' if using_ai_result else 'Quy tắc (không AI)')
with st.expander('Dấu hiệu nhận diện loại báo cáo'):
    for s in signals:
        st.write('•', s)

# ---------------- HIỆU CHỈNH TRƯỚC KHI XUẤT ----------------
st.subheader('✏️ Hiệu chỉnh trước khi xuất phiếu')
cc1, cc2 = st.columns(2)
review['du_an'] = cc1.text_input('Dự án / Công trình', review.get('du_an', ''),
                                 key=f'da_{mode_tag}')
review['hang_muc_giai_doan'] = cc2.text_input('Hạng mục / Giai đoạn',
                                              review.get('hang_muc_giai_doan', ''),
                                              key=f'hm_{mode_tag}')
review['chu_dau_tu'] = cc1.text_input('Chủ đầu tư', review.get('chu_dau_tu', ''),
                                      key=f'cdt_{mode_tag}')
review['don_vi_lap'] = cc2.text_input('Đơn vị lập báo cáo', review.get('don_vi_lap', ''),
                                      key=f'dvl_{mode_tag}')
review['so_phieu'] = cc1.text_input('Số phiếu đánh giá', review.get('so_phieu', ''),
                                    key=f'sp_{mode_tag}')
if rtype == 'dinh_ky':
    review['ky_bao_cao'] = cc2.text_input('Kỳ báo cáo', review.get('ky_bao_cao', ''),
                                          key=f'ky_{mode_tag}')

xl_opts = ['ĐẠT', 'ĐẠT CÓ ĐIỀU KIỆN', 'CẦN SỬA', 'KHÔNG ĐẠT']
xl_now = review['ket_qua_chung']['xep_loai']
review['ket_qua_chung']['xep_loai'] = st.selectbox(
    'Xếp loại', xl_opts, index=xl_opts.index(xl_now) if xl_now in xl_opts else 2,
    key=f'xl_{mode_tag}')
review['ket_qua_chung']['tom_tat'] = st.text_area(
    'Tóm tắt kết quả', review['ket_qua_chung']['tom_tat'], height=100, key=f'tt_{mode_tag}')
review['ket_luan'] = st.text_area('Kết luận', review['ket_luan'], height=120,
                                  key=f'kl_{mode_tag}')

with st.expander('Sửa trạng thái / nhận xét từng mục', expanded=False):
    st_opts = ['DAT', 'CAN_SUA', 'LOI', 'THIEU']
    for i, m in enumerate(review['danh_gia_muc']):
        a, b, c = st.columns([1.2, 1, 3])
        a.markdown(f"**{m['muc']}** — {m['ten'][:45]}")
        m['trang_thai'] = b.selectbox('Trạng thái', st_opts,
                                      index=st_opts.index(m['trang_thai'])
                                      if m['trang_thai'] in st_opts else 1,
                                      key=f'st{i}_{mode_tag}', label_visibility='collapsed')
        m['nhan_xet_ngan'] = c.text_input('Nhận xét', m['nhan_xet_ngan'],
                                          key=f'nx{i}_{mode_tag}',
                                          label_visibility='collapsed')

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
if using_ai_result:
    base += ' (AI)'
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
