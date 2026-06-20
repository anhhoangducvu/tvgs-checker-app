# -*- coding: utf-8 -*-
"""
ai_reviewer.py — Lớp AI TÙY CHỌN của tvgs-checker-app (triết lý HYBRID).

AI KHÔNG thay thế lớp rule-based: kết quả rule-based được đưa vào prompt làm điểm tựa,
AI đánh giá sâu từng mục dựa trên CHÚ THÍCH KIỂM TRA GỐC (chu_thich_goc) và câu hỏi tư duy
trong bộ tiêu chí TEXO, rồi trả JSON cùng schema để merge ngược vào phiếu đánh giá.
"""
import json

from . import llm

MAX_CHARS_PER_MUC = 2600   # cắt bớt nội dung từng mục đưa vào prompt
VALID_ST = {'DAT', 'CAN_SUA', 'LOI', 'THIEU'}
VALID_XL = {'ĐẠT', 'ĐẠT CÓ ĐIỀU KIỆN', 'CẦN SỬA', 'KHÔNG ĐẠT'}

SCHEMA = """{
  "danh_gia_muc": [
    {"muc": "Mục 1", "trang_thai": "DAT|CAN_SUA|LOI|THIEU",
     "nhan_xet_ngan": "tối đa 15 từ cho bảng tổng hợp",
     "dat_duoc": ["điểm làm tốt..."],
     "van_de": ["vấn đề cụ thể, dẫn chứng từ nội dung mục..."]}
  ],
  "phat_hien_chinh": [{"muc_do": "LOI|CAN_SUA", "noi_dung": "..."}],
  "kien_nghi": ["yêu cầu chỉnh sửa cụ thể..."],
  "tom_tat": "2-3 câu tóm tắt chất lượng báo cáo",
  "ket_luan": "đoạn kết luận của Phòng Kỹ thuật",
  "xep_loai": "ĐẠT|ĐẠT CÓ ĐIỀU KIỆN|CẦN SỬA|KHÔNG ĐẠT"
}"""


def build_prompt(extracted, criteria, base_review):
    meta = extracted.get('_meta', {})
    rtype = meta.get('loai_bao_cao', 'hoan_thanh')
    decree = meta.get('nghi_dinh', 'nd06')
    default_max = (13 if decree == 'nd207' else 12) if rtype == 'hoan_thanh' else (9 if decree == 'nd207' else 8)
    max_muc = meta.get('so_muc_chuan', default_max)
    nd_text = 'NĐ 207/2026' if decree == 'nd207' else 'NĐ 06/2021'
    loai = (f'BÁO CÁO ĐỊNH KỲ ({max_muc} mục — Phụ lục IVa, {nd_text})' if rtype == 'dinh_ky'
            else f'BÁO CÁO HOÀN THÀNH ({max_muc} mục — Phụ lục IVb, {nd_text})')

    parts = [
        f"NHIỆM VỤ: Đánh giá chất lượng {loai} của đoàn TVGS theo {nd_text} và bộ tiêu chí "
        "nội bộ TEXO. Với TỪNG MỤC: đối chiếu nội dung thực tế với CHÚ THÍCH KIỂM TRA và "
        "CÂU HỎI TƯ DUY; chỉ ra cụ thể cái gì đạt, cái gì thiếu/sai (dẫn chứng ngắn). "
        "Đặc biệt chú ý: mâu thuẫn nội bộ giữa các mục; mục viết nhầm bản chất; nội dung "
        "chung chung không dẫn số văn bản; phạm vi thời gian (định kỳ = trong kỳ); phụ lục "
        "được dẫn chiếu có tồn tại không.",
        "",
        f"== KẾT QUẢ QUÉT QUY TẮC (điểm tựa, hãy kiểm chứng lại và bổ sung) ==",
        json.dumps({
            'sections_missing': meta.get('sections_missing', []),
            'phu_luc_found': [a['ten'] for a in meta.get('phu_luc_found', [])],
            'phu_luc_referenced': meta.get('phu_luc_referenced_in_body', []),
            'chu_ky': meta.get('chu_ky', {}),
            'phat_hien_rule': [p.get('noi_dung', '')[:160]
                               for p in base_review.get('phat_hien_chinh', [])],
            'trang_thai_rule': {m['muc']: m['trang_thai']
                                for m in base_review.get('danh_gia_muc', [])},
        }, ensure_ascii=False),
        "",
        "== TIÊU CHÍ & NỘI DUNG TỪNG MỤC ==",
    ]
    for i in range(1, max_muc + 1):
        key = f'muc_{i}'
        sc = criteria.get(key, {})
        text = extracted.get(key, '')
        if len(text) > MAX_CHARS_PER_MUC:
            text = text[:MAX_CHARS_PER_MUC] + '\n...[đã cắt bớt]'
        parts.append(f"--- MỤC {i}: {sc.get('ten', '')} ---")
        if sc.get('chu_thich_goc'):
            parts.append(f"[CHÚ THÍCH KIỂM TRA GỐC] {sc['chu_thich_goc'][:700]}")
        if sc.get('cau_hoi_tu_duy'):
            parts.append("[CÂU HỎI TƯ DUY] " + " | ".join(sc['cau_hoi_tu_duy'][:5]))
        parts.append("[NỘI DUNG TRONG BÁO CÁO]")
        parts.append(text if text.strip() else "(KHÔNG TÌM THẤY MỤC NÀY)")
        parts.append("")

    kl = extracted.get('ket_luan', '')
    if kl:
        parts.append("== PHẦN KẾT LUẬN CỦA BÁO CÁO (để đối chiếu chéo) ==")
        parts.append(kl[:1500])
    parts.append("")
    parts.append("Trả về DUY NHẤT JSON theo schema sau (đủ tất cả các mục 1..%d):" % max_muc)
    parts.append(SCHEMA)
    return "\n".join(parts)


def merge(base_review, ai_data, model_label=''):
    """Merge JSON AI trả về vào phiếu rule-based (giữ metadata file, thống kê tầng 1)."""
    out = json.loads(json.dumps(base_review))  # deep copy
    by_muc = {}
    for m in ai_data.get('danh_gia_muc', []) or []:
        if isinstance(m, dict) and m.get('muc'):
            by_muc[str(m['muc']).strip()] = m
    for m in out.get('danh_gia_muc', []):
        a = by_muc.get(m['muc'])
        if not a:
            continue
        st = str(a.get('trang_thai', '')).strip().upper()
        if st in VALID_ST:
            m['trang_thai'] = st
        if a.get('nhan_xet_ngan'):
            m['nhan_xet_ngan'] = str(a['nhan_xet_ngan'])[:140]
        if isinstance(a.get('dat_duoc'), list):
            m['dat_duoc'] = [str(x) for x in a['dat_duoc'] if x][:8]
        if isinstance(a.get('van_de'), list):
            m['van_de'] = [str(x) for x in a['van_de'] if x][:10]

    ph = ai_data.get('phat_hien_chinh')
    if isinstance(ph, list) and ph:
        out['phat_hien_chinh'] = [
            {'muc_do': (p.get('muc_do', 'CAN_SUA') if isinstance(p, dict) else 'CAN_SUA'),
             'noi_dung': (p.get('noi_dung', '') if isinstance(p, dict) else str(p))}
            for p in ph][:12]
    if isinstance(ai_data.get('kien_nghi'), list) and ai_data['kien_nghi']:
        out['kien_nghi'] = [str(x) for x in ai_data['kien_nghi'] if x][:12]
    if ai_data.get('ket_luan'):
        out['ket_luan'] = str(ai_data['ket_luan'])
    xl = str(ai_data.get('xep_loai', '')).strip().upper()
    if xl in VALID_XL:
        out['ket_qua_chung']['xep_loai'] = xl
    if ai_data.get('tom_tat'):
        out['ket_qua_chung']['tom_tat'] = (str(ai_data['tom_tat'])
            + f" (Phân tích sâu bằng AI{' — ' + model_label if model_label else ''}, "
              "trên nền kiểm tra quy tắc TEXO.)")
    # cập nhật thống kê theo trạng thái mới
    out['thong_ke']['loi'] = sum(1 for m in out['danh_gia_muc'] if m['trang_thai'] == 'LOI')
    out['thong_ke']['muc_thieu'] = sum(1 for m in out['danh_gia_muc'] if m['trang_thai'] == 'THIEU')
    return out


def ai_review(provider, model, key, base_url, extracted, criteria, base_review):
    """Trả (review_merged | None, error | None, truncated)."""
    prompt = build_prompt(extracted, criteria, base_review)
    res = llm.ask_json(provider, model, key, prompt, base_url)
    if not res.get('ok'):
        return None, res.get('error', 'Lỗi không xác định'), False
    merged = merge(base_review, res['data'], model_label=model)
    return merged, None, bool(res.get('truncated'))
