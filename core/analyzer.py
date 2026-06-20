# -*- coding: utf-8 -*-
"""
analyzer.py — Tầng phân tích KHÔNG dùng AI cho tvgs-checker-app.

Gồm:
- Tầng 1: kiểm tra từ khóa/bảng biểu bắt buộc theo criteria (như skill gốc)
- Tầng 1.5 (heuristic): các quy tắc đối chiếu chéo thay cho AI —
  mâu thuẫn nội bộ, nhầm bản chất mục 5, phụ lục dẫn chiếu thiếu, chữ ký,
  văn bản pháp lý hết hiệu lực, mục quá sơ sài, lỗi chính tả phổ biến,
  ghi chú nội bộ còn sót, yêu cầu riêng của báo cáo định kỳ.

Output: dict review_data tương thích generate_review_docx.generate().
"""

import re
import datetime

# Lỗi chính tả/đánh máy hay gặp trong báo cáo TVGS (regex, mô tả)
TYPOS = [
    (r'\btham giam\b', "'tham giam' → 'tham gia'"),
    (r'\blao dộng\b', "'lao dộng' → 'lao động'"),
    (r'\bdẩy nhanh\b', "'dẩy nhanh' → 'đẩy nhanh'"),
    (r'\bxây đựng\b', "'xây đựng' → 'xây dựng'"),
    (r'\bgiám xát\b', "'giám xát' → 'giám sát'"),
    (r'\bnghiêm thu\b', "'nghiêm thu' → 'nghiệm thu'"),
    (r'\bkhổi lượng\b', "'khổi lượng' → 'khối lượng'"),
    (r'\bNhà nhà thầu\b', "'Nhà nhà thầu' → 'Nhà thầu' (lặp từ)"),
    (r'\bcos\b ?±', "'cos ±' → 'cốt ±' (nên thống nhất 'cốt')"),
]

# Mức tối thiểu độ dài nội dung (ký tự) để không bị coi là sơ sài
MIN_LEN_DEFAULT = 150
MIN_LEN_EXCEPTIONS = {'muc_7': 60, 'muc_8': 60}  # các mục có thể ngắn hợp lệ nếu "Không"


def _check_entries(text, entries):
    """Trả về (passed[mô tả], missing[mô tả])."""
    passed, missing = [], []
    for e in entries or []:
        pattern = e['pattern'] if isinstance(e, dict) else e
        desc = e.get('mo_ta', e.get('ten', pattern)) if isinstance(e, dict) else pattern
        try:
            ok = bool(re.search(pattern, text, re.IGNORECASE))
        except re.error:
            ok = pattern.lower() in text.lower()
        (passed if ok else missing).append(desc)
    return passed, missing


def _grade(n_missing, n_total, has_loi, is_missing_section):
    if is_missing_section:
        return 'THIEU'
    if has_loi:
        return 'LOI'
    if n_missing == 0:
        return 'DAT'
    return 'CAN_SUA'


def analyze(extracted, criteria, nguoi_danh_gia_ten='', truong_phong_ten='Hoàng Đức Vũ',
            so_phieu=''):
    meta = extracted.get('_meta', {})
    rtype = meta.get('loai_bao_cao', 'hoan_thanh')
    decree = meta.get('nghi_dinh', 'nd06')
    default_max = (13 if decree == 'nd207' else 12) if rtype == 'hoan_thanh' else (9 if decree == 'nd207' else 8)
    max_muc = meta.get('so_muc_chuan', default_max)
    # Vị trí một số mục THAY ĐỔI theo nghị định (NĐ207 chèn mục) — suy ra động:
    #  - hoàn thành: "tồn tại" luôn ở mục 8; "điều kiện nghiệm thu" ở mục cuối (12 hoặc 13)
    #  - định kỳ:   "tồn tại" ở mục (cuối-1); "đề xuất/kiến nghị" ở mục cuối (8 hoặc 9)
    if rtype == 'hoan_thanh':
        key_tontai = 'muc_8'
        key_dieukien = f'muc_{max_muc}'      # 12 (NĐ06) | 13 (NĐ207)
        key_dexuat = None
    else:
        key_tontai = f'muc_{max_muc - 1}'    # 7 (NĐ06) | 8 (NĐ207)
        key_dieukien = None
        key_dexuat = f'muc_{max_muc}'        # 8 (NĐ06) | 9 (NĐ207)
    tbl_counts = meta.get('tables_per_section', {})
    full_text = '\n'.join(v for k, v in extracted.items() if not k.startswith('_'))
    body_keys = ['mo_dau', 'can_cu'] + [f'muc_{i}' for i in range(1, max_muc + 1)] + ['ket_luan']
    sec_names = {f"muc_{s['id']}": s['ten']
                 for s in criteria.get('cau_truc_muc', {}).get('sections', [])}

    total_pass = total_miss = 0
    findings = []          # phát hiện chính: (muc_do, noi_dung)
    extra_issues = {}      # muc_key -> [vấn đề heuristic]
    loi_flags = set()      # muc_key bị đánh LỖI bởi heuristic

    def add_issue(key, msg, loi=False):
        extra_issues.setdefault(key, []).append(msg)
        if loi:
            loi_flags.add(key)

    # ============ TẦNG 1.5 — HEURISTIC (thay cho AI) ============
    g = lambda k: extracted.get(k, '')

    # 1) Mâu thuẫn tồn tại giữa mục tồn tại và kết luận/mục điều kiện nghiệm thu
    t_tontai = g(key_tontai).lower()
    t_kl = ((g(key_dieukien) if key_dieukien else '') + ' ' + g('ket_luan')).lower()
    if re.search(r'không (còn|có)\s.{0,15}tồn tại', t_tontai) and \
       re.search(r'còn\s.{0,25}tồn tại', t_kl):
        msg = (f"MÂU THUẪN NỘI BỘ: {key_tontai.replace('_', ' ')} khẳng định 'không còn tồn tại' "
               "nhưng phần kết luận/điều kiện nghiệm thu lại nêu 'còn tồn tại cần khắc phục' — "
               "phải thống kê danh mục tồn tại và thống nhất giữa các mục.")
        add_issue(key_tontai, msg, loi=True)
        findings.append(('LOI', msg))

    # 2) Tiến độ tự mâu thuẫn ('đạt yêu cầu' + 'nguy cơ chậm' trong cùng câu/đoạn)
    t_td = g('muc_3') + ' ' + g('ket_luan')
    if re.search(r'đạt yêu cầu[^.\n]{0,120}(nguy cơ chậm|chậm tiến độ)', t_td, re.IGNORECASE):
        msg = ("Nhận định tiến độ TỰ MÂU THUẪN: vừa 'đạt yêu cầu' vừa 'có nguy cơ chậm' "
               "trong cùng một câu — phải viết thành một nhận định duy nhất kèm số liệu so sánh.")
        add_issue('muc_3', msg, loi=True)
        findings.append(('LOI', msg))

    # 3) Mục 5 (hoàn thành) nhầm bản chất sang thí nghiệm vật liệu
    if rtype == 'hoan_thanh' and g('muc_5'):
        m5 = g('muc_5')
        vl_hits = len(re.findall(r'lấy mẫu|chứng nhận chất lượng|vật liệu đầu vào|nguồn gốc xuất xứ|CO/CQ|vật tư', m5, re.IGNORECASE))
        kd_hits = len(re.findall(r'kiểm định|quan trắc|đối chứng', m5, re.IGNORECASE))
        if vl_hits >= 2 and vl_hits > kd_hits:
            msg = ("Mục 5 có dấu hiệu VIẾT NHẦM BẢN CHẤT: mô tả kiểm tra vật liệu thông thường "
                   "(lấy mẫu, CO/CQ — thuộc Mục 4) thay vì kiểm định công trình/máy móc, "
                   "quan trắc, thí nghiệm đối chứng.")
            add_issue('muc_5', msg, loi=True)
            findings.append(('LOI', msg))

    # 4) Đối chiếu phụ lục: mục dẫn chiếu phụ lục nào thì phụ lục đó phải có trong file.
    # LƯU Ý "CÁCH 2": nhiều đơn vị lập báo cáo kiểu ghi cơ bản + "Xem phụ lục X" —
    # đây là cách HỢP LỆ. Nội dung phụ lục sẽ được GỘP vào mục khi chấm (xem dưới).
    apx = meta.get('phu_luc_found', [])
    apx_by_so = {}            # số phụ lục -> key (phu_luc_1...)
    for a in apx:
        if a.get('so') and a['so'] not in apx_by_so:
            apx_by_so[a['so']] = a['key']
    refs_per_muc = meta.get('phu_luc_refs_per_muc', {})

    def _roman(n):
        vals = [(10, 'X'), (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')]
        out_s = ''
        while n > 0:
            for v, s in vals:
                if n >= v:
                    out_s += s; n -= v; break
        return out_s

    missing_refs_all = set()
    for sec_key, nums in refs_per_muc.items():
        for n in nums:
            if n not in apx_by_so:
                missing_refs_all.add(n)
                msg = (f"{sec_key.replace('_', ' ')} dẫn chiếu 'Phụ lục {_roman(n)}' nhưng "
                       "KHÔNG tìm thấy phụ lục này trong file — phải đính kèm hoặc bỏ dẫn chiếu.")
                add_issue(sec_key, msg, loi=True)
                findings.append(('LOI', msg))

    # 5) Chữ ký
    ck = meta.get('chu_ky', {})
    if not ck.get('giam_sat_truong'):
        msg = "Thiếu chữ ký GIÁM SÁT TRƯỞNG."
        findings.append(('LOI', msg))
    if rtype == 'hoan_thanh' and not ck.get('dai_dien_phap_luat'):
        msg = ("Báo cáo HOÀN THÀNH theo PL IVb phải có chữ ký NGƯỜI ĐẠI DIỆN THEO PHÁP LUẬT "
               "của tổ chức TVGS (ký, ghi rõ họ tên, chức vụ, đóng dấu) — chưa phát hiện trong file.")
        findings.append(('CAN_SUA', msg))

    # 6) Văn bản pháp lý hết hiệu lực
    if re.search(r'15/2021/NĐ-CP|Nghị định\s*(số\s*)?15/2021|NĐ\s*15/2021', full_text, re.IGNORECASE):
        msg = ("Dẫn chiếu NĐ 15/2021 — đã được THAY THẾ bởi NĐ 175/2024 "
               "(quy định chi tiết Luật Xây dựng về quản lý hoạt động xây dựng). Cần cập nhật.")
        findings.append(('CAN_SUA', msg)); add_issue('can_cu', msg)

    # 7) Lỗi chính tả phổ biến
    typo_found = []
    for pat, desc in TYPOS:
        if re.search(pat, full_text, re.IGNORECASE):
            typo_found.append(desc)
    if typo_found:
        findings.append(('CAN_SUA', 'Lỗi chính tả cần soát: ' + '; '.join(typo_found) + '.'))

    # 8) Ghi chú nội bộ còn sót
    note_cfg = criteria.get('pattern_ghi_chu_chua_xoa', {})
    notes = []
    for pat in note_cfg.get('patterns', []):
        try:
            rx = re.compile(pat['pattern'], re.IGNORECASE)
        except re.error:
            continue
        for k in body_keys:
            for line in g(k).split('\n'):
                ls = line.strip()
                if ls and len(ls) < 150 and rx.search(ls):
                    notes.append(f"[{k}] \"{ls[:80]}\"")
    if notes:
        findings.append(('CAN_SUA', 'Nghi còn GHI CHÚ NỘI BỘ chưa xóa: ' + ' | '.join(notes[:5])))

    # 9) Riêng báo cáo định kỳ
    if rtype == 'dinh_ky':
        dau = g('mo_dau') + ' ' + g('trang_bia')
        if not re.search(r'từ ngày.{0,40}đến (hết )?ngày', dau, re.IGNORECASE):
            if re.search(r'tháng\s*\d{1,2}\s*[/\-.]\s*\d{2,4}|quý\s*[IVX1-4]', dau, re.IGNORECASE):
                # ghi kỳ theo tháng/quý — chấp nhận được, chỉ nhắc nhẹ
                findings.append(('CAN_SUA',
                    "Kỳ báo cáo ghi theo tháng/quý — nên ghi rõ 'từ ngày... đến ngày...' "
                    "theo đúng mẫu Phụ lục IVa."))
            else:
                msg = "Báo cáo ĐỊNH KỲ bắt buộc ghi rõ KỲ BÁO CÁO (từ ngày... đến ngày...) — chưa thấy."
                findings.append(('LOI', msg)); add_issue('mo_dau', msg, loi=True)
        # phía trên có vấn đề mà mục đề xuất/kiến nghị (mục cuối) không kiến nghị
        has_van_de = bool(re.search(r'chậm tiến độ|chưa khắc phục|tồn tại', g('muc_3') + g(key_tontai), re.IGNORECASE)) \
            and not re.search(r'không có tồn tại|không còn tồn tại', g(key_tontai), re.IGNORECASE)
        if has_van_de and re.search(r'không có (đề xuất|kiến nghị)|chưa có kiến nghị', g(key_dexuat), re.IGNORECASE):
            msg = (f"VÔ LÝ: các mục phía trên nêu vấn đề (tồn tại/chậm tiến độ) nhưng Mục {max_muc} lại ghi "
                   "'không có đề xuất, kiến nghị' — phải có kiến nghị tương ứng.")
            findings.append(('LOI', msg)); add_issue(key_dexuat, msg, loi=True)

    # 10) Công trình cấp I mà mục 5 (kiểm định/quan trắc) ghi Không — CHỈ với báo cáo hoàn thành
    # (báo cáo định kỳ mục 5 là thống kê nghiệm thu, không liên quan quan trắc)
    if rtype == 'hoan_thanh' and re.search(r'cấp\s*I\b|công trình cấp 1\b', g('muc_1'), re.IGNORECASE):
        m5 = g('muc_5')
        if m5 and re.search(r'^|\n\s*[-+]?\s*Không\b', m5) and not re.search(r'quan trắc', m5, re.IGNORECASE):
            msg = ("Công trình CẤP I: quan trắc (lún/nghiêng/chuyển vị) thường BẮT BUỘC — "
                   "Mục 5 không thể hiện quan trắc và không nêu căn cứ miễn.")
            findings.append(('CAN_SUA', msg)); add_issue('muc_5', msg)

    # ============ TẦNG 1 — THEO TỪNG MỤC ============
    danh_gia_muc = []
    n_thieu = 0
    sections_missing = meta.get('sections_missing', [])
    for i in range(1, max_muc + 1):
        key = f'muc_{i}'
        sc = criteria.get(key, {})
        ten = sc.get('ten', sec_names.get(key, key))
        text = g(key)
        if key == 'mo_dau':
            text += '\n' + g('can_cu')
        is_miss = key in sections_missing or not text.strip()
        dat_duoc, van_de = [], list(extra_issues.get(key, []))

        if is_miss:
            n_thieu += 1
            van_de.insert(0, "KHÔNG tìm thấy mục này trong báo cáo — kiểm tra xem nội dung có bị "
                             "gộp vào mục khác không; nếu thiếu thật phải bổ sung (mẫu PL IV bắt buộc đủ mục).")
            danh_gia_muc.append({'muc': f'Mục {i}', 'ten': ten, 'trang_thai': 'THIEU',
                                 'nhan_xet_ngan': 'KHÔNG có mục riêng trong báo cáo',
                                 'dat_duoc': [], 'van_de': van_de})
            continue

        # "CÁCH 2": mục dẫn chiếu phụ lục → GỘP nội dung phụ lục vào mục khi chấm
        refs = refs_per_muc.get(key, [])
        resolved = [n for n in refs if n in apx_by_so]
        aug_text = text
        n_tbl = tbl_counts.get(key, 0)
        for n in resolved:
            ak = apx_by_so[n]
            aug_text += '\n' + g(ak)
            n_tbl += tbl_counts.get(ak, 0)
        if resolved:
            dat_duoc.append('Trình bày theo cách dẫn chiếu: chi tiết tại Phụ lục '
                            + ', '.join(_roman(n) for n in resolved)
                            + ' (đã kiểm tra gộp nội dung phụ lục).')

        passed, missing = _check_entries(aug_text, sc.get('tu_khoa_bat_buoc', []))
        tbl_passed, tbl_missing = _check_entries(aug_text, sc.get('bang_bieu_bat_buoc', []))
        total_pass += len(passed) + len(tbl_passed)
        total_miss += len(missing)

        if passed:
            dat_duoc.append('Đạt %d/%d tiêu chí từ khóa bắt buộc.' % (len(passed), len(passed) + len(missing)))
        if tbl_passed:
            dat_duoc.append('Có bảng biểu yêu cầu: ' + '; '.join(tbl_passed) + '.')
        elif n_tbl:
            dat_duoc.append(f'Mục có {n_tbl} bảng số liệu.')
        for m in missing:
            van_de.append('Thiếu nội dung bắt buộc: ' + m + '.')
        for m in tbl_missing:
            if n_tbl > 0:
                van_de.append(f'Có {n_tbl} bảng nhưng không khớp mẫu cột "{m}" — cần kiểm tra thủ công.')
            else:
                van_de.append('Thiếu bảng bắt buộc: ' + m + '.')

        # sơ sài? (KHÔNG cảnh báo nếu mục dẫn chiếu phụ lục — cách 2 hợp lệ)
        min_len = MIN_LEN_EXCEPTIONS.get(key, MIN_LEN_DEFAULT)
        if len(text) < min_len and not resolved and not re.search(r'Không\b', text):
            van_de.append(f'Nội dung rất ngắn ({len(text)} ký tự) — có dấu hiệu sơ sài, cần rà soát.')

        st = _grade(len(missing) + len(tbl_missing), len(passed) + len(missing),
                    key in loi_flags, False)
        if st == 'DAT' and len(van_de) > 0:
            st = 'CAN_SUA'
        nx = ('Đủ nội dung bắt buộc theo mẫu' if st == 'DAT'
              else (van_de[0][:90] if van_de else 'Cần rà soát'))
        danh_gia_muc.append({'muc': f'Mục {i}', 'ten': ten, 'trang_thai': st,
                             'nhan_xet_ngan': nx, 'dat_duoc': dat_duoc, 'van_de': van_de})

    # ============ TỔNG HỢP ============
    n_loi = sum(1 for m in danh_gia_muc if m['trang_thai'] == 'LOI') + \
        sum(1 for sev, _ in findings if sev == 'LOI' and True) - len(loi_flags & set(f'muc_{i}' for i in range(1, max_muc + 1)))
    n_loi = max(n_loi, sum(1 for sev, _ in findings if sev == 'LOI'))
    n_can_sua = sum(1 for m in danh_gia_muc if m['trang_thai'] == 'CAN_SUA')

    if n_thieu >= 3 or n_loi >= 5:
        xep_loai = 'KHÔNG ĐẠT'
    elif n_loi > 0 or n_thieu > 0:
        xep_loai = 'CẦN SỬA'
    elif n_can_sua > 0:
        xep_loai = 'ĐẠT CÓ ĐIỀU KIỆN'
    else:
        xep_loai = 'ĐẠT'

    # kiến nghị tự sinh
    kien_nghi = []
    for m in danh_gia_muc:
        if m['trang_thai'] in ('LOI', 'THIEU') and m['van_de']:
            kien_nghi.append(f"{m['muc']} ({m['ten']}): {m['van_de'][0]}")
    for sev, msg in findings:
        if sev == 'LOI' and not any(msg[:40] in k for k in kien_nghi):
            kien_nghi.append(msg)
    if typo_found:
        kien_nghi.append('Soát lỗi chính tả toàn văn trước khi trình ký.')
    kien_nghi = kien_nghi[:10]

    nd_text = 'NĐ 207/2026' if decree == 'nd207' else 'NĐ 06/2021'
    loai_text = (f'Báo cáo định kỳ ({max_muc} mục — Phụ lục IVa, {nd_text})' if rtype == 'dinh_ky'
                 else f'Báo cáo hoàn thành ({max_muc} mục — Phụ lục IVb, {nd_text})')
    tom_tat = (f"Kiểm tra tự động (không AI) theo mẫu {loai_text.split('(')[1][:-1]}: "
               f"{max_muc - n_thieu}/{max_muc} mục hiện diện; "
               f"{sum(1 for m in danh_gia_muc if m['trang_thai']=='DAT')} mục đạt, "
               f"{n_can_sua} mục cần sửa, "
               f"{sum(1 for m in danh_gia_muc if m['trang_thai']=='LOI')} mục lỗi, {n_thieu} mục thiếu. "
               "Kết quả do công cụ quét quy tắc — cần người có chuyên môn xác nhận lại trước khi phát hành.")
    ket_luan = (f"Báo cáo được Phòng Kỹ thuật kiểm tra tự động theo bộ tiêu chí TEXO "
                f"(xây dựng từ Phụ lục IV {nd_text} và chú thích kiểm tra nội bộ). "
                f"Xếp loại sơ bộ: {xep_loai}. "
                + ("Đề nghị đơn vị lập báo cáo xử lý các kiến nghị nêu trên và gửi lại Phòng Kỹ thuật "
                   "soát xét trước khi phát hành chính thức. " if xep_loai != 'ĐẠT' else
                   "Báo cáo cơ bản đáp ứng yêu cầu về cấu trúc và nội dung bắt buộc. ")
                + "Lưu ý: đây là kết quả kiểm tra bằng công cụ quy tắc, chưa thay thế "
                  "đánh giá chuyên môn của người kiểm tra.")

    review = {
        'so_phieu': so_phieu or f"Số: ...../{datetime.date.today().year}/PKT-ĐG",
        'tieu_de': ('ĐÁNH GIÁ BÁO CÁO ĐỊNH KỲ CÔNG TÁC TVGS' if rtype == 'dinh_ky'
                    else 'ĐÁNH GIÁ BÁO CÁO HOÀN THÀNH CÔNG TÁC TVGS'),
        'du_an': '', 'hang_muc_giai_doan': '', 'ky_bao_cao': '',
        'loai_bao_cao_text': loai_text,
        'file_bao_cao': meta.get('source_file', ''),
        'don_vi_lap': '', 'chu_dau_tu': '',
        'ngay_danh_gia': datetime.date.today().strftime('%d/%m/%Y'),
        'nguoi_danh_gia': 'Phòng Kỹ thuật — TEXO',
        'nguoi_danh_gia_ten': nguoi_danh_gia_ten,
        'truong_phong_ten': truong_phong_ten,
        'ket_qua_chung': {'xep_loai': xep_loai, 'tom_tat': tom_tat},
        'thong_ke': {'dat': total_pass, 'can_sua': total_miss,
                     'loi': sum(1 for sev, _ in findings if sev == 'LOI'),
                     'muc_thieu': n_thieu},
        'danh_gia_muc': danh_gia_muc,
        'phat_hien_chinh': [{'muc_do': sev, 'noi_dung': msg} for sev, msg in findings],
        'ket_luan': ket_luan,
    }
    return review


def guess_project_info(extracted):
    """Đoán tên dự án/hạng mục/CĐT từ trang bìa + mở đầu (regex, có thể sửa tay trên UI)."""
    src = (extracted.get('trang_bia', '') + '\n' + extracted.get('mo_dau', '')
           + '\n' + extracted.get('muc_1', ''))
    out = {}
    pats = {
        'du_an': r'(?:CÔNG TRÌNH|DỰ ÁN)\s*:\s*(.+)',
        'hang_muc_giai_doan': r'(?:GIAI ĐOẠN/HẠNG MỤC|HẠNG MỤC|GIAI ĐOẠN|GÓI THẦU)\s*:\s*(.+)',
        'chu_dau_tu': r'(?:CHỦ ĐẦU TƯ|Chủ đầu tư)\s*:\s*(.+)',
        'don_vi_lap': r'(?:TV GIÁM SÁT|ĐƠN VỊ TVGS|TVGS)\s*:\s*(.+)',
        'ky_bao_cao': r'(từ ngày[^\n]{5,60}đến (?:hết )?ngày[^\n]{0,20}|tháng\s*\d{1,2}\s*[/\-.]\s*\d{2,4})',
    }
    for k, p in pats.items():
        m = re.search(p, src, re.IGNORECASE)
        if m:
            out[k] = m.group(1).strip().rstrip('.;')[:160]
    return out
