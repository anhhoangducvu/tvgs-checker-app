#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_sections.py v2.0 — Trích xuất các mục từ báo cáo TVGS (.docx)

Cải tiến so với v1:
- Tự nhận diện LOẠI báo cáo: định kỳ (8 mục, PL IVa) hay hoàn thành (12 mục, PL IVb)
- Đọc đánh số TỰ ĐỘNG của Word (numbering.xml): số thường (1,2,3), số La Mã (I,II,III),
  chữ cái (a,b,c) — kể cả khi số không nằm trong text
- Bỏ qua trang bìa / trang lót, tách riêng vào key "trang_bia"
- Đọc nội dung TRONG BẢNG (text trong bảng được gắn vào mục chứa nó, đếm số bảng/mục)
- Tách phần PHỤ LỤC riêng (không bỏ qua), phát hiện dẫn chiếu phụ lục trong thân
- Phát hiện khối chữ ký (Giám sát trưởng / Người đại diện pháp luật)

Usage:
    python extract_sections.py <file.docx> [--type dinh_ky|hoan_thanh] [--output-json] [--output-file out.json]

Với PDF: cần pdftotext (poppler), hoặc dùng skill pdf-to-word để convert trước.
"""

import sys, os, re, json, argparse, zipfile, subprocess
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


# ---------------------------------------------------------------- numbering ---

def int_to_roman(n):
    vals = [(1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),(100,'C'),(90,'XC'),
            (50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')]
    out = ''
    for v, s in vals:
        while n >= v:
            out += s; n -= v
    return out

def roman_to_int(s):
    vals = {'I':1,'V':5,'X':10,'L':50,'C':100,'D':500,'M':1000}
    s = s.upper(); total = 0; prev = 0
    for ch in reversed(s):
        v = vals.get(ch, 0)
        if v < prev: total -= v
        else: total += v; prev = v
    return total

def fmt_num(fmt, n):
    if fmt == 'decimal': return str(n)
    if fmt == 'upperRoman': return int_to_roman(n)
    if fmt == 'lowerRoman': return int_to_roman(n).lower()
    if fmt == 'upperLetter': return chr(ord('A') + (n-1) % 26)
    if fmt == 'lowerLetter': return chr(ord('a') + (n-1) % 26)
    if fmt in ('bullet', 'none'): return ''
    return str(n)


class Numbering:
    """Resolve Word automatic numbering (numId/ilvl -> rendered label like 'I.' or '2.1.')."""

    def __init__(self, zf):
        self.num2abs = {}; self.levels = {}; self.overrides = {}
        self.counters = {}
        if 'word/numbering.xml' not in zf.namelist():
            return
        root = ET.fromstring(zf.read('word/numbering.xml'))
        for an in root.findall(W+'abstractNum'):
            aid = an.get(W+'abstractNumId'); lvls = {}
            for lvl in an.findall(W+'lvl'):
                il = int(lvl.get(W+'ilvl'))
                g = lambda tag, d=None: (lvl.find(W+tag).get(W+'val')
                                          if lvl.find(W+tag) is not None else d)
                lvls[il] = {'fmt': g('numFmt','decimal'), 'text': g('lvlText','%1.'),
                            'start': int(g('start','1') or 1)}
            self.levels[aid] = lvls
        for num in root.findall(W+'num'):
            nid = num.get(W+'numId')
            a = num.find(W+'abstractNumId')
            self.num2abs[nid] = a.get(W+'val') if a is not None else None
            for ov in num.findall(W+'lvlOverride'):
                il = int(ov.get(W+'ilvl'))
                so = ov.find(W+'startOverride')
                if so is not None:
                    self.overrides[(nid, il)] = int(so.get(W+'val'))

    def label(self, num_id, ilvl):
        aid = self.num2abs.get(num_id)
        if aid is None or aid not in self.levels:
            return ''
        lvls = self.levels[aid]
        if ilvl not in lvls:
            return ''
        if lvls[ilvl]['fmt'] in ('bullet', 'none'):
            return lvls[ilvl]['text'].strip() if lvls[ilvl]['fmt'] == 'bullet' else ''
        key = aid  # share counters across numIds pointing to same abstract
        cnt = self.counters.setdefault(key, {})
        start = self.overrides.get((num_id, ilvl), lvls[ilvl]['start'])
        cnt[ilvl] = cnt.get(ilvl, start - 1) + 1
        for deeper in list(cnt):
            if deeper > ilvl:
                del cnt[deeper]
        text = lvls[ilvl]['text']
        for m in range(1, 9):
            if f'%{m}' in text:
                lvl_i = m - 1
                v = cnt.get(lvl_i, self.overrides.get((num_id, lvl_i),
                            lvls.get(lvl_i, {}).get('start', 1)))
                text = text.replace(f'%{m}', fmt_num(lvls.get(lvl_i, {}).get('fmt','decimal'), v))
        return text


def style_numbering_map(zf):
    """Map style id -> (numId, ilvl) for styles that carry numbering."""
    out = {}
    if 'word/styles.xml' not in zf.namelist():
        return out
    root = ET.fromstring(zf.read('word/styles.xml'))
    for st in root.findall(W+'style'):
        sid = st.get(W+'styleId')
        ppr = st.find(W+'pPr')
        if ppr is None: continue
        npr = ppr.find(W+'numPr')
        if npr is None: continue
        nid = npr.find(W+'numId'); il = npr.find(W+'ilvl')
        if nid is not None:
            out[sid] = (nid.get(W+'val'), int(il.get(W+'val')) if il is not None else 0)
    return out


# ---------------------------------------------------------------- docx walk ---

def para_text(p):
    return ''.join(t.text or '' for t in p.iter(W+'t')).strip()

def walk_docx(path):
    """Yield items in document order:
       {'kind':'p','text':rendered,'raw':text,'label':label,'ilvl':int|None,'caps':bool}
       {'kind':'table','rows':n,'cols':n,'text':flattened}"""
    zf = zipfile.ZipFile(path)
    numbering = Numbering(zf)
    style_num = style_numbering_map(zf)
    root = ET.fromstring(zf.read('word/document.xml'))
    body = root.find(W+'body')
    items = []
    for child in list(body):
        if child.tag == W+'p':
            text = para_text(child)
            label = ''; ilvl = None
            ppr = child.find(W+'pPr')
            num_id = None
            if ppr is not None:
                npr = ppr.find(W+'numPr')
                if npr is not None:
                    nid = npr.find(W+'numId'); il = npr.find(W+'ilvl')
                    if nid is not None:
                        num_id = nid.get(W+'val')
                        ilvl = int(il.get(W+'val')) if il is not None else 0
                if num_id is None:
                    ps = ppr.find(W+'pStyle')
                    if ps is not None and ps.get(W+'val') in style_num:
                        num_id, ilvl = style_num[ps.get(W+'val')]
            if num_id is not None and text:
                label = numbering.label(num_id, ilvl)
            if not text:
                continue
            rendered = (label + ' ' + text).strip() if label else text
            letters = [c for c in text if c.isalpha()]
            caps = bool(letters) and sum(1 for c in letters if c.isupper()) / len(letters) > 0.85
            items.append({'kind': 'p', 'text': rendered, 'raw': text,
                          'label': label, 'ilvl': ilvl, 'caps': caps})
        elif child.tag == W+'tbl':
            rows = child.findall(W+'tr')
            lines = []
            for tr in rows:
                cells = []
                for tc in tr.findall(W+'tc'):
                    ct = ' '.join(para_text(p) for p in tc.findall(W+'p') if para_text(p))
                    cells.append(ct)
                dedup = []
                for c in cells:
                    if not dedup or c != dedup[-1]:
                        dedup.append(c)
                line = ' | '.join(c for c in dedup if c)
                if line:
                    lines.append(line)
            ncols = len(rows[0].findall(W+'tc')) if rows else 0
            items.append({'kind': 'table', 'rows': len(rows), 'cols': ncols,
                          'text': '\n'.join(lines)})
    return items


def _pdf_text_pdftotext(path):
    """Đọc PDF bằng poppler (pdftotext -layout) — tốt nhất nếu có cài."""
    r = subprocess.run(['pdftotext', '-layout', path, '-'],
                       capture_output=True, timeout=120)
    if r.returncode != 0:
        raise RuntimeError(r.stderr.decode('utf-8', 'ignore')[:200] or 'pdftotext failed')
    return r.stdout.decode('utf-8', 'ignore')


def _pdf_text_pypdf(path):
    """Fallback thuần Python (pypdf) — chạy được mọi nơi, không cần poppler."""
    import logging
    logging.getLogger('pypdf').setLevel(logging.ERROR)  # ém warning PDF không chuẩn
    try:
        from pypdf import PdfReader
    except ImportError:
        from PyPDF2 import PdfReader  # tên gói cũ
    reader = PdfReader(path)
    return '\n'.join((page.extract_text() or '') for page in reader.pages)


def walk_pdf(path):
    """Đọc PDF: thử pdftotext (poppler) trước, không có thì tự chuyển sang pypdf."""
    txt, errs = None, []
    for fn in (_pdf_text_pdftotext, _pdf_text_pypdf):
        try:
            txt = fn(path)
            if txt and txt.strip():
                break
            errs.append(f'{fn.__name__}: không trích xuất được text')
            txt = None
        except Exception as e:
            errs.append(f'{fn.__name__}: {str(e)[:120]}')
    if txt is None:
        raise RuntimeError(
            'Không đọc được PDF (' + ' | '.join(errs) + '). '
            'PDF dạng scan/ảnh không trích xuất được chữ — hãy convert sang .docx rồi tải lại. '
            'Nếu chạy trên máy cá nhân: pip install pypdf.')
    items = []
    for line in txt.split('\n'):
        t = ' '.join(line.split())
        if not t: continue
        letters = [c for c in t if c.isalpha()]
        caps = bool(letters) and sum(1 for c in letters if c.isupper()) / len(letters) > 0.85
        items.append({'kind': 'p', 'text': t, 'raw': t, 'label': '', 'ilvl': None, 'caps': caps})
    return items


# ------------------------------------------------------------ heading logic ---

# =====================================================================
# Từ khóa nhận diện tiêu đề mục — THEO LOẠI BÁO CÁO và NGHỊ ĐỊNH.
# Hỗ trợ CẢ HAI nghị định (dùng song song trong giai đoạn chuyển tiếp):
#   - NĐ 06/2021: định kỳ 8 mục (IVa), hoàn thành 12 mục (IVb)
#   - NĐ 207/2026: định kỳ 9 mục (IVa), hoàn thành 13 mục (IVb)
# =====================================================================

# --- HOÀN THÀNH — NĐ 06/2021 (12 mục) ---
KW_HOAN_THANH_06 = {
    1:  [r'quy mô(,| và| công trình| thông tin|.{0,15}công năng)', r'quy mô công trình'],
    2:  [r'năng lực.{0,40}nhà thầu', r'^nhà thầu thi công xây dựng\s*:?$', r'sự phù hợp về năng lực'],
    3:  [r'khối lượng.{0,30}tiến độ', r'tình hình thi công', r'khối lượng, tiến độ'],
    4:  [r'thí nghiệm.{0,40}(kiểm tra )?vật liệu', r'vật liệu.{0,30}thiết bị lắp đặt', r'công tác thí nghiệm'],
    5:  [r'kiểm định.{0,30}quan trắc', r'quan trắc.{0,40}(thí nghiệm )?đối chứng', r'kiểm định, quan trắc'],
    6:  [r'nghiệm thu công việc', r'tổ chức nghiệm thu'],
    7:  [r'thay đổi thiết kế', r'sửa đổi thiết kế', r'thiết kế điều chỉnh'],
    8:  [r'tồn tại.{0,25}khiếm khuyết', r'khiếm khuyết về chất lượng', r'tồn tại khiếm khuyết'],
    9:  [r'hồ sơ quản lý chất lượng', r'hồ sơ QLCL'],
    10: [r'tuân thủ.{0,40}pháp luật', r'môi trường.{0,30}(PCCC|phòng cháy)', r'pháp luật về môi trường'],
    11: [r'quy trình vận hành', r'vận hành.{0,25}bảo trì'],
    12: [r'điều kiện nghiệm thu (hoàn thành|.{0,30}(giai đoạn|gói thầu|hạng mục|công trình))',
         r'đánh giá về các điều kiện nghiệm thu'],
}

# --- HOÀN THÀNH — NĐ 207/2026 (13 mục): chèn mục 10 PCCC, tách môi trường ---
KW_HOAN_THANH_207 = {
    1:  KW_HOAN_THANH_06[1],
    2:  KW_HOAN_THANH_06[2],
    3:  KW_HOAN_THANH_06[3],
    4:  KW_HOAN_THANH_06[4],
    5:  KW_HOAN_THANH_06[5],
    6:  KW_HOAN_THANH_06[6],
    7:  KW_HOAN_THANH_06[7],
    8:  KW_HOAN_THANH_06[8],
    9:  KW_HOAN_THANH_06[9],
    10: [r'(phòng cháy|PCCC).{0,50}(thẩm định|phê duyệt|thẩm duyệt)',
         r'thi công.{0,40}thiết kế.{0,25}(về )?(phòng cháy|PCCC)',
         r'thiết kế.{0,20}(về )?(phòng cháy|PCCC).{0,20}(được )?(thẩm định|phê duyệt|thẩm duyệt)',
         r'(công tác )?phòng cháy.{0,15}(và )?chữa cháy'],
    11: [r'(bảo vệ )?môi trường', r'pháp luật.{0,20}môi trường'],
    12: [r'quy trình vận hành', r'vận hành.{0,25}bảo trì'],
    13: [r'điều kiện nghiệm thu (hoàn thành|.{0,30}(giai đoạn|gói thầu|hạng mục|công trình))',
         r'đánh giá về các điều kiện nghiệm thu'],
}

# --- ĐỊNH KỲ — NĐ 06/2021 (8 mục) ---
KW_DINH_KY_06 = {
    1: [r'(sự )?phù hợp về quy mô.{0,15}công năng', r'quy mô.{0,15}công năng'],
    2: [r'năng lực.{0,40}nhà thầu', r'sự phù hợp về năng lực'],
    3: [r'khối lượng.{0,30}tiến độ', r'tình hình thi công'],
    4: [r'(thống kê.{0,40})?thí nghiệm', r'kiểm tra vật liệu'],
    5: [r'(thống kê.{0,40})?nghiệm thu', r'nghiệm thu trong kỳ', r'nghiệm thu giai đoạn'],
    6: [r'thay đổi thiết kế'],
    7: [r'tồn tại.{0,25}khiếm khuyết', r'khiếm khuyết về chất lượng'],
    8: [r'đề xuất.{0,15}kiến nghị', r'kiến nghị'],
}

# --- ĐỊNH KỲ — NĐ 207/2026 (9 mục): chèn mục 5 'đối chiếu KQ thí nghiệm với thiết kế' ---
KW_DINH_KY_207 = {
    1: KW_DINH_KY_06[1],
    2: KW_DINH_KY_06[2],
    3: KW_DINH_KY_06[3],
    4: [r'(thống kê.{0,40})?(công tác )?thí nghiệm.{0,40}(được thực hiện|trong kỳ|kiểm soát)',
        r'kiểm soát chất lượng.{0,20}thí nghiệm', r'(thống kê.{0,30})?thí nghiệm', r'kiểm tra vật liệu'],
    5: [r'(sự )?phù hợp.{0,60}(kết quả thí nghiệm|quan trắc|kiểm định).{0,60}(so với|yêu cầu).{0,15}thiết kế',
        r'(kết quả thí nghiệm|quan trắc|kiểm định).{0,60}so với.{0,15}(yêu cầu )?thiết kế',
        r'so với yêu cầu thiết kế'],
    6: [r'(thống kê.{0,40})?nghiệm thu', r'nghiệm thu trong kỳ', r'nghiệm thu giai đoạn'],
    7: [r'thay đổi thiết kế'],
    8: [r'tồn tại.{0,25}khiếm khuyết', r'khiếm khuyết về chất lượng'],
    9: [r'đề xuất.{0,15}kiến nghị', r'kiến nghị'],
}

# alias để giữ tương thích phần nhận diện loại báo cáo (dùng bản NĐ06 làm cơ sở)
KW_HOAN_THANH = KW_HOAN_THANH_06
KW_DINH_KY = KW_DINH_KY_06

# thứ tự ưu tiên khi khớp từ khóa (mục đặc thù trước, mục chung chung sau) — theo (loại, nghị định)
KW_PRIORITY = {
    ('hoan_thanh', 'nd06'):  [12, 11, 10, 9, 5, 7, 8, 4, 6, 2, 3, 1],
    ('hoan_thanh', 'nd207'): [13, 12, 11, 10, 9, 5, 8, 7, 6, 4, 2, 3, 1],
    ('dinh_ky', 'nd06'):     [8, 6, 7, 5, 4, 2, 3, 1],
    ('dinh_ky', 'nd207'):    [9, 7, 8, 6, 5, 4, 2, 3, 1],
    # alias chỉ-theo-loại (mặc định NĐ06) cho phần nhận diện loại báo cáo
    'hoan_thanh': [12, 11, 10, 9, 5, 7, 8, 4, 6, 2, 3, 1],
    'dinh_ky':    [8, 6, 7, 5, 4, 2, 3, 1],
}

# tra cứu (bản đồ từ khóa, ưu tiên, số mục chuẩn) theo loại + nghị định
def kw_for(report_type, decree):
    decree = decree if decree in ('nd06', 'nd207') else 'nd06'
    if report_type == 'hoan_thanh':
        kw = KW_HOAN_THANH_207 if decree == 'nd207' else KW_HOAN_THANH_06
        max_muc = 13 if decree == 'nd207' else 12
    else:
        kw = KW_DINH_KY_207 if decree == 'nd207' else KW_DINH_KY_06
        max_muc = 9 if decree == 'nd207' else 8
    return kw, KW_PRIORITY[(report_type, decree)], max_muc

RE_CAN_CU  = re.compile(r'^(các\s+)?căn cứ', re.IGNORECASE)
RE_KET_LUAN = re.compile(r'^kết luận', re.IGNORECASE)
RE_PHU_LUC = re.compile(r'^phụ\s*lục\s*([IVX0-9]+)?', re.IGNORECASE)
RE_HINH_ANH = re.compile(r'(một số )?hình ảnh', re.IGNORECASE)
RE_NUM_HEAD = re.compile(r'^\s*(\d{1,2})\s*[\.\):/](?!\d)\s*')   # bắt cả dạng "4/ ..."
RE_ROMAN_HEAD = re.compile(r'^\s*([IVX]{1,4})\s*[\.\):/]\s+')
RE_MUC_HEAD = re.compile(r'^\s*Mục\s+(\d{1,2})\b', re.IGNORECASE)
RE_TOC_LINE = re.compile(r'\.{5,}\s*\d*\s*$')   # dòng MỤC LỤC: "..... 12"


def heading_ordinal(item):
    """Trả về số thứ tự mục nếu paragraph trông như tiêu đề mục cấp 1, else None."""
    t = item['text']
    if len(t) > 220 or RE_TOC_LINE.search(t):
        return None
    if item['label'] and item.get('ilvl') == 0:
        lm = re.match(r'^([IVX]{1,4})[\.\)]?$', item['label'].strip())
        if lm:
            return roman_to_int(lm.group(1))
        dm = re.match(r'^(\d{1,2})[\.\)]?$', item['label'].strip())
        if dm:
            return int(dm.group(1))
    m = RE_MUC_HEAD.match(t) or RE_NUM_HEAD.match(t)
    if m:
        return int(m.group(1))
    m = RE_ROMAN_HEAD.match(t)
    if m:
        return roman_to_int(m.group(1))
    return None


def ordinal_format(item):
    """('roman'|'decimal', n) nếu paragraph có số thứ tự cấp 1, else (None, None)."""
    if item['label'] and item.get('ilvl') == 0:
        lab = item['label'].strip()
        m = re.match(r'^([IVX]{1,4})[\.\)]?$', lab)
        if m: return 'roman', roman_to_int(m.group(1))
        m = re.match(r'^(\d{1,2})[\.\)]?$', lab)
        if m: return 'decimal', int(m.group(1))
    t = item['text']
    m = RE_ROMAN_HEAD.match(t)
    if m: return 'roman', roman_to_int(m.group(1))
    m = RE_MUC_HEAD.match(t) or RE_NUM_HEAD.match(t)
    if m: return 'decimal', int(m.group(1))
    return None, None


def strip_leading_number(t):
    t = RE_MUC_HEAD.sub('', t)
    t = RE_NUM_HEAD.sub('', t)
    t = RE_ROMAN_HEAD.sub('', t)
    return t.strip()


def match_keywords(text, kw_map, priority):
    body = strip_leading_number(text)
    for muc in priority:
        for p in kw_map.get(muc, []):
            if re.search(p, body, re.IGNORECASE):
                return muc
    return None


def is_heading_candidate(item):
    t = item['text']
    if item['kind'] != 'p' or not t or len(t) > 220:
        return False
    if RE_TOC_LINE.search(t):   # dòng mục lục không phải tiêu đề thật
        return False
    if heading_ordinal(item) is not None:
        return True
    if item['caps'] and len(t) < 150:
        return True
    return False


# ------------------------------------------------------------ type detection ---

def detect_type(items, forced=None):
    signals = []
    if forced:
        return forced, 99, [f'--type {forced} (do người dùng chỉ định)']
    full = '\n'.join(it['text'] for it in items)
    head = '\n'.join(it['text'] for it in items[:80])
    score = {'dinh_ky': 0, 'hoan_thanh': 0}
    if re.search(r'BÁO CÁO ĐỊNH KỲ', head, re.IGNORECASE):
        score['dinh_ky'] += 10; signals.append("Tiêu đề chứa 'BÁO CÁO ĐỊNH KỲ'")
    if re.search(r'BÁO CÁO HOÀN THÀNH|HOÀN THÀNH CÔNG TÁC.{0,20}GIÁM SÁT', head, re.IGNORECASE):
        score['hoan_thanh'] += 10; signals.append("Tiêu đề chứa 'BÁO CÁO HOÀN THÀNH'")
    n_ky = len(re.findall(r'trong kỳ báo cáo|kỳ báo cáo này', full, re.IGNORECASE))
    if n_ky >= 3:
        score['dinh_ky'] += 2; signals.append(f"'trong kỳ báo cáo' xuất hiện {n_ky} lần")
    if re.search(r'từ ngày.{0,40}đến (hết )?ngày', head, re.IGNORECASE):
        score['dinh_ky'] += 1; signals.append("Mở đầu có 'từ ngày... đến ngày...' (kỳ báo cáo)")
    for it in items:
        if is_heading_candidate(it):
            if match_keywords(it['text'], KW_HOAN_THANH, KW_PRIORITY['hoan_thanh']) in (9, 10, 11, 12):
                score['hoan_thanh'] += 2
            if re.search(r'đề xuất.{0,15}kiến nghị', it['text'], re.IGNORECASE):
                score['dinh_ky'] += 1
    if re.search(r'điều kiện nghiệm thu (hoàn thành|.{0,30}(giai đoạn|gói thầu|hạng mục))|đủ điều kiện nghiệm thu', full, re.IGNORECASE):
        score['hoan_thanh'] += 2; signals.append("Có nội dung 'điều kiện nghiệm thu hoàn thành' (mục 12 IVb)")
    if score['hoan_thanh'] == score['dinh_ky']:
        return 'unknown', 0, signals
    typ = 'hoan_thanh' if score['hoan_thanh'] > score['dinh_ky'] else 'dinh_ky'
    return typ, abs(score['hoan_thanh'] - score['dinh_ky']), signals


def detect_decree(items, report_type, forced=None):
    """Nhận diện NGHỊ ĐỊNH áp dụng: 'nd06' (06/2021) hay 'nd207' (207/2026).
    Cả hai mẫu dùng song song trong giai đoạn chuyển tiếp nên phải tự đoán."""
    if forced in ('nd06', 'nd207'):
        return forced, 99, [f'--nghidinh {forced} (do người dùng chỉ định)']
    full = '\n'.join(it['text'] for it in items)
    signals = []
    score = {'nd207': 0, 'nd06': 0}

    # 1) Trích dẫn số nghị định trong căn cứ pháp lý (tín hiệu mạnh nhất)
    if re.search(r'207\s*/\s*2026', full):
        score['nd207'] += 6; signals.append("Dẫn chiếu 'Nghị định 207/2026'")
    if re.search(r'\b06\s*/\s*2021', full):
        score['nd06'] += 4; signals.append("Dẫn chiếu 'Nghị định 06/2021'")

    # 2) Cấu trúc đặc trưng theo tiêu đề mục
    headings = [it for it in items if is_heading_candidate(it)]
    if report_type == 'hoan_thanh':
        for it in headings:
            b = strip_leading_number(it['text'])
            if (re.search(r'(phòng cháy|PCCC)', b, re.IGNORECASE) and
                    re.search(r'thẩm định|phê duyệt|thẩm duyệt|thiết kế', b, re.IGNORECASE) and
                    not re.search(r'môi trường', b, re.IGNORECASE)):
                score['nd207'] += 4
                signals.append("Có mục RIÊNG về thi công theo thiết kế PCCC được thẩm duyệt (đặc trưng NĐ207)")
                break
        if any(heading_ordinal(it) == 13 for it in headings):
            score['nd207'] += 3; signals.append("Có mục số 13 (NĐ207 hoàn thành = 13 mục)")
        for it in headings:
            b = strip_leading_number(it['text'])
            if re.search(r'môi trường', b, re.IGNORECASE) and re.search(r'phòng cháy|PCCC', b, re.IGNORECASE):
                score['nd06'] += 3
                signals.append("Có mục GỘP 'môi trường + PCCC' (đặc trưng NĐ06)")
                break
    else:  # dinh_ky
        if any(heading_ordinal(it) == 9 for it in headings):
            score['nd207'] += 3; signals.append("Có mục số 9 (NĐ207 định kỳ = 9 mục)")
        for it in headings:
            b = strip_leading_number(it['text'])
            if re.search(r'(kết quả thí nghiệm|quan trắc|kiểm định).{0,70}(so với|yêu cầu).{0,15}thiết kế'
                         r'|so với yêu cầu thiết kế', b, re.IGNORECASE):
                score['nd207'] += 3
                signals.append("Có mục đối chiếu kết quả thí nghiệm với thiết kế (đặc trưng NĐ207)")
                break

    if score['nd207'] == 0 and score['nd06'] == 0:
        return 'nd06', 0, signals + [
            'Không có dấu hiệu rõ ràng — MẶC ĐỊNH NĐ 06/2021 (mẫu cũ còn phổ biến trong giai đoạn '
            'chuyển tiếp). Nếu là mẫu NĐ 207/2026, hãy chọn/ép thủ công.']
    if score['nd207'] > score['nd06']:
        return 'nd207', score['nd207'] - score['nd06'], signals
    if score['nd06'] > score['nd207']:
        return 'nd06', score['nd06'] - score['nd207'], signals
    return 'nd06', 0, signals + ['Hai nghị định ngang điểm — mặc định NĐ 06/2021, kiểm tra lại thủ công.']


# ------------------------------------------------------------ main parsing ---

def find_body_start(items):
    """Trang bìa/lót: mọi thứ trước khối tiêu đề chính thức (BÁO CÁO... ngay trước 'Kính gửi')."""
    kg = None
    for i, it in enumerate(items):
        if it['kind'] == 'p' and re.match(r'^Kính gửi', it['raw'], re.IGNORECASE):
            kg = i; break
    if kg is None:
        return 0
    start = kg
    for j in range(kg - 1, max(-1, kg - 12), -1):
        it = items[j]
        if it['kind'] == 'p' and ('BÁO CÁO' in it['raw'].upper() or 'CỘNG HÒA' in it['raw'].upper()):
            start = j
        elif it['kind'] == 'table' and 'CỘNG HÒA' in it['text'].upper():
            start = j  # bảng letterhead quốc hiệu
    return start


def parse(items, report_type, decree='nd06'):
    kw_map, priority, max_muc = kw_for(report_type, decree)

    body_start = find_body_start(items)
    sections = {'trang_bia': [], 'mo_dau': []}
    tables_count = {}
    appendices = []      # [{ten, key}]
    current = 'mo_dau'
    in_appendix = False
    last_muc = 0
    fmt_seen = set()     # định dạng số của các tiêu đề mục đã khớp ('roman'/'decimal')
    has_gst_sign = False
    has_daidien_sign = False

    for idx, it in enumerate(items):
        txt_for_sign = it['text']
        if re.search(r'GIÁM SÁT TRƯỞNG|TVGS TRƯỞNG|T\.?GIÁM SÁT', txt_for_sign, re.IGNORECASE):
            has_gst_sign = True
        if re.search(r'ĐẠI DIỆN|TỔNG GIÁM ĐỐC|GIÁM ĐỐC', txt_for_sign) and 'NHÀ THẦU' not in txt_for_sign.upper():
            has_daidien_sign = True

        if idx < body_start:
            sections['trang_bia'].append(it['text'] if it['kind'] == 'p'
                                          else '[BẢNG] ' + it['text'].replace('\n', ' / '))
            continue

        if it['kind'] == 'table':
            sections.setdefault(current, []).append(f"[BẢNG {it['rows']}x{it['cols']}]\n{it['text']}")
            tables_count[current] = tables_count.get(current, 0) + 1
            continue

        t = it['text']
        # --- phụ lục / hình ảnh ---
        if it['kind'] == 'p' and len(t) < 150 and (RE_PHU_LUC.match(t.strip()) or
                (RE_HINH_ANH.search(t) and it['caps'])):
            in_appendix = True
            key = f'phu_luc_{len(appendices) + 1}'
            appendices.append({'ten': t.strip(), 'key': key})
            current = key
            sections.setdefault(current, []).append(t)
            continue

        if not in_appendix and is_heading_candidate(it):
            stripped = strip_leading_number(t)
            if RE_CAN_CU.match(stripped):
                current = 'can_cu'
                sections.setdefault(current, []).append(t)
                continue
            if RE_KET_LUAN.match(stripped):
                # với báo cáo định kỳ, "Kết luận" thực chất là mục cuối (đề xuất/kiến nghị)
                current = f'muc_{max_muc}' if (report_type == 'dinh_ky') else 'ket_luan'
                sections.setdefault(current, []).append(t)
                last_muc = max_muc if report_type == 'dinh_ky' else last_muc
                continue
            ofmt, ordn = ordinal_format(it)
            if fmt_seen and ofmt and ofmt not in fmt_seen:
                # kiểu đánh số khác với tiêu đề cấp 1 đã xác lập (vd cấp 1 là I,II,III
                # mà đây là 1.,2.,3.) → đây là TIÊU ĐỀ CON, không đổi mục
                muc = None
            else:
                muc = match_keywords(t, kw_map, priority)
                if muc is not None:
                    if ofmt:
                        fmt_seen.add(ofmt)
                elif (ordn is not None and 1 <= ordn <= max_muc and ordn == last_muc + 1):
                    # fallback theo số thứ tự khi không khớp từ khóa, đúng trình tự
                    muc = ordn
            if muc is not None:
                current = f'muc_{muc}'
                last_muc = max(last_muc, muc)
                sections.setdefault(current, []).append(t)
                continue

        sections.setdefault(current, []).append(t)

    out = {}
    for k, lines in sections.items():
        text = '\n'.join(lines).strip()
        if text:
            out[k] = text

    expected = [f'muc_{i}' for i in range(1, max_muc + 1)]
    found = [k for k in expected if k in out]
    missing = [k for k in expected if k not in out]

    body_text = '\n'.join(out.get(k, '') for k in ['mo_dau', 'can_cu'] + expected + ['ket_luan'])
    refs = sorted(set(m.group(0).strip() for m in
                      re.finditer(r'[Pp]hụ lục\s+[IVX0-9]+', body_text)))

    # --- đánh SỐ cho từng phụ lục (từ tiêu đề; nếu tiêu đề bị cắt dòng thì nội suy theo thứ tự) ---
    def _ref_num(s):
        s = s.strip().upper()
        return int(s) if s.isdigit() else roman_to_int(s)

    prev_so = 0
    for a in appendices:
        m = re.search(r'PHỤ\s*LỤC\s*[:\.]?\s*([IVX]+\b|\d{1,2})', a['ten'].upper())
        if m is None:  # tiêu đề không có số (vd 'MỘT SỐ HÌNH ẢNH...') → thử tìm trong nội dung đầu
            m = re.search(r'PHỤ\s*LỤC\s*[:\.]?\s*([IVX]+\b|\d{1,2})', out.get(a['key'], '')[:120].upper())
        a['so'] = _ref_num(m.group(1)) if m else prev_so + 1
        prev_so = a['so']

    # --- mục nào dẫn chiếu phụ lục nào ("Xem phụ lục III" → muc_4: [3]) ---
    refs_per_muc = {}
    for k in ['mo_dau', 'can_cu'] + expected + ['ket_luan']:
        txt = out.get(k, '')
        nums = sorted(set(_ref_num(m.group(1)) for m in
                          re.finditer(r'[Pp]hụ\s*lục\s*[:\s]?\s*([IVX]+\b|\d{1,2})', txt)))
        if nums:
            refs_per_muc[k] = nums

    meta = {
        'sections_found': found,
        'sections_missing': missing,
        'so_muc_chuan': max_muc,
        'tables_per_section': tables_count,
        'trang_bia_detected': bool(out.get('trang_bia')),
        'phu_luc_found': appendices,
        'phu_luc_referenced_in_body': refs,
        'phu_luc_refs_per_muc': refs_per_muc,
        'chu_ky': {'giam_sat_truong': has_gst_sign, 'dai_dien_phap_luat': has_daidien_sign},
    }
    return out, meta


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('path')
    ap.add_argument('--type', choices=['dinh_ky', 'hoan_thanh'], help='Ép loại báo cáo')
    ap.add_argument('--nghidinh', choices=['nd06', 'nd207'],
                    help='Ép nghị định: nd06 (06/2021) hoặc nd207 (207/2026)')
    ap.add_argument('--output-json', action='store_true')
    ap.add_argument('--output-file')
    args = ap.parse_args()

    if not os.path.exists(args.path):
        print(json.dumps({'error': f'File not found: {args.path}'})); sys.exit(1)

    items = walk_pdf(args.path) if args.path.lower().endswith('.pdf') else walk_docx(args.path)
    rtype, conf, signals = detect_type(items, args.type)
    if rtype == 'unknown':
        rtype = 'hoan_thanh'
        signals.append('KHÔNG chắc chắn loại báo cáo — mặc định hoan_thanh, hãy kiểm tra lại hoặc dùng --type')

    decree, dconf, dsignals = detect_decree(items, rtype, args.nghidinh)
    sections, meta = parse(items, rtype, decree)
    output = {
        '_meta': {
            'source_file': os.path.basename(args.path),
            'loai_bao_cao': rtype,
            'nghi_dinh': decree,
            'do_tin_cay': conf,
            'do_tin_cay_nghi_dinh': dconf,
            'dau_hieu_nhan_dien': signals,
            'dau_hieu_nghi_dinh': dsignals,
            **meta,
        },
        **sections,
    }
    js = json.dumps(output, ensure_ascii=False, indent=2)
    if args.output_file:
        with open(args.output_file, 'w', encoding='utf-8') as f:
            f.write(js)
        print(f'Saved to {args.output_file}')
    if args.output_json:
        print(js)
    if not args.output_json and not args.output_file or args.output_file:
        m = output['_meta']
        nd_label = 'NĐ 207/2026' if decree == 'nd207' else 'NĐ 06/2021'
        so_muc = m['so_muc_chuan']
        print(f"File: {m['source_file']}")
        print(f"Loại báo cáo: {'ĐỊNH KỲ' if rtype=='dinh_ky' else 'HOÀN THÀNH'} — {nd_label} "
              f"({so_muc} mục, PL IV{'a' if rtype=='dinh_ky' else 'b'}) "
              f"[tin cậy loại: {conf}, tin cậy nghị định: {dconf}]")
        for s in m['dau_hieu_nhan_dien']: print(f'  • {s}')
        print(f"Nghị định áp dụng: {nd_label}")
        for s in m['dau_hieu_nghi_dinh']: print(f'  • {s}')
        print(f"Trang bìa/lót: {'CÓ — đã tách riêng' if m['trang_bia_detected'] else 'không phát hiện'}")
        print(f"Mục tìm thấy: {len(m['sections_found'])}/{m['so_muc_chuan']}")
        if m['sections_missing']: print(f"  THIẾU: {', '.join(m['sections_missing'])}")
        for k in ['trang_bia','mo_dau','can_cu'] + [f'muc_{i}' for i in range(1, m['so_muc_chuan']+1)] + ['ket_luan']:
            if k in sections:
                ntb = m['tables_per_section'].get(k, 0)
                preview = sections[k][:75].replace('\n', ' ')
                print(f"  {k}: {len(sections[k])} ký tự, {ntb} bảng — {preview}...")
        print(f"Phụ lục trong file: {len(m['phu_luc_found'])}")
        for a in m['phu_luc_found']: print(f"  • {a['ten'][:80]} ({a['key']})")
        if m['phu_luc_referenced_in_body']:
            print(f"Phụ lục được dẫn chiếu trong thân: {', '.join(m['phu_luc_referenced_in_body'])}")
        ck = m['chu_ky']
        print(f"Chữ ký: GST={'CÓ' if ck['giam_sat_truong'] else 'THIẾU'}, Đại diện PL={'CÓ' if ck['dai_dien_phap_luat'] else 'THIẾU/không cần (định kỳ)'}")


if __name__ == '__main__':
    main()
