# -*- coding: utf-8 -*-
"""render_html.py — render review_data dict thành HTML (xem trực tiếp + tải về)."""

import html

C = {
    'primary': '#1F4E79', 'accent': '#E8731A', 'grey': '#595959', 'light': '#EAF1F8',
}
ST = {
    'DAT': ('ĐẠT', '#2E7D32', '#E2F0E4'),
    'CAN_SUA': ('CẦN SỬA', '#B7791F', '#FDF3D7'),
    'LOI': ('LỖI', '#C0392B', '#FADBD8'),
    'THIEU': ('THIẾU MỤC', '#C0392B', '#FADBD8'),
}
XL = {
    'ĐẠT': ('#2E7D32', '#E2F0E4'), 'ĐẠT CÓ ĐIỀU KIỆN': ('#B7791F', '#FDF3D7'),
    'CẦN SỬA': ('#B7791F', '#FDF3D7'), 'KHÔNG ĐẠT': ('#C0392B', '#FADBD8'),
}


def e(s):
    return html.escape(str(s or ''))


def badge(st):
    txt, fg, bg = ST.get(st, ST['CAN_SUA'])
    return (f'<span style="background:{bg};color:{fg};font-weight:700;'
            f'padding:2px 10px;border-radius:12px;font-size:12px;white-space:nowrap">{txt}</span>')


def build_html(d):
    kq = d.get('ket_qua_chung', {})
    xl = kq.get('xep_loai', 'CẦN SỬA')
    fg, bg = XL.get(xl.upper(), ('#B7791F', '#FDF3D7'))
    tk = d.get('thong_ke', {})

    info_rows = ''.join(
        f'<tr><td class="lbl">{e(k)}</td><td>{e(v)}</td></tr>'
        for k, v in [
            ('Dự án / Công trình', d.get('du_an')),
            ('Hạng mục / Giai đoạn', d.get('hang_muc_giai_doan')),
            ('Loại báo cáo', d.get('loai_bao_cao_text')),
            ('Kỳ báo cáo', d.get('ky_bao_cao')),
            ('Báo cáo được đánh giá', d.get('file_bao_cao')),
            ('Đơn vị lập báo cáo', d.get('don_vi_lap')),
            ('Chủ đầu tư', d.get('chu_dau_tu')),
            ('Ngày / Người đánh giá',
             f"{d.get('ngay_danh_gia','')} — {d.get('nguoi_danh_gia','Phòng Kỹ thuật')}"),
        ] if v)

    rows = ''
    for i, m in enumerate(d.get('danh_gia_muc', [])):
        zebra = 'background:#F5F8FC' if i % 2 else ''
        rows += (f'<tr style="{zebra}"><td class="c">{e(m.get("muc"))}</td>'
                 f'<td>{e(m.get("ten"))}</td><td class="c">{badge(m.get("trang_thai"))}</td>'
                 f'<td>{e(m.get("nhan_xet_ngan"))}</td></tr>')

    ph = ''
    for p in d.get('phat_hien_chinh', []):
        sev = p.get('muc_do', 'CAN_SUA') if isinstance(p, dict) else 'CAN_SUA'
        txt = p.get('noi_dung', '') if isinstance(p, dict) else str(p)
        _, pfg, pbg = ST.get(sev, ST['CAN_SUA'])
        ph += (f'<li style="margin:6px 0"><span style="color:{pfg};font-weight:700">'
               f'[{ST.get(sev, ST["CAN_SUA"])[0]}]</span> {e(txt)}</li>')

    detail = ''
    for m in d.get('danh_gia_muc', []):
        detail += (f'<h3 style="color:{C["primary"]};margin:18px 0 6px">'
                   f'{e(m.get("muc"))} — {e(m.get("ten"))} {badge(m.get("trang_thai"))}</h3>')
        for x in m.get('dat_duoc', []):
            detail += f'<div class="li ok">— {e(x)}</div>'
        for x in m.get('van_de', []):
            detail += f'<div class="li bad">■ {e(x)}</div>'

    kn = ''.join(f'<li style="margin:5px 0">{e(x)}</li>' for x in d.get('kien_nghi', []))

    return f"""<!DOCTYPE html>
<html lang="vi"><head><meta charset="utf-8">
<title>{e(d.get('tieu_de'))}</title>
<style>
 body{{font-family:'Segoe UI',Calibri,Arial,sans-serif;color:#262626;max-width:900px;
      margin:0 auto;padding:24px;font-size:14.5px;line-height:1.5}}
 .hdr{{display:flex;justify-content:space-between;border-bottom:3px solid {C['primary']};
      padding-bottom:8px;margin-bottom:18px}}
 table{{border-collapse:collapse;width:100%;margin:10px 0}}
 td,th{{border:1px solid #D9E4F0;padding:7px 10px;vertical-align:top}}
 td.lbl{{background:{C['light']};color:{C['primary']};font-weight:700;width:220px}}
 th{{background:{C['primary']};color:#fff;text-align:left}}
 td.c,th.c{{text-align:center}}
 .banner{{display:flex;border:2px solid {fg};border-radius:8px;overflow:hidden;margin:16px 0}}
 .banner .l{{background:{fg};color:#fff;padding:16px 22px;text-align:center;min-width:130px}}
 .banner .r{{background:{bg};padding:14px 18px;flex:1}}
 h2{{color:{C['primary']};border-bottom:2px solid #BFD3E6;padding-bottom:4px;margin-top:28px}}
 h2 .n{{color:{C['accent']}}}
 .li{{margin:4px 0 4px 14px}} .li.bad{{color:#9C3D2E}} .li.ok{{color:#262626}}
 .concl{{border:2px solid {C['primary']};background:{C['light']};border-radius:8px;
        padding:14px 18px;margin:10px 0}}
 .sig{{display:flex;justify-content:space-around;margin-top:40px;text-align:center}}
 .chip{{font-weight:700;font-size:18px;margin-right:4px}}
 footer{{margin-top:36px;border-top:2px solid {C['primary']};padding-top:6px;
        color:{C['grey']};font-size:12px;display:flex;justify-content:space-between}}
 @media print{{ body{{padding:0}} }}
</style></head><body>
<div class="hdr">
  <div><b style="color:{C['primary']}">CÔNG TY CỔ PHẦN TEXO TƯ VẤN VÀ ĐẦU TƯ</b><br>
       <span style="color:{C['grey']};font-weight:700;font-size:12px">PHÒNG KỸ THUẬT</span></div>
  <div style="text-align:right"><b style="color:{C['accent']}">PHIẾU ĐÁNH GIÁ BÁO CÁO TVGS</b><br>
       <span style="color:{C['grey']};font-size:12px">{e(d.get('so_phieu'))}</span></div>
</div>
<div style="color:{C['accent']};font-weight:700;font-size:12.5px">PHÒNG KỸ THUẬT • ĐÁNH GIÁ CHẤT LƯỢNG BÁO CÁO TƯ VẤN GIÁM SÁT</div>
<h1 style="color:{C['primary']};border-bottom:4px solid {C['accent']};padding-bottom:6px;margin-top:4px">{e(d.get('tieu_de'))}</h1>
<table>{info_rows}</table>
<div class="banner">
  <div class="l"><div style="font-size:11px;font-weight:700">KẾT QUẢ</div>
    <div style="font-size:21px;font-weight:800">{e(xl)}</div></div>
  <div class="r">{e(kq.get('tom_tat'))}</div>
</div>
<div>
 <span class="chip" style="color:#2E7D32">{tk.get('dat', 0)}</span>tiêu chí ĐẠT&nbsp;&nbsp;&nbsp;
 <span class="chip" style="color:#B7791F">{tk.get('can_sua', 0)}</span>điểm CẦN SỬA&nbsp;&nbsp;&nbsp;
 <span class="chip" style="color:#C0392B">{tk.get('loi', 0)}</span>LỖI / mâu thuẫn&nbsp;&nbsp;&nbsp;
 <span class="chip" style="color:#C0392B">{tk.get('muc_thieu', 0)}</span>mục THIẾU
</div>
<h2><span class="n">1</span> TỔNG HỢP ĐÁNH GIÁ THEO MỤC</h2>
<table><tr><th class="c" style="width:60px">Mục</th><th>Nội dung</th>
<th class="c" style="width:110px">Kết quả</th><th>Nhận xét chính</th></tr>{rows}</table>
<h2><span class="n">2</span> CÁC PHÁT HIỆN CHÍNH</h2>
<ul style="padding-left:18px">{ph or '<li>Không có phát hiện đáng kể.</li>'}</ul>
<h2><span class="n">3</span> ĐÁNH GIÁ CHI TIẾT TỪNG MỤC</h2>
{detail}
<h2><span class="n">4</span> YÊU CẦU CHỈNH SỬA &amp; KIẾN NGHỊ</h2>
<ol style="padding-left:20px">{kn or '<li>Không có.</li>'}</ol>
<h2><span class="n">5</span> KẾT LUẬN</h2>
<div class="concl">{e(d.get('ket_luan'))}</div>
<div class="sig">
  <div><b style="color:{C['primary']}">NGƯỜI ĐÁNH GIÁ</b><br>
       <i style="color:{C['grey']};font-size:12px">(Ký, ghi rõ họ tên)</i><br><br><br><br>
       <b>{e(d.get('nguoi_danh_gia_ten'))}</b></div>
  <div><b style="color:{C['primary']}">TRƯỞNG PHÒNG KỸ THUẬT</b><br>
       <i style="color:{C['grey']};font-size:12px">(Ký, ghi rõ họ tên)</i><br><br><br><br>
       <b>{e(d.get('truong_phong_ten'))}</b></div>
</div>
<footer><span>Phòng Kỹ thuật — TEXO • {e(d.get('ngay_danh_gia'))}</span>
<span>tvgs-checker-app (kiểm tra tự động không AI — cần chuyên môn xác nhận)</span></footer>
</body></html>"""
