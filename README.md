# tvgs-checker-app

Web app kiểm tra **báo cáo Tư vấn Giám sát (TVGS)** theo Phụ lục IV — NĐ 06/2021, của Phòng Kỹ thuật — Công ty CP TEXO Tư vấn và Đầu tư.

Hỗ trợ cả 2 loại: **báo cáo định kỳ** (8 mục — PL IVa) và **báo cáo hoàn thành** (12 mục — PL IVb). App **không dùng AI** — kiểm tra bằng bộ quy tắc (từ khóa bắt buộc, bảng biểu, đối chiếu chéo mâu thuẫn nội bộ, phụ lục, chữ ký, văn bản pháp lý, chính tả) xây dựng từ chú thích kiểm tra nội bộ TEXO. Kết quả xuất ra **phiếu đánh giá .docx** (A4, định dạng Phòng Kỹ thuật) và **HTML** xem trực tiếp.

## Cấu trúc

```
tvgs-checker-app/
├── app.py                      # Giao diện Streamlit
├── requirements.txt            # streamlit, python-docx
├── packages.txt                # poppler-utils (đọc PDF trên Streamlit Cloud)
├── .streamlit/config.toml      # theme màu TEXO
├── core/
│   ├── extract_sections.py     # Trích xuất mục (bìa, số La Mã tự động, bảng, phụ lục)
│   ├── analyzer.py             # Tầng 1 (từ khóa) + Tầng 1.5 (heuristic thay AI)
│   ├── generate_review_docx.py # Render phiếu đánh giá .docx
│   └── render_html.py          # Render phiếu đánh giá HTML
└── criteria/
    ├── dinh_ky_8muc.json       # Tiêu chí báo cáo định kỳ (PL IVa)
    └── hoan_thanh_12muc.json   # Tiêu chí báo cáo hoàn thành (PL IVb)
```

## Chạy thử trên máy

```bash
pip install -r requirements.txt
streamlit run app.py
```

Mở http://localhost:8501 — mật khẩu mặc định: `texo2026`.

## Deploy lên Streamlit Community Cloud (miễn phí)

1. Đưa toàn bộ folder này lên 1 repo GitHub (repo **private** được hỗ trợ).
2. Vào https://share.streamlit.io → **New app** → chọn repo, branch `main`, file `app.py` → Deploy.
3. (Khuyến nghị) Đổi mật khẩu: trong trang quản lý app → **Settings → Secrets**, thêm:

```toml
APP_PASSWORD = "mat-khau-moi"
```

Không đặt secret thì mật khẩu mặc định là `texo2026`.

## Lưu ý chất lượng

- Khuyến nghị upload **.docx** (đọc được bảng + đánh số tự động của Word). PDF đọc bằng poppler, kém chính xác hơn.
- Kết quả là **kiểm tra tự động bằng quy tắc** — phát hiện tốt: thiếu mục, thiếu nội dung bắt buộc, mâu thuẫn tồn tại/tiến độ, Mục 5 viết nhầm bản chất, phụ lục dẫn chiếu thiếu, thiếu chữ ký, dẫn NĐ hết hiệu lực, chính tả phổ biến. KHÔNG đánh giá được chiều sâu chuyên môn — bản đầy đủ dùng skill `tvgs-checker` trên Claude.
- Trước khi tải phiếu, có thể hiệu chỉnh trực tiếp trên app: thông tin dự án, xếp loại, trạng thái/nhận xét từng mục, kết luận.
