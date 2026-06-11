# tvgs-checker-app

Web app kiểm tra **báo cáo Tư vấn Giám sát (TVGS)** theo Phụ lục IV — NĐ 06/2021, của Phòng Kỹ thuật — Công ty CP TEXO Tư vấn và Đầu tư.

Hỗ trợ cả 2 loại: **báo cáo định kỳ** (8 mục — PL IVa) và **báo cáo hoàn thành** (12 mục — PL IVb). Kết quả xuất ra **phiếu đánh giá .docx** (A4, định dạng Phòng Kỹ thuật) và **HTML** xem trực tiếp.

## Hai chế độ phân tích (HYBRID)

| | Không AI (mặc định) | Dùng AI (tùy chọn) |
|---|---|---|
| Chi phí | Miễn phí | Theo API key của bạn |
| Dữ liệu | Xử lý cục bộ trên server app | **Toàn văn báo cáo gửi tới nhà cung cấp AI** |
| Cách hoạt động | Bộ quy tắc TEXO: từ khóa bắt buộc, bảng biểu, đối chiếu chéo mâu thuẫn, phụ lục, chữ ký, NĐ hết hiệu lực, chính tả | AI đánh giá sâu từng mục theo **chú thích kiểm tra gốc** + câu hỏi tư duy, **trên nền** kết quả quy tắc (không thay thế) |
| Cần gì | Không cần gì | API key tự nhập: Claude / Gemini / OpenAI / tương thích OpenAI (OpenRouter...) |

Không có key → app vẫn chạy đầy đủ ở chế độ quy tắc. AI lỗi (401/404/429...) → app báo gợi ý tiếng Việt và tự quay về kết quả quy tắc.

## Cấu trúc

```
tvgs-checker-app/
├── app.py                      # Giao diện Streamlit (cổng mật khẩu, 2 chế độ, hiệu chỉnh, tải về)
├── requirements.txt            # streamlit, python-docx, requests
├── packages.txt                # poppler-utils (đọc PDF trên Streamlit Cloud)
├── .streamlit/config.toml      # theme màu TEXO
├── core/
│   ├── extract_sections.py     # Trích xuất mục (bìa, số La Mã tự động, bảng, phụ lục, chữ ký)
│   ├── analyzer.py             # Lớp quy tắc: Tầng 1 (từ khóa) + Tầng 1.5 (heuristic)
│   ├── llm.py                  # Gọi AI đa nhà cung cấp qua REST: lấy danh sách model thật,
│   │                           #   kiểm tra kết nối, vá JSON cắt cụt, map lỗi tiếng Việt
│   ├── ai_reviewer.py          # Lớp AI: prompt theo chú thích kiểm tra gốc + merge về phiếu
│   ├── generate_review_docx.py # Render phiếu đánh giá .docx
│   └── render_html.py          # Render phiếu đánh giá HTML
└── criteria/
    ├── dinh_ky_8muc.json       # Tiêu chí báo cáo định kỳ (PL IVa) — kèm chu_thich_goc
    └── hoan_thanh_12muc.json   # Tiêu chí báo cáo hoàn thành (PL IVb) — kèm chu_thich_goc
```

## Chạy thử trên máy

```bash
pip install -r requirements.txt
streamlit run app.py
```

Mở http://localhost:8501 — mật khẩu mặc định: `texo2026`.

## Deploy lên Streamlit Community Cloud (miễn phí)

1. Đưa toàn bộ folder này lên 1 repo GitHub (**nên để private** — bộ tiêu chí là kinh nghiệm nội bộ).
2. Vào https://share.streamlit.io → **New app** → chọn repo, branch `main`, file `app.py` → Deploy.
3. (Khuyến nghị) Đổi mật khẩu: trang quản lý app → **Settings → Secrets**:

```toml
APP_PASSWORD = "mat-khau-moi"
```

Không đặt secret thì mật khẩu mặc định là `texo2026`.

## Dùng chế độ AI

1. Sidebar → chọn **Dùng AI** → chọn nhà cung cấp → dán API key (key chỉ giữ trong phiên, app không lưu).
2. Bấm **Lấy danh sách model** → chọn model từ danh sách thật của key (không gõ tay — tên model đổi liên tục).
3. Bấm **Kiểm tra kết nối** để xác nhận key + model.
4. Upload báo cáo → bấm **Phân tích sâu bằng AI** → hiệu chỉnh → tải .docx / .html.

Lấy key: Claude (console.anthropic.com) · Gemini (aistudio.google.com — có bậc miễn phí) · OpenAI (platform.openai.com) · OpenRouter (openrouter.ai, dùng loại "Tương thích OpenAI" với Base URL `https://openrouter.ai/api/v1`).

## Bảo mật

- API key chỉ tồn tại trong phiên làm việc; app không ghi key ra đĩa, không gửi đi đâu ngoài nhà cung cấp AI bạn chọn.
- Bật AI nghĩa là toàn văn báo cáo được gửi tới nhà cung cấp AI — báo cáo nhạy cảm hãy dùng chế độ Không AI.
- Không dán key trên máy công cộng; không commit key lên GitHub; lộ key do cách dùng thuộc trách nhiệm người dùng.

## Lưu ý chất lượng

- Khuyến nghị upload **.docx** (đọc được bảng + đánh số tự động của Word). PDF kém chính xác hơn.
- Kết quả (kể cả khi dùng AI) **không thay thế đánh giá chuyên môn** — người kiểm tra xác nhận lại trước khi phát hành. Trước khi tải phiếu có thể hiệu chỉnh trực tiếp: thông tin dự án, xếp loại, trạng thái/nhận xét từng mục, kết luận.
- Phần gọi API viết theo "API giả định" (chưa chạy bằng key thật trong môi trường phát triển) — lần đầu dùng hãy bấm **Kiểm tra kết nối** trước.
