# HƯỚNG DẪN ĐƯA WEB APP LÊN INTERNET

## 1. Tải mã nguồn
Giải nén thư mục dự án.

## 2. Tạo GitHub repository
- Đăng nhập GitHub.
- New repository.
- Đặt tên: `ai-quan-ly-phong-tro`
- Có thể để Public để dùng Streamlit Community Cloud miễn phí.
- Upload các file:
  - `app.py`
  - `requirements.txt`
  - `README.md`
  - `.streamlit/config.toml`

## 3. Tạo Gemini API key
Tạo API key trong Google AI Studio. Không đưa API key vào `app.py` hoặc GitHub.

## 4. Deploy
- Mở Streamlit Community Cloud.
- Đăng nhập bằng GitHub.
- Chọn Create app.
- Chọn repository vừa tạo.
- Entrypoint: `app.py`
- Deploy.

## 5. Đặt API key trong Secrets
Trong app đã deploy:
- Settings / Secrets
- Thêm:

```toml
GEMINI_API_KEY = "DÁN_API_KEY_CỦA_ANH_VÀO_ĐÂY"
```

- Save.
- Reboot/re-run app.

## 6. Kết quả
Streamlit cấp cho app một URL dạng:
`https://ten-app-cua-anh.streamlit.app`

Anh có thể gửi URL đó cho người khác.

## Lưu ý
- Đây là web app xử lý theo từng lần upload; file Excel không được lưu lâu dài trên server.
- File kết quả được tạo trong phiên làm việc và người dùng tải xuống.
- Không commit API key vào GitHub.
- Gói miễn phí phù hợp cho dự án cá nhân/giáo dục và lưu lượng nhỏ. Nếu số người dùng tăng mạnh hoặc cần uptime/đảm bảo sản xuất, nên chuyển sang hạ tầng trả phí.
