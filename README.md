# AI Quản Lý Phòng Trọ

Web app Streamlit dành cho người lớn tuổi:
1. Chọn một hoặc nhiều ảnh chỉ số điện/nước.
2. Chọn bất kỳ file `.xlsx` theo mẫu phòng trọ.
3. Gemini đọc ảnh và trả dữ liệu có cấu trúc.
4. Người dùng kiểm tra/sửa trực tiếp.
5. App chỉ ghi chỉ số mới vào cột I của đúng tháng/phòng, giữ công thức và định dạng.
6. Tải file Excel kết quả.

## Chạy thử
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Gemini API key
Tạo `GEMINI_API_KEY` và đặt trong biến môi trường hoặc Streamlit Secrets.

## Deploy miễn phí
Đẩy `app.py` + `requirements.txt` lên GitHub rồi deploy bằng Streamlit Community Cloud.
