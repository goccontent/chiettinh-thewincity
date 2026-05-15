# Chiết Tính – THE WIN CITY (Streamlit)

App online tra giá theo Mã sản phẩm từ file Giỏ hàng và xuất file CHIETTINH tự động.

## Chạy thử trên máy
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```
Trình duyệt sẽ tự mở tại http://localhost:8501

## Deploy lên Streamlit Community Cloud (miễn phí)

1. **Tạo repo GitHub** chứa toàn bộ thư mục này (bao gồm 2 file Excel `CHIETTINH.xlsx` và `GIỎ HÀNG ... .xlsx`, các file `streamlit_app.py`, `requirements.txt`, `packages.txt`, `.streamlit/config.toml`).

   ```bash
   cd "D:\CHIET TINH THE WIN CITY"
   git init
   git add .
   git commit -m "Initial commit: Chiet Tinh app"
   git branch -M main
   git remote add origin https://github.com/<USER>/<REPO>.git
   git push -u origin main
   ```

2. Truy cập <https://share.streamlit.io> → đăng nhập GitHub → **New app**.

3. Chọn:
   - Repository: `<USER>/<REPO>`
   - Branch: `main`
   - Main file path: `streamlit_app.py`

4. Bấm **Deploy**. Sau 2–3 phút app sẽ chạy tại URL dạng `https://<tên-app>.streamlit.app`.

### Lưu ý
- `packages.txt` đã khai báo `libreoffice` → server sẽ tự cài để tính lại công thức trong file Excel xuất ra. Lần build đầu tiên có thể mất ~5 phút.
- Nếu đổi tên file Excel, sửa biến `GIO_HANG_FILE` và `CHIETTINH_FILE` ở đầu `streamlit_app.py`.
- Khi cập nhật Giỏ hàng mới: thay file Excel trong repo → `git push` → Streamlit Cloud tự deploy lại.
