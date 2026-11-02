# Dashboard trưng bày Daily - KHÔNG CẦN PYTHON

Bộ này dùng **Windows PowerShell + Microsoft Excel desktop**. Không dùng Python.

## Thay vào repo hiện tại
1. Giải nén ZIP.
2. Copy toàn bộ file/folder bên trong vào folder repo `GITHUB_DASHBOARD_DAILY_FIXED` hiện tại.
3. Chọn **Replace** khi Windows hỏi ghi đè.
4. Không xóa folder `.git` của repo hiện tại.

## Data nguồn hằng ngày
Đặt đúng 2 file trong `source`:
- `New Microsoft Excel Worksheet.xlsx`: sheet `Doanh số`, và sheet `Clear` nếu có.
- `Trung bay.xlsx`: mức trưng bày hiện tại.

Giữ nguyên tên file và tên cột. Script đọc trực tiếp 2 workbook này. File nguồn không được commit lên GitHub vì `.gitignore` đã chặn `source/*.xlsx`.

## Run hằng ngày
1. Đóng 2 file Excel nguồn nếu đang mở.
2. Ghi đè file mới vào `source`.
3. Double-click `RUN_UPDATE_DAILY.bat`.
4. BAT sẽ:
   - mở Excel ẩn và chạy RefreshAll cho workbook doanh số;
   - đọc doanh số và mức trưng bày hiện tại;
   - dùng nguyên logic trong `source/template.html`;
   - tạo `docs/index.html` cho ASM;
   - tạo `docs/npp/<ma-npp>/index.html` riêng từng NPP;
   - tạo `LINK_NPP.csv`;
   - `git add`, `commit`, `push`.
5. Khi xong, `LINK_NPP.csv` tự mở. Link không đổi qua các ngày.

## Cấu trúc link
- ASM: `https://vycharming2012-cpu.github.io/GITHUB_DASHBOARD_DAILY_FIXED/`
- NPP: `https://vycharming2012-cpu.github.io/GITHUB_DASHBOARD_DAILY_FIXED/npp/<npp>/`

## Điều kiện máy
- Windows PowerShell có sẵn trên Windows.
- Microsoft Excel desktop đã cài, vì script dùng Excel COM để đọc/refresh workbook.
- Git/GitHub Desktop đã đăng nhập để push.

## Nếu Push failed
Output HTML vẫn đã tạo trong `docs`. Mở GitHub Desktop và bấm `Push origin`.
