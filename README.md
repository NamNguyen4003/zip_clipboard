# OneNote → New Outlook: nén ảnh trước khi paste

Report OneNote nặng (~30 MB ảnh) paste vào New Outlook làm máy treo. `Nen-Clipboard.bat` thu mỗi ảnh trong clipboard về đúng kích thước đang hiển thị trên trang (PNG, hoặc JPEG chất lượng 85 khi ảnh không trong suốt và JPEG nhỏ hơn PNG ít nhất 3 lần; phóng to ảnh trong email sẽ kém nét hơn bản gốc), bỏ các định dạng clipboard nặng, rồi mới để bạn paste. Dùng được với OneNote web (ảnh nhúng dạng `data:`) và OneNote desktop (ảnh dạng `file:`). Cả hai mới được thử bằng dữ liệu giả trong SelfTest; OneNote desktop chưa thử với một trang thật. Mọi thứ nằm trong một file duy nhất, không cần cài gì.

## Cài một lần
Mở **cmd** (Win+R → `cmd` → Enter), dán lệnh sau rồi Enter. Không cần git:
```bat
curl -fsSL https://raw.githubusercontent.com/NamNguyen4003/zip_clipboard/main/install.cmd -o install.cmd && .\install.cmd && del install.cmd
```
Lệnh tải `Nen-Clipboard.bat` vào thư mục **Downloads** và tạo shortcut `Nen-Clipboard` trên Desktop (phím `Ctrl + Alt + C`, chạy Minimized).
- Đừng xóa file trong Downloads: mất file là phím tắt hỏng. Lỡ xóa thì chạy lại lệnh cài.
- Cập nhật: chạy lại lệnh cài; file và shortcut cũ sẽ bị ghi đè.

## Dùng
1. OneNote: bấm **Ctrl+A** vài lần tới khi chọn cả trang (một lần có thể chỉ chọn một đoạn chữ, không có ảnh), rồi **Ctrl+C**.
2. **Ctrl+Alt+C** (hoặc bấm đúp file), chờ thông báo "N ảnh: X MB -> Y MB" (vài giây). Nếu có ảnh bị bỏ qua, thông báo thêm "Bỏ qua k ảnh." Bấm lần hai trên clipboard đã nén thì thông báo "N ảnh đã đủ nhỏ, cứ Ctrl+V vào Outlook." và clipboard không bị ghi lại.
3. New Outlook: **Ctrl+V**.

Nếu có lỗi hoặc không nén được ảnh nào, file hiện thông báo nêu lý do và **không đụng vào clipboard**:
- "không có ảnh": phần đã copy không chứa ảnh; chọn lại cả trang rồi copy.
- "Không nén được ảnh nào (...)": liệt kê lý do và số ảnh (file ảnh không tồn tại, link web, định dạng không hỗ trợ...).

**Win+V:** lịch sử Win+V không hiện mục lớn hơn 4 MB, đó là bình thường (Ctrl+V và script vẫn đọc được). Đừng dùng Win+V để kiểm tra việc copy, và đừng bấm chọn mục cũ trong đó vì nó thay luôn nội dung clipboard.

## Kiểm tra
```powershell
powershell -STA -NoProfile -ExecutionPolicy Bypass -Command "& ([scriptblock]::Create((Get-Content -LiteralPath 'Nen-Clipboard.bat' -Raw -Encoding UTF8))) -SelfTest"
```
Lệnh này chạy qua clipboard thật hai report giả: kiểu OneNote web (17 ảnh `data:` PNG, khoảng 30 MB, nén hai lần để thấy lần hai không đổi gì) và kiểu OneNote desktop (31 ảnh `file:`, khoảng 33 MB). Nó kiểm tra: chọn PNG/JPEG, kích thước ảnh, tiếng Việt, header CF_HTML, tỉ lệ giảm, thời gian, bộ nhớ. Kết quả lần chạy gần nhất: web 29.6 MB → 0.1 MB trong 6.9 giây; desktop 32.8 MB → 2.4 MB trong 2.2 giây.

## Chẩn đoán (khi paste không ra ảnh)
Bấm **Win+R**, dán lệnh sau (thay đường dẫn cho đúng chỗ bạn để file), rồi Enter:
- `"C:\...\Nen-Clipboard.bat" -TestPaste`: đặt hai ảnh test lên clipboard, A (đỏ) nhúng trong HTML, B (xanh) trỏ tới file. Ctrl+V vào Outlook để xem ảnh nào hiện.
- `"C:\...\Nen-Clipboard.bat" -Diag`: copy trang OneNote trước, rồi chạy lệnh này. Nó ghi `nen-clipboard-diag.txt` ra Desktop. File chỉ chứa cấu trúc, không có chữ trong report; đường dẫn user được thay bằng `<user>`.

Lưu ý khi sửa file: lưu dạng **UTF-8 không BOM** và xuống dòng **CRLF**. Nếu sai một trong hai, cmd sẽ chạy nhầm cả phần PowerShell.
