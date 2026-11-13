# CoolWord - Bộ Công Cụ Tiếng Việt Mở Rộng Cho Microsoft Word (2016+)

**CoolWord** là Add-in native (C# .NET COM Add-in + WPF Fluent UI) dành cho **Microsoft Word 2016, 2019, 2021, 2024 và Microsoft 365** trên Windows. Add-in hoạt động hoàn toàn offline, tích hợp sâu vào **Ribbon Tab**, **Menu Chuột Phải (Context Menu)** và **Bảng Điều Khiển (Custom Task Pane)** của Word, đồng thời được thiết kế theo kiến trúc **Plugin / Module (`IToolModule`)** dễ dàng mở rộng.

---

## 1. Tính năng nổi bật

### 🔢 Module 1: Đổi Số Ra Chữ Tiếng Việt (`NumberToTextModule`)
- **Chuyển đổi trực tiếp trên văn bản Word**:
  - Bôi đen con số trong văn bản hoặc bảng biểu (hoặc chỉ cần đặt con trỏ vào giữa con số) và nhấn **Đổi Số Đang Chọn**.
  - **Menu Chuột Phải thông minh**: Khi bôi đen số và nhấp chuột phải, CoolWord tự động đọc số và hiển thị **xem trước kết quả bằng chữ ngay trên dòng đầu tiên của menu chuột phải** (`Đổi số: "Một triệu hai trăm năm mươi nghìn đồng chẵn"`).
- **4 Chế độ thay thế / chèn linh hoạt**:
  1. **Chèn trong ngoặc sau số**: `1.250.000` → `1.250.000 (Một triệu hai trăm năm mươi nghìn đồng chẵn)`
  2. **Thay thế trực tiếp số gốc**: `1.250.000` → `Một triệu hai trăm năm mươi nghìn đồng chẵn`
  3. **Chèn kèm nhãn trong ngoặc**: `1.250.000` → `1.250.000 (Bằng chữ: Một triệu hai trăm năm mươi nghìn đồng chẵn)`
  4. **Xuống dòng mới**: Chèn dòng `Bằng chữ: ...` ngay bên dưới.
  - Hỗ trợ **Đổi hàng loạt mọi con số trong vùng bôi đen** và hỗ trợ `Ctrl+Z` (Undo 1 bước).
- **Chuẩn âm vị & quy tắc kế toán Tiếng Việt**:
  - Xử lý chính xác các quy tắc đọc đặc biệt: `mười`, `mốt` (`21`), `tư`/`bốn` (`14` đọc là *mười bốn*, `24` đọc là *hai mươi tư*), `lăm`/`năm` (`15` đọc là *mười lăm*, `105` đọc là *một trăm lẻ năm*), `không trăm` (`1.005`), và các cấp hàng `tỷ`, `nghìn tỷ` (`1.002.000.000.000` → *Một nghìn không trăm lẻ hai tỷ*).
  - **Đơn vị đọc**: Tiền Việt Nam (`đồng` / `chẵn`), Ngoại tệ `USD` (`đô la Mỹ` và `xu`, tự động nhận diện ký hiệu `$` hoặc `USD`), Số thuần túy (đọc số âm, số thập phân `phẩy`), hoặc Đơn vị tùy chỉnh.
  - **Phương ngữ**: Phổ thông hành chính (`nghìn, lẻ, tư`), Chuẩn Bắc Bộ (`nghìn, linh, tư`), Chuẩn Nam Bộ (`ngàn, lẻ, bốn`).

### 🏛️ Module 2: Chuẩn Hóa Văn Bản Hành Chính Nghị Định 30/2020/NĐ-CP (`AdminFormatModule`)
Được xây dựng mới 100% độc lập dựa trên quy chuẩn pháp lý công khai tại **Phụ lục I - Nghị định 30/2020/NĐ-CP của Chính phủ về công tác văn thư**:
- **Chuẩn hóa toàn diện 1-Click (`ApplyFullStandardOneClick`)**:
  - Thiết lập khổ giấy **A4**, căn lề chuẩn văn thư (Lề trên `20-25 mm`, Lề dưới `20-25 mm`, Lề trái `30-35 mm`, Lề phải `15-20 mm`).
  - Chuẩn hóa phông chữ **Times New Roman** cỡ `13-14 pt`, căn đều 2 bên (`Justify`), thụt lề dòng đầu `1,27 cm`, khoảng cách đoạn `6 pt`, giãn dòng `1.15` cho toàn bộ nội dung (tự động nhận diện và giữ nguyên bố cục các bảng biểu).
  - **Đánh số trang chuẩn NĐ 30/2020**: Cỡ chữ `13 pt` Times New Roman, đặt canh giữa theo chiều ngang ở lề trên (`Header`) và tự động ẩn số trang ở trang thứ nhất (`DifferentFirstPageHeaderFooter`).
  - **Dọn dẹp định dạng rác toàn văn bản**: Xóa khoảng trắng kép, xóa dấu Tab thừa đầu dòng, xóa khoảng trắng thừa trước dấu câu và gộp các dòng trống thừa liên tiếp.
- **Tự động tạo Khung Đầu Văn Bản (`InsertHeaderBlock`)**:
  - Tạo bảng ẩn viền 2x2 chuẩn thể thức: **Cơ quan chủ quản / Cơ quan ban hành** (kèm đường kẻ ngang), **Quốc hiệu - Tiêu ngữ** (`CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM` / `Độc lập - Tự do - Hạnh phúc` kèm đường kẻ ngang), **Số & ký hiệu**, **Địa danh & Ngày tháng** (tự động thêm số `0` trước các ngày `< 10` và các tháng `1, 2` đúng quy định NĐ 30/2020), cùng **Tên loại văn bản & Trích yếu nội dung**.
- **Tự động tạo Khung Cuối Văn Bản (`InsertFooterSignatureBlock`)**:
  - Tạo bảng ẩn viền 1x2 gồm **Nơi nhận** (`12 pt` đậm nghiêng, danh sách `11 pt` đứng) và **Quyền hạn, Chức vụ, Họ tên người ký** (`13-14 pt` in đậm, căn giữa).
- **Đọc ngày tháng năm ra chữ Tiếng Việt**: Hỗ trợ cả định dạng ngày hành chính NĐ 30 và đọc toàn bộ ngày tháng năm bằng chữ cho biên bản/hợp đồng.

### ✨ Module 3: Chuẩn Hóa Văn Bản & Đổi Chữ Hoa/Thường (`TextCleanModule`)
- Xóa khoảng trắng kép, xóa dấu cách thừa trước dấu câu (`, . ; : ! ? )`) và chuẩn hóa khoảng cách sau dấu câu.
- Chuyển đổi kiểu chữ Hoa/Thường chuẩn Tiếng Việt (`Viết Hoa Đầu Mỗi Từ`, `Viết hoa đầu câu`, `IN HOA TOÀN BỘ`, `chữ thường`).

### 🖨️ Module 4: In Ấn & Xuất Bản Thông Minh (`PrintModule`)
- **In 2 mặt thủ công thông minh (Manual Duplex)**: Dành cho máy in 1 mặt, tự động in trang lẻ trước, hiển thị hướng dẫn lật giấy trực quan và in tiếp trang chẵn.
- **In nhanh trang hiện tại & vùng chọn**: In ngay lập tức 1-Click trang đang thao tác hoặc vùng văn bản được bôi đen.
- **In 2 trang trên 1 mặt giấy A4 (2 Pages per sheet / N-up)**: Co tỷ lệ 2 trang tài liệu trên một mặt giấy A4 nằm ngang giúp tiết kiệm giấy và đọc duyệt tài liệu nhanh.
- **Xuất nhanh PDF chuẩn in ấn**: Xuất trực tiếp văn bản sang file PDF giữ nguyên toàn bộ định dạng và phông chữ.
- **Tính năng nghiệp vụ kế toán (dịch vụ logic nền)**: Hỗ trợ hàm chuyên biệt in chứng từ kế toán nhiều liên (Phiếu Thu - Chi 2 liên/3 liên).

### 📑 Module 5: Tiện Ích & Xuất Bản PDF (`PdfModule`)
- **Nghiệp vụ kế toán: Xuất PDF/A Chuẩn Lưu Trữ Kế Toán & Thuế (ISO 19005-1)**:
  - 1-Click xuất Báo cáo tài chính, Sổ kế toán, Chứng từ điện tử chuẩn quốc tế ISO 19005-1 (PDF/A).
  - Tự động nhúng 100% phông chữ hệ thống và cấu trúc dữ liệu toàn vẹn, đáp ứng quy định của Luật Kế toán và Tổng cục Thuế về lưu trữ hồ sơ điện tử lâu dài (10 - 20 năm).
- **Xuất nhanh PDF theo phạm vi**:
  - Chỉ xuất đoạn văn bản hoặc bảng biểu đang bôi đen ra file PDF riêng biệt.
  - Chỉ xuất trang tài liệu hiện tại nơi con trỏ đang đứng ra file PDF.
- **Tối ưu hóa dung lượng (Web/Email)**: Nén kích thước file PDF nhỏ gọn phục vụ gửi nhanh qua Email, Zalo.
- **Hệ thống đóng dấu chìm kế toán & bản quyền (Watermark)**:
  - Đóng dấu chìm xoay nghiêng 45° màu xám mờ trong suốt trên toàn bộ trang.
  - Hỗ trợ các mẫu nghiệp vụ kế toán: `ĐÃ THANH TOÁN`, `ĐÃ KIỂM TOÁN`, `BẢN SAO KẾ TOÁN`, `LƯU HÀNH NỘI BỘ`, `DỰ THẢO` hoặc dấu chìm tùy biến.
  - Hỗ trợ gỡ bỏ toàn bộ dấu chìm 1-Click an toàn.

### 📝 Module 6: Nhập Liệu Nhanh (`QuickInputModule`)
- **Menu Nhập Liệu Nhanh trên Ribbon & Menu Chuột Phải**:
  - Chèn nhanh các **Căn cứ pháp lý** thông dụng: Luật Doanh nghiệp 59/2020, Bộ luật Lao động 45/2019, Nghị định 30/2020/NĐ-CP, Kính gửi Ban Giám đốc, Điều khoản hiệu lực thi hành.
  - Chèn nhanh **Mẫu biểu Kế toán & Thanh toán**: Giấy đề nghị tạm ứng, Giấy đề nghị thanh toán, Thông tin số tài khoản ngân hàng & Xuất hóa đơn VAT, Điều khoản thanh toán, Cam đoan số liệu kế toán trung thực.
  - Chèn nhanh **Điều khoản Hợp đồng**: Điều khoản bảo mật (NDA), Điều khoản bất khả kháng, Điều khoản giải quyết tranh chấp.
  - Chèn nhanh **Ký hiệu & Đánh dấu**: Ô vuông kiểm tra `☐`, Ô vuông đã chọn `☑`, Mũi tên `➔`, Dấu sao `⭐ Ghi chú`.
- **Tạo & Quản lý mẫu người dùng tùy biến**:
  - Bôi đen đoạn văn bản bất kỳ trong Word và nhấn **Lưu đoạn đang chọn thành mẫu mới** để tái sử dụng lâu dài.
  - Tìm kiếm mẫu thông minh, phân loại theo nhóm (*Hành chính, Kế toán, Hợp đồng, Cá nhân, Ký hiệu*) trên bảng điều khiển Task Pane.
  - Hỗ trợ các biến tự động động hóa ngày giờ: `{DATE}` (ngày hiện tại), `{DOC_DATE}` (ngày tháng theo chuẩn văn thư), `{TIME}`, `{YEAR}`.

### 👥 Module 7: Tách Họ - Tên (`NameSplitModule`)
- **Nút bấm Ribbon `Tách Họ-Tên` & Menu Chuột Phải**:
  - **Tách 2 cột trong bảng (Họ đệm & Tên)**: Tự động chèn 1 cột "Tên" bên phải, đổi tiêu đề thành "Họ và tên đệm" và "Tên" để phục vụ sắp xếp danh sách nhân sự, học sinh theo thứ tự A-Z.
  - **Tách 3 cột trong bảng (Họ | Đệm | Tên)**: Phân tách chi tiết 3 cột riêng biệt, tự động nhận diện họ đơn và các họ ghép thông dụng tiếng Việt (*Tôn Thất, Tôn Nữ, Nguyễn Phước, Đặng Vũ...*).
  - **Gộp 2 cột thành 1 cột "Họ và tên"**: Gộp nhanh cột Họ đệm và cột Tên liền kề thành 1 cột Họ và tên hoàn chỉnh và tự động xóa cột phụ.
  - **Xử lý danh sách ngoài bảng**: Tách Họ - Tên theo định dạng dấu Tab `\t` (tiện sao chép/dán sang Excel) hoặc dấu gạch đứng `|`.
  - **Chuẩn hóa danh từ riêng**: Tự động sửa lỗi viết hoa (*nguyễn văn an* $\rightarrow$ *Nguyễn Văn An*), giữ nguyên tiền tố các họ tên dân tộc thiểu số (*H'Hen Niê, Y Moan...*).

### 🔄 Module 8: Word ➔ Word (`WordToWordModule`)
- **Nút Ribbon `Word -> Word` với biểu tượng Word chuẩn**:
  - Tích hợp menu thả xuống kích thước lớn (Large menu) với nhãn 2 dòng `Word ->` / `Word ˅` và biểu tượng Word hiện đại (`FileSaveAsWordDocx`).
- **Chuyển đổi bảng mã Tiếng Việt cũ sang Unicode chuẩn (Times New Roman)**:
  - Tự động nhận diện và chuyển đổi bảng mã **TCVN3 (ABC)** từ các phông `.VnTime`, `.VnArial`, `.VnTimeH`... sang Unicode dựng sẵn.
  - Tự động nhận diện và chuyển đổi bảng mã **VNI-Windows** từ các phông `VNI-Times`, `VNI-Helve`... sang Unicode dựng sẵn.
  - Tự động định dạng lại phông chữ sang **Times New Roman** sau khi chuyển mã.
  - Hoạt động 100% offline, tốc độ cao, hỗ trợ thao tác hoàn tác Undo (`Ctrl+Z`).
- **Nâng cấp & Chuyển đổi định dạng tệp Word (.doc ↔ .docx)**:
  - 1-Click nâng cấp tệp tài liệu cũ `*.doc` (Word 97-2003) sang chuẩn OpenXML `*.docx` hiện đại, giảm dung lượng và tối ưu hóa hiển thị.
  - Lưu bản sao tương thích `*.doc` khi cần nộp hồ sơ, gửi cho hệ thống phần mềm cũ chỉ hỗ trợ chuẩn Word 97-2003.
- **Gộp & Tách tệp Word tự động (Merge & Split)**:
  - **Gộp nhiều tệp Word**: Chọn nhiều tệp `*.docx`, `*.doc` và tự động ghép nối vào tài liệu đang mở, chèn ngắt trang tự động và giữ nguyên cấu trúc.
  - **Tách tệp Word theo Section**: Chia nhỏ tài liệu nhiều chương, phụ lục, hợp đồng nhiều phần thành từng tệp Word riêng biệt (`Phan_01.docx`, `Phan_02.docx`...).
  - **Tách tệp Word theo Trang**: Chia tách tài liệu thành các tệp Word riêng cho mỗi trang (`Trang_001.docx`, `Trang_002.docx`...).

### 📊 Module 9: Word ➔ Excel (`WordToExcelModule`)
- **Nút Ribbon `Word -> Excel` với biểu tượng Excel chuẩn**:
  - Tích hợp menu thả xuống kích thước lớn với nhãn 2 dòng `Word ->` / `Excel ˅` và biểu tượng Excel hiện đại (`FileSaveAsExcelXlsx`).
- **Xuất Bảng biểu Word sang Microsoft Excel (.xlsx)**:
  - **Xuất tất cả bảng biểu**: Trích xuất toàn bộ bảng trong Word sang các bảng tính Excel với đường viền, căn lề và định dạng ô hoàn hảo.
  - **Xuất bảng đang chọn**: Nhanh chóng xuất bảng biểu nơi con trỏ đang đứng hoặc bôi đen.
  - **Tùy chọn phân chia Sheet**: Mỗi bảng trên một Sheet riêng biệt hoặc tất cả bảng xếp nối tiếp trên cùng một Sheet.
- **Tính năng Kế Toán: Xuất Bảng Kê & Báo Cáo Tài Chính Kèm Công Thức SUM**:
  - Tự động nhận diện các cột tiền tệ (`Số tiền`, `Thành tiền`, `Đơn giá`, `Phát sinh Nợ/Có`...) và gán định dạng số kế toán (`#,##0`).
  - Tự động chèn dòng **TỔNG CỘNG** ở chân bảng với công thức `=SUM(...)` chuẩn Excel, kẻ viền kép chuẩn kế toán.
- **Xuất Dòng Văn Bản sang Cột Excel**:
  - Chuyển đổi danh sách đoạn văn, danh mục hàng hóa, câu hỏi, điều khoản thành các dòng trong Excel kèm cột STT tự động.

### 📽️ Module 10: Word ➔ PowerPoint (`WordToPowerPointModule`)
- **Nút Ribbon `Word -> PowerPoint` với biểu tượng PowerPoint chuẩn**:
  - Tích hợp menu thả xuống kích thước lớn với nhãn 2 dòng `Word ->` / `PowerPoint ˅` và biểu tượng PowerPoint hiện đại (`FileSaveAsPowerPointPptx`).
- **Tạo Slide thuyết trình từ Đề mục & Dàn ý (Heading 1-2-3)**:
  - Phân tích cấu trúc phân cấp tài liệu: Heading 1 làm Tiêu đề slide, Heading 2-3 và các ý chính làm các gạch đầu dòng (Bullets) tinh gọn.
- **Tạo Slide linh hoạt theo Trang & Section**:
  - **Theo từng Trang**: Trích xuất mỗi trang tài liệu Word thành một slide trình chiếu tương ứng.
  - **Theo từng Section**: Phù hợp cho báo cáo nhiều chương và đề án phân đoạn.
- **Xuất Bảng biểu & Bộ Slide Báo Cáo Tài Chính Tổng Kết**:
  - **Xuất bảng biểu sang Slide**: Copy và căn giữa từng bảng biểu lên một slide riêng biệt.
  - **Tạo bộ Slide Báo Cáo Tài Chính**: Tự động tạo bài thuyết trình hoàn chỉnh gồm slide bìa, căn cứ số liệu, bảng biểu tài chính và slide kết luận.

---

## 2. Cấu trúc dự án & Kiến trúc mở rộng

```text
c:\projects\CoolWord\
├── CoolWord.sln                    # Giải pháp Visual Studio
├── build.ps1                       # Script biên dịch & chạy bộ kiểm thử tự động
├── install.ps1                     # Đăng ký Add-in vào MS Word (Current User - không cần quyền Admin)
├── uninstall.ps1                   # Gỡ đăng ký Add-in khỏi MS Word
├── bin\                            # Thư mục chứa file thực thi sau khi build
│   ├── CoolWord.dll                # Thư viện COM Add-in chính nạp vào MS Word
│   ├── CoolWord.Studio.exe         # Ứng dụng Desktop quản lý cài đặt & dùng độc lập
│   └── CoolWord.Tests.exe          # Bộ kiểm thử tự động (109 test cases)
├── src\
│   ├── CoolWord\
│   │   ├── AddIn\                  # Tích hợp COM (IDTExtensibility2, Ribbon XML, Custom Task Pane, Installer)
│   │   ├── Core\                   # Kiến trúc lõi (IToolModule, ModuleRegistry, WordDocumentService, SettingsService)
│   │   ├── Modules\
│   │   │   ├── NumberToText\       # Module Đổi Số Ra Chữ Tiếng Việt
│   │   │   ├── AdminFormat\        # Module Chuẩn Hóa Văn Bản Nghị Định 30/2020
│   │   │   ├── TextUtilities\      # Module Chuẩn Hóa Văn Bản & Chữ Hoa/Thường
│   │   │   ├── Printing\           # Module In Ấn & Xuất Bản Thông Minh
│   │   │   ├── Pdf\                # Module Tiện Ích & Xuất Bản PDF
│   │   │   ├── QuickInput\         # Module Nhập Liệu Nhanh (Mẫu văn bản, Kế toán & Snippets)
│   │   │   ├── NameSplit\          # Module Tách Họ - Tên (Bảng biểu & Danh sách nhân sự)
│   │   │   ├── WordToWord\         # Module Word ➔ Word (Chuyển mã TCVN3/VNI, .doc/.docx, Gộp/Tách file)
│   │   │   ├── WordToExcel\        # Module Word ➔ Excel (Xuất bảng biểu, chứng từ kế toán kèm SUM)
│   │   │   └── WordToPowerPoint\   # Module Word ➔ PowerPoint (Tạo slide từ Heading 1-2-3, Báo cáo tài chính)
│   │   └── UI\                     # Giao diện WPF Fluent Design (ThemeColors, MainTaskPaneControl)
│   └── CoolWord.Studio\            # Ứng dụng quản lý & chạy thử giao diện độc lập
└── tests\
    └── CoolWord.Tests\             # Unit tests kiểm tra thuật toán đọc số Tiếng Việt & các Module
```

### Cách mở rộng thêm tính năng (Module) mới sau này
1. Tạo một lớp mới thực thi giao diện `IToolModule` (trong `src/CoolWord/Core/IToolModule.cs`):
   - Khai báo `Id`, `Title`, `Subtitle`, `IconText`, `Order`.
   - Cài đặt `ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)` cho thao tác nhanh.
   - Cài đặt `CreateModuleView(...)` trả về `UIElement` (WPF UserControl) cho giao diện chi tiết.
2. Đăng ký module vào `ModuleRegistry.EnsureInitialized()` (trong `src/CoolWord/Core/ModuleRegistry.cs`).
3. Thanh điều hướng trên `MainTaskPaneControl` sẽ tự động hiển thị tab công cụ mới mà không cần sửa đổi khung giao diện chính.

---

## 3. Hướng dẫn Biên dịch & Cài đặt

- **Biên dịch và chạy kiểm thử**:
  ```powershell
  powershell -ExecutionPolicy Bypass -File .\build.ps1
  ```
- **Cài đặt Add-in vào MS Word (không cần quyền Admin)**:
  ```powershell
  powershell -ExecutionPolicy Bypass -File .\install.ps1
  ```
  Hoặc mở trực tiếp `bin\CoolWord.Studio.exe` để cài đặt/gỡ cài đặt bằng giao diện đồ họa và trải nghiệm ngay các tính năng.
