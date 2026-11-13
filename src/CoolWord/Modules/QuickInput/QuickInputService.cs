using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using CoolWord.Core;
using Word = Microsoft.Office.Interop.Word;

namespace CoolWord.Modules.QuickInput
{
    public static class QuickInputService
    {
        private static readonly List<QuickSnippet> BuiltInSnippets = new List<QuickSnippet>();

        static QuickInputService()
        {
            InitializeBuiltInSnippets();
        }

        private static void InitializeBuiltInSnippets()
        {
            // 1. Căn Cứ Pháp Lý & Hành Chính
            BuiltInSnippets.Add(new QuickSnippet(
                "snip-luat-dn",
                "Căn cứ Luật Doanh nghiệp số 59/2020/QH14",
                "Hành chính",
                "Căn cứ Luật Doanh nghiệp số 59/2020/QH14 ngày 17 tháng 6 năm 2020 của Quốc hội;",
                "ccdn",
                true));

            BuiltInSnippets.Add(new QuickSnippet(
                "snip-luat-ld",
                "Căn cứ Bộ luật Lao động số 45/2019/QH14",
                "Hành chính",
                "Căn cứ Bộ luật Lao động số 45/2019/QH14 ngày 20 tháng 11 năm 2019 của Quốc hội;",
                "ccld",
                true));

            BuiltInSnippets.Add(new QuickSnippet(
                "snip-nd30",
                "Căn cứ Nghị định 30/2020/NĐ-CP về công tác văn thư",
                "Hành chính",
                "Căn cứ Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư;",
                "ccnd30",
                true));

            BuiltInSnippets.Add(new QuickSnippet(
                "snip-kinh-gui",
                "Kính gửi Ban Giám đốc và Cán bộ nhân viên",
                "Hành chính",
                "Kính gửi: Ban Giám đốc cùng toàn thể Cán bộ, Nhân viên Công ty.",
                "kgbgd",
                true));

            BuiltInSnippets.Add(new QuickSnippet(
                "snip-hieu-luc-qd",
                "Điều khoản hiệu lực thi hành quyết định",
                "Hành chính",
                "Quyết định này có hiệu lực thi hành kể từ ngày ký. Các phòng ban, đơn vị trực thuộc và các cá nhân có tên tại Điều 1 chịu trách nhiệm thi hành Quyết định này./.",
                "hlqd",
                true));

            // 2. Nghiệp Vụ Kế Toán & Biểu Mẫu
            BuiltInSnippets.Add(new QuickSnippet(
                "snip-dn-tam-ung",
                "Giấy đề nghị tạm ứng (Mẫu chuẩn)",
                "Kế toán",
                "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r\nĐộc lập - Tự do - Hạnh phúc\r\n\r\nGIẤY ĐỀ NGHỊ TẠM ỨNG\r\n\r\nKính gửi: Ban Giám đốc và Phòng Kế toán - Tài chính\r\n\r\nTôi tên là: [Họ và tên]\r\nBộ phận / Phòng ban: [Tên phòng ban]\r\nĐề nghị tạm ứng số tiền: [Số tiền] VNĐ (Bằng chữ: [Bằng chữ])\r\nLý do tạm ứng: [Nêu rõ mục đích công tác / mua sắm]\r\nThời hạn thanh toán tạm ứng: Ngày {DATE}\r\n\r\n{DOC_DATE}\r\nNGƯỜI ĐỀ NGHỊ               KẾ TOÁN TRƯỞNG               GIÁM ĐỐC",
                "gntu",
                true));

            BuiltInSnippets.Add(new QuickSnippet(
                "snip-dn-thanh-toan",
                "Giấy đề nghị thanh toán (Mẫu chuẩn)",
                "Kế toán",
                "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r\nĐộc lập - Tự do - Hạnh phúc\r\n\r\nGIẤY ĐỀ NGHỊ THANH TOÁN\r\n\r\nKính gửi: Ban Giám đốc và Phòng Kế toán - Tài chính\r\n\r\nTôi tên là: [Họ và tên]\r\nBộ phận / Phòng ban: [Tên phòng ban]\r\nĐề nghị thanh toán số tiền: [Số tiền] VNĐ (Bằng chữ: [Bằng chữ])\r\nNội dung thanh toán: [Chi tiết các khoản chi]\r\nChứng từ gốc đính kèm gồm: [Số lượng] hóa đơn, biên nhận.\r\nHình thức thanh toán: Chuyển khoản qua số tài khoản: [Số TK] - Ngân hàng [Tên NH]\r\n\r\n{DOC_DATE}\r\nNGƯỜI ĐỀ NGHỊ               KẾ TOÁN TRƯỞNG               GIÁM ĐỐC",
                "dntt",
                true));

            BuiltInSnippets.Add(new QuickSnippet(
                "snip-tt-nh",
                "Thông tin tài khoản ngân hàng & Xuất hóa đơn",
                "Kế toán",
                "THÔNG TIN THANH TOÁN & XUẤT HÓA ĐƠN:\r\n- Tên đơn vị: [TÊN CÔNG TY / DOANH NGHIỆP]\r\n- Mã số thuế: [MÃ SỐ THUẾ]\r\n- Địa chỉ: [ĐỊA CHỈ ĐĂNG KÝ KINH DOANH]\r\n- Số tài khoản: [SỐ TÀI KHOẢN]\r\n- Ngân hàng: [TÊN NGÂN HÀNG] - Chi nhánh: [TÊN CHI NHÁNH]\r\n- Người thụ hưởng: [TÊN CHỦ TÀI KHOẢN]",
                "ttnh",
                true));

            BuiltInSnippets.Add(new QuickSnippet(
                "snip-dk-tt",
                "Điều khoản thanh toán & Hóa đơn tài chính VAT",
                "Kế toán",
                "Giá trị hợp đồng đã bao gồm thuế Giá trị gia tăng (VAT). Bên A sẽ tiến hành thanh toán cho Bên B bằng hình thức chuyển khoản trong vòng 05 (năm) ngày làm việc kể từ ngày nhận đủ hóa đơn điện tử hợp lệ kèm Biên bản nghiệm thu và bàn giao./.",
                "dktt",
                true));

            BuiltInSnippets.Add(new QuickSnippet(
                "snip-cam-ket-kt",
                "Cam đoan số liệu báo cáo kế toán chính xác, hợp pháp",
                "Kế toán",
                "Chúng tôi xin cam đoan toàn bộ các số liệu, chứng từ và tài liệu giải trình nêu trên là hoàn toàn chính xác, trung thực, hợp lệ và chịu trách nhiệm hoàn toàn trước pháp luật về tính pháp lý của hồ sơ./.",
                "ckkt",
                true));

            // 3. Hợp Đồng & Thỏa Thuận
            BuiltInSnippets.Add(new QuickSnippet(
                "snip-bao-mat",
                "Điều khoản bảo mật thông tin (NDA)",
                "Hợp đồng",
                "ĐIỀU KHOẢN BẢO MẬT: Hai Bên cam kết giữ bí mật tuyệt đối mọi thông tin kỹ thuật, dữ liệu kinh doanh, kế toán tài chính và các nội dung trao đổi trong quá trình thực hiện Hợp đồng. Không bên nào được phép cung cấp hoặc tiết lộ thông tin cho bên thứ ba khi chưa có sự đồng ý bằng văn bản của bên còn lại, ngoại trừ yêu cầu của cơ quan quản lý Nhà nước có thẩm quyền theo quy định của pháp luật.",
                "bmtt",
                true));

            BuiltInSnippets.Add(new QuickSnippet(
                "snip-bat-kha-khang",
                "Điều khoản sự kiện bất khả kháng (Force Majeure)",
                "Hợp đồng",
                "SỰ KIỆN BẤT KHẢ KHÁNG: Sự kiện bất khả kháng là sự kiện xảy ra một cách khách quan không thể lường trước được và không thể khắc phục được mặc dù đã áp dụng mọi biện pháp cần thiết và khả năng cho phép (bao gồm thiên tai, hỏa hoạn, lũ lụt, chiến tranh, dịch bệnh hoặc sự thay đổi của chính sách pháp luật). Bên gặp sự kiện bất khả kháng phải thông báo kịp thời bằng văn bản cho bên kia trong vòng 24 giờ kể từ khi xảy ra sự kiện.",
                "bkkh",
                true));

            BuiltInSnippets.Add(new QuickSnippet(
                "snip-tranh-chap",
                "Điều khoản giải quyết tranh chấp & Trọng tài",
                "Hợp đồng",
                "GIẢI QUYẾT TRANH CHẤP: Mọi tranh chấp phát sinh từ hoặc liên quan đến Hợp đồng này trước hết sẽ được giải quyết thông qua thương lượng hòa giải giữa hai Bên. Trường hợp không tự hòa giải được trong vòng 30 ngày, tranh chấp sẽ được đưa ra giải quyết tại Tòa án nhân dân có thẩm quyền theo quy định của pháp luật Việt Nam. Phán quyết của Tòa án là chung thẩm và có hiệu lực bắt buộc đối với cả hai Bên.",
                "tcpl",
                true));

            // 4. Ký Hiệu & Đánh Dấu Văn Bản
            BuiltInSnippets.Add(new QuickSnippet(
                "snip-box-unchecked",
                "Ô vuông chưa chọn (Checkbox trống)",
                "Ký hiệu",
                "☐ ",
                "box",
                true));

            BuiltInSnippets.Add(new QuickSnippet(
                "snip-box-checked",
                "Ô vuông đã đánh dấu chọn (Checkbox tích)",
                "Ký hiệu",
                "☑ ",
                "check",
                true));

            BuiltInSnippets.Add(new QuickSnippet(
                "snip-arrow",
                "Mũi tên chỉ dẫn (➔)",
                "Ký hiệu",
                "➔ ",
                "arrow",
                true));

            BuiltInSnippets.Add(new QuickSnippet(
                "snip-star",
                "Dấu sao ghi chú (⭐ Ghi chú: )",
                "Ký hiệu",
                "⭐ Ghi chú: ",
                "star",
                true));
        }

        public static IReadOnlyList<QuickSnippet> GetBuiltInSnippets()
        {
            return BuiltInSnippets.AsReadOnly();
        }

        public static List<QuickSnippet> GetAllSnippets()
        {
            List<QuickSnippet> list = new List<QuickSnippet>(BuiltInSnippets);
            if (SettingsService.Current != null && SettingsService.Current.CustomSnippets != null)
            {
                list.AddRange(SettingsService.Current.CustomSnippets);
            }
            return list;
        }

        public static bool AddCustomSnippet(string title, string category, string content, string shortcut, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (string.IsNullOrWhiteSpace(title) || string.IsNullOrWhiteSpace(content))
            {
                statusMessage = "Tiêu đề và nội dung mẫu không được để trống.";
                return false;
            }

            QuickSnippet item = new QuickSnippet(
                Guid.NewGuid().ToString("N"),
                title.Trim(),
                string.IsNullOrWhiteSpace(category) ? "Cá nhân" : category.Trim(),
                content.Trim(),
                shortcut != null ? shortcut.Trim() : string.Empty,
                false);

            if (SettingsService.Current.CustomSnippets == null)
            {
                SettingsService.Current.CustomSnippets = new List<QuickSnippet>();
            }

            SettingsService.Current.CustomSnippets.Insert(0, item);
            SettingsService.Save();

            statusMessage = string.Format("Đã lưu mẫu mới \"{0}\" vào nhóm \"{1}\".", item.Title, item.Category);
            return true;
        }

        public static bool DeleteCustomSnippet(string id, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (string.IsNullOrWhiteSpace(id) || SettingsService.Current.CustomSnippets == null)
            {
                statusMessage = "Không tìm thấy mẫu cần xóa.";
                return false;
            }

            int count = SettingsService.Current.CustomSnippets.RemoveAll(s => string.Equals(s.Id, id, StringComparison.OrdinalIgnoreCase));
            if (count > 0)
            {
                SettingsService.Save();
                statusMessage = "Đã xóa mẫu thành công.";
                return true;
            }

            statusMessage = "Không thể xóa mẫu mặc định của hệ thống.";
            return false;
        }

        public static string ResolvePlaceholders(string template)
        {
            if (string.IsNullOrEmpty(template)) return string.Empty;

            DateTime now = DateTime.Now;
            string dateStr = now.ToString("dd/MM/yyyy");
            string timeStr = now.ToString("HH:mm");
            string yearStr = now.Year.ToString();

            string dayPad = now.Day < 10 ? "0" + now.Day : now.Day.ToString();
            string monthPad = (now.Month == 1 || now.Month == 2) ? "0" + now.Month : now.Month.ToString();
            string docDate = string.Format("ngày {0} tháng {1} năm {2}", dayPad, monthPad, now.Year);

            string resolved = template
                .Replace("{DATE}", dateStr)
                .Replace("{TIME}", timeStr)
                .Replace("{YEAR}", yearStr)
                .Replace("{DOC_DATE}", docDate);

            return resolved;
        }

        /// <summary>
        /// Chèn nội dung mẫu vào vị trí con trỏ (hoặc thay thế vùng bôi đen) trong văn bản Word.
        /// </summary>
        public static bool InsertSnippetIntoWord(Word.Application app, string rawContent, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            if (string.IsNullOrEmpty(rawContent))
            {
                statusMessage = "Nội dung mẫu cần chèn trống.";
                return false;
            }

            try
            {
                string textToInsert = ResolvePlaceholders(rawContent);

                Word.Selection sel = app.Selection;
                if (sel == null)
                {
                    statusMessage = "Không tìm thấy vị trí con trỏ trong tài liệu.";
                    return false;
                }

                StartUndoRecord(app, "CoolWord: Nhập liệu nhanh");
                try
                {
                    sel.Range.Text = textToInsert;
                    // Di chuyển con trỏ ra sau đoạn vừa chèn
                    object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                    sel.Collapse(ref collapseEnd);
                }
                finally
                {
                    EndUndoRecord(app);
                }

                string preview = textToInsert.Length > 30 ? textToInsert.Substring(0, 30) + "..." : textToInsert;
                statusMessage = "Đã chèn nội dung: " + preview.Replace("\r", " ").Replace("\n", " ");
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi chèn mẫu vào Word: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Lấy đoạn văn bản đang chọn trong Word và lưu thành mẫu người dùng.
        /// </summary>
        public static bool SaveSelectionAsCustomSnippet(Word.Application app, string title, string category, string shortcut, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                Word.Selection sel = app.Selection;
                if (sel == null || sel.Range == null || sel.Range.Start == sel.Range.End)
                {
                    statusMessage = "Vui lòng bôi đen đoạn văn bản muốn lưu thành mẫu trước khi bấm lưu.";
                    return false;
                }

                string selectedText = sel.Range.Text;
                if (string.IsNullOrWhiteSpace(selectedText))
                {
                    statusMessage = "Vùng văn bản đang chọn trống.";
                    return false;
                }

                string clean = selectedText.TrimEnd('\r', '\n', '\a');
                string autoTitle = string.IsNullOrWhiteSpace(title)
                    ? (clean.Length > 25 ? clean.Substring(0, 25) + "..." : clean)
                    : title.Trim();

                return AddCustomSnippet(autoTitle, category, clean, shortcut, out statusMessage);
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi lưu đoạn chọn: " + ex.Message;
                return false;
            }
        }

        private static void StartUndoRecord(Word.Application app, string name)
        {
            try
            {
                if (app.UndoRecord != null && !app.UndoRecord.IsRecordingCustomRecord)
                {
                    app.UndoRecord.StartCustomRecord(name);
                }
            }
            catch { }
        }

        private static void EndUndoRecord(Word.Application app)
        {
            try
            {
                if (app.UndoRecord != null && app.UndoRecord.IsRecordingCustomRecord)
                {
                    app.UndoRecord.EndCustomRecord();
                }
            }
            catch { }
        }
    }
}
