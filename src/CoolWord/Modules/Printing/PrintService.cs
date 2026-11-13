using System;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace CoolWord.Modules.Printing
{
    /// <summary>
    /// Bộ máy xử lý các tác vụ in ấn và xuất bản thông minh cho Microsoft Word.
    /// Hỗ trợ in 2 mặt thủ công, in trang chẵn/lẻ, in nhanh trang/vùng chọn, xuất PDF
    /// và in chứng từ kế toán nhiều liên (tính năng nghiệp vụ kế toán chuyên sâu).
    /// </summary>
    public static class PrintService
    {
        private static object Missing = Type.Missing;
        private static object TrueObj = true;
        private static object FalseObj = false;

        /// <summary>
        /// In nhanh trang hiện tại (nơi con trỏ đang đứng).
        /// </summary>
        public static bool PrintCurrentPage(Word.Application app, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                Word.Document doc = app.ActiveDocument;
                object range = Word.WdPrintOutRange.wdPrintCurrentPage;
                object copies = 1;
                object pageType = Word.WdPrintOutPages.wdPrintAllPages;

                doc.PrintOut(
                    ref FalseObj, ref Missing, ref range, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref copies,
                    ref Missing, ref pageType, ref FalseObj, ref TrueObj,
                    ref Missing, ref FalseObj, ref Missing, ref Missing,
                    ref Missing, ref Missing);

                statusMessage = "Đã gửi lệnh in trang hiện tại thành công.";
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi in trang hiện tại: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// In vùng văn bản hoặc bảng biểu đang được bôi đen.
        /// </summary>
        public static bool PrintSelection(Word.Application app, out string statusMessage)
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
                    statusMessage = "Vui lòng bôi đen vùng nội dung muốn in trước khi thực hiện.";
                    return false;
                }

                Word.Document doc = app.ActiveDocument;
                object range = Word.WdPrintOutRange.wdPrintSelection;
                object copies = 1;
                object pageType = Word.WdPrintOutPages.wdPrintAllPages;

                doc.PrintOut(
                    ref FalseObj, ref Missing, ref range, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref copies,
                    ref Missing, ref pageType, ref FalseObj, ref TrueObj,
                    ref Missing, ref FalseObj, ref Missing, ref Missing,
                    ref Missing, ref Missing);

                statusMessage = "Đã gửi lệnh in vùng văn bản đang chọn thành công.";
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi in vùng chọn: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Chỉ in các trang lẻ (1, 3, 5...).
        /// </summary>
        public static bool PrintOddPagesOnly(Word.Application app, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                Word.Document doc = app.ActiveDocument;
                object range = Word.WdPrintOutRange.wdPrintAllDocument;
                object copies = 1;
                object pageType = Word.WdPrintOutPages.wdPrintOddPagesOnly;

                doc.PrintOut(
                    ref FalseObj, ref Missing, ref range, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref copies,
                    ref Missing, ref pageType, ref FalseObj, ref TrueObj,
                    ref Missing, ref FalseObj, ref Missing, ref Missing,
                    ref Missing, ref Missing);

                statusMessage = "Đã gửi lệnh in toàn bộ các trang lẻ (1, 3, 5...).";
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi in trang lẻ: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Chỉ in các trang chẵn (2, 4, 6...).
        /// </summary>
        public static bool PrintEvenPagesOnly(Word.Application app, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                Word.Document doc = app.ActiveDocument;
                object range = Word.WdPrintOutRange.wdPrintAllDocument;
                object copies = 1;
                object pageType = Word.WdPrintOutPages.wdPrintEvenPagesOnly;

                doc.PrintOut(
                    ref FalseObj, ref Missing, ref range, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref copies,
                    ref Missing, ref pageType, ref FalseObj, ref TrueObj,
                    ref Missing, ref FalseObj, ref Missing, ref Missing,
                    ref Missing, ref Missing);

                statusMessage = "Đã gửi lệnh in toàn bộ các trang chẵn (2, 4, 6...).";
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi in trang chẵn: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// In 2 mặt thủ công thông minh (Manual Duplex):
        /// Bước 1: In toàn bộ trang lẻ.
        /// Bước 2: Hiển thị thông báo hướng dẫn người dùng lật mặt giấy.
        /// Bước 3: In toàn bộ trang chẵn sau khi người dùng xác nhận đã nạp lại giấy.
        /// </summary>
        public static bool PrintManualDuplex(Word.Application app, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                // Bước 1: In trang lẻ
                string oddStatus;
                if (!PrintOddPagesOnly(app, out oddStatus))
                {
                    statusMessage = "Không thể bắt đầu in 2 mặt: " + oddStatus;
                    return false;
                }

                // Bước 2: Hộp thoại hướng dẫn người dùng
                DialogResult dr = MessageBox.Show(
                    "Đã gửi lệnh in các TRANG LẺ thành công.\n\n" +
                    "Vui lòng thực hiện các bước sau:\n" +
                    "1. Lấy toàn bộ xấp giấy vừa in ra khỏi khay ra của máy in.\n" +
                    "2. Giữ nguyên hướng đầu trang, lật úp mặt giấy lại và nạp vào khay chứa giấy.\n" +
                    "3. Nhấn 'OK' để tiếp tục in các TRANG CHẴN, hoặc 'Cancel' để dừng lại.",
                    "CoolWord - Hướng Dẫn In 2 Mặt Thủ Công",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Information);

                if (dr != DialogResult.OK)
                {
                    statusMessage = "Đã dừng in trang chẵn theo yêu cầu của người dùng.";
                    return true;
                }

                // Bước 3: In trang chẵn
                string evenStatus;
                if (!PrintEvenPagesOnly(app, out evenStatus))
                {
                    statusMessage = "Lỗi khi in các trang chẵn: " + evenStatus;
                    return false;
                }

                statusMessage = "Hoàn tất in 2 mặt thủ công (đã gửi lệnh in trang lẻ và trang chẵn).";
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi quy trình in 2 mặt: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// In 2 trang tài liệu trên 1 mặt giấy A4 nằm ngang (2 Pages per Sheet - N-up).
        /// </summary>
        public static bool PrintTwoPagesPerSheet(Word.Application app, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                Word.Document doc = app.ActiveDocument;
                object range = Word.WdPrintOutRange.wdPrintAllDocument;
                object copies = 1;
                object pageType = Word.WdPrintOutPages.wdPrintAllPages;
                object zoomCol = 2; // 2 cột
                object zoomRow = 1; // 1 dòng

                doc.PrintOut(
                    ref FalseObj, ref Missing, ref range, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref copies,
                    ref Missing, ref pageType, ref FalseObj, ref TrueObj,
                    ref Missing, ref FalseObj, ref zoomCol, ref zoomRow,
                    ref Missing, ref Missing);

                statusMessage = "Đã gửi lệnh in 2 trang trên 1 mặt giấy A4.";
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi in 2 trang/mặt giấy: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Xuất tài liệu hiện tại sang định dạng PDF chuẩn in ấn.
        /// </summary>
        public static bool ExportQuickPdf(Word.Application app, string targetPath, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                Word.Document doc = app.ActiveDocument;
                string destFile = targetPath;
                if (string.IsNullOrWhiteSpace(destFile))
                {
                    string docPath = string.Empty;
                    try { docPath = doc.FullName; } catch { }

                    if (!string.IsNullOrEmpty(docPath) && File.Exists(docPath))
                    {
                        string dir = Path.GetDirectoryName(docPath);
                        string fileNameNoExt = Path.GetFileNameWithoutExtension(docPath);
                        destFile = Path.Combine(dir, fileNameNoExt + ".pdf");
                    }
                    else
                    {
                        string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                        destFile = Path.Combine(desktop, "CoolWord_Export_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".pdf");
                    }
                }

                doc.ExportAsFixedFormat(
                    destFile,
                    Word.WdExportFormat.wdExportFormatPDF,
                    OpenAfterExport: true,
                    OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint);

                statusMessage = "Đã xuất PDF thành công: " + Path.GetFileName(destFile);
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi xuất PDF: " + ex.Message;
                return false;
            }
        }

        #region Nghiệp vụ Kế toán (Tính năng in ấn kế toán chuyên biệt - KHÔNG hiển thị trên giao diện người dùng)

        /// <summary>
        /// TÍNH NĂNG KẾ TOÁN: In chứng từ kế toán nhiều liên (Phiếu Thu - Chi / Hóa đơn 2 liên - 3 liên)
        /// Hỗ trợ nghiệp vụ kế toán khi in các biểu mẫu phiếu thu, phiếu chi, phiếu nhập kho, phiếu xuất kho:
        /// - In tự động 2 liên (Liên 1: Gốc lưu kế toán, Liên 2: Giao người nộp/nhận)
        /// - Hoặc 3 liên (Liên 1: Gốc lưu, Liên 2: Giao khách, Liên 3: Thủ kho/Thủ quỹ)
        /// - Hỗ trợ in ghép 2 liên trên cùng 1 trang A4 nằm ngang để tiết kiệm giấy kế toán.
        /// LƯU Ý: Theo yêu cầu nghiệp vụ, tính năng này hoạt động độc lập ở tầng logic/dịch vụ và
        /// KHÔNG hiển thị nút bấm hay thông tin trực tiếp trên giao diện Ribbon / Task Pane.
        /// </summary>
        public static bool PrintMultiCopyAccountingVoucher(
            Word.Application app,
            int copies,
            string voucherType,
            bool fitTwoPerSheet,
            out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            if (copies <= 0) copies = 2;
            string vName = string.IsNullOrWhiteSpace(voucherType) ? "Chứng từ kế toán" : voucherType.Trim();

            try
            {
                Word.Document doc = app.ActiveDocument;
                object range = Word.WdPrintOutRange.wdPrintAllDocument;
                object copiesObj = copies;
                object pageType = Word.WdPrintOutPages.wdPrintAllPages;
                object zoomCol = fitTwoPerSheet ? 2 : (object)Missing;
                object zoomRow = fitTwoPerSheet ? 1 : (object)Missing;

                doc.PrintOut(
                    ref FalseObj, ref Missing, ref range, ref Missing,
                    ref Missing, ref Missing, ref Missing, ref copiesObj,
                    ref Missing, ref pageType, ref FalseObj, ref TrueObj,
                    ref Missing, ref FalseObj, ref zoomCol, ref zoomRow,
                    ref Missing, ref Missing);

                string layoutDetail = fitTwoPerSheet ? "ghép 2 liên trên 1 mặt A4" : "in liên tiếp các liên";
                statusMessage = string.Format(
                    "Đã in {0} liên {1} ({2}) thành công (Liên 1: Lưu gốc kế toán, Liên 2: Giao đối tác...).",
                    copies, vName, layoutDetail);
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi in chứng từ kế toán nhiều liên: " + ex.Message;
                return false;
            }
        }

        #endregion
    }
}
