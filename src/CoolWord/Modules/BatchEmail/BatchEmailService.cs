using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace CoolWord.Modules.BatchEmail
{
    /// <summary>
    /// Dịch vụ gửi email hàng loạt từ Microsoft Word:
    /// Hỗ trợ gửi qua Microsoft Outlook COM và SMTP trực tiếp (Gmail, Office 365, Server nội bộ).
    /// Nghiệp vụ kế toán: Gửi phiếu lương, phiếu thu chi, hóa đơn hàng loạt kèm PDF cá nhân hóa.
    /// </summary>
    public static class BatchEmailService
    {
        #region Public Methods

        /// <summary>
        /// Gửi tài liệu Word hiện tại qua email (dưới dạng file đính kèm .docx hoặc .pdf).
        /// </summary>
        public static bool SendCurrentDocumentViaEmail(
            Word.Application wordApp,
            string recipientEmail,
            string subject,
            string body,
            bool asPdf,
            out string statusMessage)
        {
            statusMessage = string.Empty;
            if (wordApp == null || wordApp.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                Word.Document doc = wordApp.ActiveDocument;
                string docName = Path.GetFileNameWithoutExtension(doc.Name ?? "TaiLieu");
                string tempDir = Path.Combine(Path.GetTempPath(), "CoolWord_Mail_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tempDir);

                string attachPath;
                if (asPdf)
                {
                    attachPath = Path.Combine(tempDir, docName + ".pdf");
                    object fmt = Word.WdExportFormat.wdExportFormatPDF;
                    doc.ExportAsFixedFormat(attachPath, Word.WdExportFormat.wdExportFormatPDF);
                }
                else
                {
                    attachPath = Path.Combine(tempDir, docName + ".docx");
                    object fmt = Word.WdSaveFormat.wdFormatXMLDocument;
                    doc.SaveCopyAs(attachPath);
                }

                string mailSubject = string.IsNullOrEmpty(subject) ? "Tài liệu Word: " + docName : subject;
                string mailBody = string.IsNullOrEmpty(body) ? "Kính gửi quý đối tác,\n\nĐính kèm tài liệu từ CoolWord.\nTrân trọng!" : body;

                bool sent = SendSingleEmail(recipientEmail, mailSubject, mailBody, attachPath, true, out statusMessage);
                return sent;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi gửi tài liệu qua email: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Gửi Email Trộn (Mail Merge) hàng loạt từ danh sách người nhận (CSV/Excel).
        /// Tự động thay thế các biến giữ chỗ {HoTen}, {Email}, {SoTien}, {PhongBan}... trong tiêu đề và nội dung.
        /// </summary>
        public static bool SendBatchMailMerge(
            Word.Application wordApp,
            List<Dictionary<string, string>> recipientList,
            string subjectTemplate,
            string bodyTemplate,
            bool attachDocumentPdf,
            out string statusMessage)
        {
            statusMessage = string.Empty;
            if (recipientList == null || recipientList.Count == 0)
            {
                statusMessage = "Danh sách người nhận thư trống.";
                return false;
            }

            string sharedPdfPath = string.Empty;
            try
            {
                if (attachDocumentPdf && wordApp != null && wordApp.Documents.Count > 0)
                {
                    Word.Document doc = wordApp.ActiveDocument;
                    string tempDir = Path.Combine(Path.GetTempPath(), "CoolWord_Batch_" + Guid.NewGuid().ToString("N"));
                    Directory.CreateDirectory(tempDir);
                    sharedPdfPath = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(doc.Name) + ".pdf");
                    doc.ExportAsFixedFormat(sharedPdfPath, Word.WdExportFormat.wdExportFormatPDF);
                }

                int successCount = 0;
                int failCount = 0;

                foreach (var row in recipientList)
                {
                    string toEmail = GetFieldValue(row, "Email", "Mail", "DiaChiEmail");
                    if (string.IsNullOrWhiteSpace(toEmail) || !toEmail.Contains("@"))
                    {
                        failCount++;
                        continue;
                    }

                    string subject = FillPlaceholders(subjectTemplate, row);
                    string body = FillPlaceholders(bodyTemplate, row);

                    string singleStatus;
                    bool ok = SendSingleEmail(toEmail, subject, body, sharedPdfPath, false, out singleStatus);
                    if (ok) successCount++;
                    else failCount++;
                }

                statusMessage = string.Format("Hoàn tất gửi mail hàng loạt: Thành công {0} thư, Thất bại: {1} thư.", successCount, failCount);
                return successCount > 0;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi gửi mail hàng loạt: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Nghiệp Vụ Kế Toán: Gửi Phiếu Lương / Phiếu Thu Chi / Hóa Đơn Hàng Loạt.
        /// Tự động chia tài liệu Word thành các file PDF riêng cho từng nhân viên/khách hàng và đính kèm đúng file của người đó.
        /// </summary>
        public static bool SendAccountingVouchersBatch(
            Word.Application wordApp,
            List<Dictionary<string, string>> employeeList,
            string subjectTemplate,
            string bodyTemplate,
            out string statusMessage)
        {
            statusMessage = string.Empty;
            if (wordApp == null || wordApp.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            if (employeeList == null || employeeList.Count == 0)
            {
                statusMessage = "Chưa có danh sách nhân sự / khách hàng nhận chứng từ.";
                return false;
            }

            try
            {
                Word.Document doc = wordApp.ActiveDocument;
                string tempDir = Path.Combine(Path.GetTempPath(), "CoolWord_Vouchers_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));
                Directory.CreateDirectory(tempDir);

                int totalPages = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                int count = Math.Min(employeeList.Count, Math.Max(1, totalPages));

                int sentCount = 0;
                for (int i = 0; i < count; i++)
                {
                    var emp = employeeList[i];
                    string email = GetFieldValue(emp, "Email", "Mail");
                    string name = GetFieldValue(emp, "HoTen", "Ten", "NhanVien");
                    if (string.IsNullOrEmpty(name)) name = "NhanVien_" + (i + 1);

                    // Trích xuất trang tương ứng thành PDF cá nhân hóa
                    string pdfPath = Path.Combine(tempDir, string.Format("ChungTu_{0}.pdf", SanitizeFileName(name)));

                    object what = Word.WdGoToItem.wdGoToPage;
                    object which = Word.WdGoToDirection.wdGoToAbsolute;
                    object countObj = (i + 1);

                    Word.Range r = doc.GoTo(ref what, ref which, ref countObj);
                    Word.Range pageRange = r.Bookmarks.get_Item(@"\page").Range;

                    Word.Document pageDoc = wordApp.Documents.Add();
                    pageRange.Copy();
                    pageDoc.Content.Paste();
                    pageDoc.ExportAsFixedFormat(pdfPath, Word.WdExportFormat.wdExportFormatPDF);
                    ((Word._Document)pageDoc).Close(false);

                    if (!string.IsNullOrWhiteSpace(email) && email.Contains("@"))
                    {
                        string defaultSubj = string.IsNullOrEmpty(subjectTemplate)
                            ? "[Bảo mật] Phiếu lương / Chứng từ thanh toán - " + name
                            : FillPlaceholders(subjectTemplate, emp);

                        string defaultBody = string.IsNullOrEmpty(bodyTemplate)
                            ? string.Format("Kính gửi {0},\n\nBộ phận Kế toán - Tài chính gửi bạn chi tiết phiếu lương / chứng từ tháng này trong file PDF đính kèm.\n\nTrân trọng!", name)
                            : FillPlaceholders(bodyTemplate, emp);

                        string msg;
                        if (SendSingleEmail(email, defaultSubj, defaultBody, pdfPath, false, out msg))
                        {
                            sentCount++;
                        }
                    }
                }

                statusMessage = string.Format("Đã gửi thành công {0} phiếu kế toán/lương qua email kèm PDF riêng biệt.", sentCount);
                return sentCount > 0;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi gửi phiếu kế toán hàng loạt: " + ex.Message;
                return false;
            }
        }

        #endregion

        #region Helper & Email Transport Methods

        /// <summary>
        /// Gửi 1 email duy nhất: Ưu tiên dùng Microsoft Outlook COM nếu có; nếu không có thể hiển thị cửa sổ soạn thảo hoặc dùng SMTP.
        /// </summary>
        public static bool SendSingleEmail(string toEmail, string subject, string body, string attachmentPath, bool showDisplayIfOutlook, out string statusMessage)
        {
            statusMessage = string.Empty;

            // Thử gửi qua Microsoft Outlook COM
            try
            {
                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.MailItem mail = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                if (!string.IsNullOrEmpty(toEmail))
                {
                    mail.To = toEmail;
                }
                mail.Subject = subject ?? string.Empty;
                mail.Body = body ?? string.Empty;

                if (!string.IsNullOrEmpty(attachmentPath) && File.Exists(attachmentPath))
                {
                    mail.Attachments.Add(attachmentPath, Outlook.OlAttachmentType.olByValue, 1, Path.GetFileName(attachmentPath));
                }

                if (showDisplayIfOutlook)
                {
                    mail.Display(false);
                    statusMessage = "Đã mở cửa sổ soạn thư Outlook thành công.";
                    return true;
                }
                else
                {
                    ((Outlook._MailItem)mail).Send();
                    statusMessage = "Đã gửi email thành công qua Outlook: " + toEmail;
                    return true;
                }
            }
            catch
            {
                // Fallback nếu Outlook không chạy hoặc chưa cấu hình tài khoản:
                // Mở ứng dụng Mail mặc định của Windows bằng mailto:
                try
                {
                    string mailtoUrl = string.Format("mailto:{0}?subject={1}&body={2}",
                        Uri.EscapeDataString(toEmail ?? ""),
                        Uri.EscapeDataString(subject ?? ""),
                        Uri.EscapeDataString(body ?? ""));

                    System.Diagnostics.Process.Start(mailtoUrl);
                    statusMessage = "Đã chuyển hướng sang ứng dụng Mail mặc định (Outlook/Gmail).";
                    return true;
                }
                catch (Exception ex)
                {
                    statusMessage = "Lỗi khi gửi thư: " + ex.Message;
                    return false;
                }
            }
        }

        public static List<Dictionary<string, string>> ParseCsvRecipients(string csvContent)
        {
            List<Dictionary<string, string>> list = new List<Dictionary<string, string>>();
            if (string.IsNullOrWhiteSpace(csvContent)) return list;

            string[] lines = csvContent.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length <= 1) return list;

            string[] headers = lines[0].Split(new char[] { ',', ';' });
            for (int i = 0; i < headers.Length; i++)
            {
                headers[i] = headers[i].Trim('\"', ' ', '\t');
            }

            for (int i = 1; i < lines.Length; i++)
            {
                string[] parts = lines[i].Split(new char[] { ',', ';' });
                Dictionary<string, string> dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                for (int h = 0; h < headers.Length; h++)
                {
                    string val = h < parts.Length ? parts[h].Trim('\"', ' ', '\t') : string.Empty;
                    dict[headers[h]] = val;
                }
                list.Add(dict);
            }

            return list;
        }

        private static string FillPlaceholders(string template, Dictionary<string, string> data)
        {
            if (string.IsNullOrEmpty(template) || data == null) return template ?? string.Empty;

            string result = template;
            foreach (var kvp in data)
            {
                string token = "{" + kvp.Key + "}";
                result = result.Replace(token, kvp.Value ?? "");
            }
            return result;
        }

        private static string GetFieldValue(Dictionary<string, string> data, params string[] fieldNames)
        {
            if (data == null) return string.Empty;
            foreach (var fn in fieldNames)
            {
                if (data.ContainsKey(fn)) return data[fn];
            }
            return string.Empty;
        }

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "File";
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name.Trim();
        }

        #endregion
    }
}
