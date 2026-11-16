using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using QRCoder;
using Word = Microsoft.Office.Interop.Word;

namespace CoolWord.Modules.QRCode
{
    public class VietQrBankInfo
    {
        public string ShortName { get; set; }
        public string FullName { get; set; }
        public string Bin { get; set; }

        public VietQrBankInfo(string shortName, string fullName, string bin)
        {
            ShortName = shortName;
            FullName = fullName;
            Bin = bin;
        }

        public override string ToString()
        {
            return string.Format("{0} - {1}", ShortName, FullName);
        }
    }

    /// <summary>
    /// Dịch vụ tạo và chèn mã QR Code vào tài liệu Word:
    /// - QR Code từ URL / Văn bản bất kỳ
    /// - VietQR chuẩn thanh toán ngân hàng NAPAS 247
    /// - vCard Danh thiếp điện tử
    /// - Tra cứu văn bản hành chính theo NĐ 30/2020
    /// </summary>
    public static class QRCodeService
    {
        private static object Missing = Type.Missing;

        public static readonly List<VietQrBankInfo> PopularBanks = new List<VietQrBankInfo>
        {
            new VietQrBankInfo("VCB", "Vietcombank (Ngân hàng Ngoại thương)", "970436"),
            new VietQrBankInfo("MB", "MBBank (Ngân hàng Quân đội)", "970422"),
            new VietQrBankInfo("TCB", "Techcombank (Ngân hàng Kỹ Thương)", "970407"),
            new VietQrBankInfo("BIDV", "BIDV (Ngân hàng Đầu tư và Phát triển)", "970418"),
            new VietQrBankInfo("CTG", "VietinBank (Ngân hàng Công Thương)", "970415"),
            new VietQrBankInfo("ACB", "ACB (Ngân hàng Á Châu)", "970416"),
            new VietQrBankInfo("VPB", "VPBank (Ngân hàng Việt Nam Thịnh Vượng)", "970432"),
            new VietQrBankInfo("TPB", "TPBank (Ngân hàng Tiên Phong)", "970423"),
            new VietQrBankInfo("VBA", "Agribank (Ngân hàng Nông nghiệp & PTNT)", "970405"),
            new VietQrBankInfo("STB", "Sacombank (Ngân hàng Sài Gòn Thương Tín)", "970403"),
            new VietQrBankInfo("VIB", "VIB (Ngân hàng Quốc tế)", "970441"),
            new VietQrBankInfo("HDB", "HDBank (Ngân hàng Phát triển TP.HCM)", "970437"),
            new VietQrBankInfo("SHB", "SHB (Ngân hàng Sài Gòn - Hà Nội)", "970443"),
            new VietQrBankInfo("OCB", "OCB (Ngân hàng Phương Đông)", "970448"),
            new VietQrBankInfo("MSB", "MSB (Ngân hàng Hàng Hải)", "970426"),
            new VietQrBankInfo("LPB", "LPBank (Ngân hàng Bưu điện Liên Việt)", "970449"),
            new VietQrBankInfo("SEAB", "SeABank (Ngân hàng Đông Nam Á)", "970440"),
            new VietQrBankInfo("NAB", "Nam A Bank (Ngân hàng Nam Á)", "970428")
        };

        #region Sinh Ma QR Code (Bitmap)

        /// <summary>
        /// Tạo hình ảnh QR Code chất lượng cao dạng Bitmap.
        /// </summary>
        public static Bitmap GenerateQrBitmap(string content, int pixelsPerModule = 10, Color? darkColor = null, Color? lightColor = null, Bitmap logo = null)
        {
            if (string.IsNullOrEmpty(content)) return null;

            Color dark = darkColor ?? Color.Black;
            Color light = lightColor ?? Color.White;
            QRCodeGenerator.ECCLevel ecc = (logo != null) ? QRCodeGenerator.ECCLevel.H : QRCodeGenerator.ECCLevel.M;

            using (QRCodeGenerator generator = new QRCodeGenerator())
            using (QRCodeData data = generator.CreateQrCode(content, ecc))
            using (QRCoder.QRCode qrCode = new QRCoder.QRCode(data))
            {
                if (logo != null)
                {
                    return qrCode.GetGraphic(pixelsPerModule, dark, light, logo, 18, 6, true);
                }
                return qrCode.GetGraphic(pixelsPerModule, dark, light, true);
            }
        }

        #endregion

        #region Chèn QR Code vào Tài liệu Word

        /// <summary>
        /// Chèn hình ảnh QR Code vào vị trí con trỏ (hoặc thay thế vùng chọn) trong tài liệu Word.
        /// </summary>
        public static bool InsertQrCode(
            Word.Application app,
            string content,
            float widthPt,
            bool center,
            string caption,
            out string statusMessage)
        {
            return InsertQrCode(app, content, widthPt, center, caption, null, null, out statusMessage);
        }

        public static bool InsertQrCode(
            Word.Application app,
            string content,
            float widthPt,
            bool center,
            string caption,
            Color? darkColor,
            Color? lightColor,
            out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(content))
            {
                statusMessage = "Nội dung mã QR Code không được để trống.";
                return false;
            }

            string tempFile = null;
            try
            {
                using (Bitmap bmp = GenerateQrBitmap(content, 12, darkColor, lightColor))
                {
                    if (bmp == null)
                    {
                        statusMessage = "Không thể tạo mã QR Code từ nội dung đã cung cấp.";
                        return false;
                    }

                    tempFile = Path.Combine(Path.GetTempPath(), "CoolWord_QR_" + Guid.NewGuid().ToString("N") + ".png");
                    bmp.Save(tempFile, ImageFormat.Png);
                }

                Word.Selection sel = app.Selection;
                Word.Range insertRange = (sel != null && sel.Range != null) ? sel.Range : app.ActiveDocument.Content;

                object linkToFile = false;
                object saveWithDoc = true;
                object rangeObj = insertRange;

                Word.InlineShape shape = app.ActiveDocument.InlineShapes.AddPicture(
                    tempFile,
                    ref linkToFile,
                    ref saveWithDoc,
                    ref rangeObj);

                if (shape != null)
                {
                    if (widthPt > 0)
                    {
                        shape.Width = widthPt;
                        shape.Height = widthPt;
                    }

                    if (center)
                    {
                        try
                        {
                            shape.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                        catch { }
                    }

                    if (!string.IsNullOrWhiteSpace(caption))
                    {
                        try
                        {
                            Word.Range capRange = shape.Range.Duplicate;
                            object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                            capRange.Collapse(ref collapseEnd);

                            capRange.Text = "\r" + caption;
                            capRange.Font.Size = 9.5f;
                            capRange.Font.Italic = 1;
                            capRange.Font.ColorIndex = Word.WdColorIndex.wdGray50;
                            capRange.ParagraphFormat.Alignment = center
                                ? Word.WdParagraphAlignment.wdAlignParagraphCenter
                                : Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        }
                        catch { }
                    }
                }

                statusMessage = "Đã chèn mã QR Code thành công vào tài liệu.";
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi chèn mã QR Code: " + ex.Message;
                return false;
            }
            finally
            {
                if (!string.IsNullOrEmpty(tempFile) && File.Exists(tempFile))
                {
                    try { File.Delete(tempFile); } catch { }
                }
            }
        }

        /// <summary>
        /// Tạo và chèn nhanh QR Code từ đoạn văn bản hoặc URL đang bôi đen trong Word.
        /// </summary>
        public static bool InsertQrFromSelection(
            Word.Application app,
            float widthPt,
            bool center,
            string caption,
            out string statusMessage)
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
                    statusMessage = "Vui lòng bôi đen đường link hoặc đoạn văn bản cần tạo mã QR Code.";
                    return false;
                }

                string text = sel.Range.Text;
                if (string.IsNullOrWhiteSpace(text))
                {
                    statusMessage = "Vùng văn bản đang chọn không có nội dung.";
                    return false;
                }

                text = text.Trim();
                // Bỏ ký tự kết thúc đoạn \r\a nếu bôi đen cả ô bảng
                text = text.TrimEnd('\r', '\n', '\a', '\t');

                string cap = caption;
                if (string.IsNullOrEmpty(cap) && (text.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || text.StartsWith("https://", StringComparison.OrdinalIgnoreCase)))
                {
                    cap = "Quét mã để truy cập liên kết";
                }

                return InsertQrCode(app, text, widthPt, center, cap, null, null, out statusMessage);
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi tạo QR Code từ vùng chọn: " + ex.Message;
                return false;
            }
        }

        #endregion

        #region VietQR - Chuẩn Thanh Toán Ngân Hàng Việt Nam

        /// <summary>
        /// Tạo chuỗi EMVCo VietQR chuẩn NAPAS 247 cho thanh toán ngân hàng Việt Nam.
        /// Cho phép quét chuyển khoản trực tiếp trên mọi App ngân hàng (Vietcombank, MBBank, Techcombank...).
        /// </summary>
        public static string GenerateVietQrPayload(string bankBin, string accountNo, string accountName, decimal amount = 0, string description = "")
        {
            if (string.IsNullOrEmpty(bankBin) || string.IsNullOrEmpty(accountNo))
                return string.Empty;

            string cleanBin = bankBin.Trim();
            string cleanAcc = accountNo.Trim();
            string cleanName = string.IsNullOrEmpty(accountName) ? string.Empty : RemoveDiacritics(accountName.Trim().ToUpperInvariant());
            string cleanDesc = string.IsNullOrEmpty(description) ? string.Empty : RemoveDiacritics(description.Trim());

            // 1. Tag 38: Consumer Account Information (Napas VietQR)
            // 38.00 = A000000727 (Napas AID)
            // 38.01 = 00 <bin.length> <bin> 01 <acc.length> <acc>
            // 38.02 = QRIBFTTA (Service code: Quick Transfer to Account)
            string tag38_00 = FormatTlv("00", "A000000727");
            string tag38_01_inner = FormatTlv("00", cleanBin) + FormatTlv("01", cleanAcc);
            string tag38_01 = FormatTlv("01", tag38_01_inner);
            string tag38_02 = FormatTlv("02", "QRIBFTTA");
            string tag38 = FormatTlv("38", tag38_00 + tag38_01 + tag38_02);

            // 2. Tag 00: Payload Format Indicator ("01")
            string tag00 = FormatTlv("00", "01");

            // 3. Tag 01: Point of Initiation Method ("12" - Dynamic QR nếu có số tiền, "11" nếu tĩnh)
            string tag01 = FormatTlv("01", amount > 0 ? "12" : "11");

            // 4. Tag 53: Transaction Currency (VND = 704)
            string tag53 = FormatTlv("53", "704");

            // 5. Tag 54: Transaction Amount
            string tag54 = string.Empty;
            if (amount > 0)
            {
                tag54 = FormatTlv("54", ((long)amount).ToString());
            }

            // 6. Tag 58: Country Code ("VN")
            string tag58 = FormatTlv("58", "VN");

            // 7. Tag 59: Merchant Name
            string tag59 = string.Empty;
            if (!string.IsNullOrEmpty(cleanName))
            {
                tag59 = FormatTlv("59", cleanName.Length > 25 ? cleanName.Substring(0, 25) : cleanName);
            }

            // 8. Tag 60: Merchant City ("HA NOI")
            string tag60 = FormatTlv("60", "HA NOI");

            // 9. Tag 62: Additional Data Field (Purpose of transaction)
            string tag62 = string.Empty;
            if (!string.IsNullOrEmpty(cleanDesc))
            {
                string tag62_08 = FormatTlv("08", cleanDesc.Length > 25 ? cleanDesc.Substring(0, 25) : cleanDesc);
                tag62 = FormatTlv("62", tag62_08);
            }

            // Kết hợp chuỗi trước khi tính CRC
            string raw = tag00 + tag01 + tag38 + tag53 + tag54 + tag58 + tag59 + tag60 + tag62 + "6304";
            string crc = CalculateCrc16Ccitt(raw);

            return raw + crc;
        }

        private static string FormatTlv(string tag, string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            return string.Format("{0}{1:D2}{2}", tag, Encoding.UTF8.GetByteCount(value), value);
        }

        private static string CalculateCrc16Ccitt(string input)
        {
            byte[] bytes = Encoding.ASCII.GetBytes(input);
            ushort crc = 0xFFFF;
            for (int i = 0; i < bytes.Length; i++)
            {
                crc ^= (ushort)(bytes[i] << 8);
                for (int j = 0; j < 8; j++)
                {
                    if ((crc & 0x8000) != 0)
                        crc = (ushort)((crc << 1) ^ 0x1021);
                    else
                        crc = (ushort)(crc << 1);
                }
            }
            return crc.ToString("X4");
        }

        private static string RemoveDiacritics(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            string normalized = text.Normalize(NormalizationForm.FormD);
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < normalized.Length; i++)
            {
                System.Globalization.UnicodeCategory c = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(normalized[i]);
                if (c != System.Globalization.UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(normalized[i]);
                }
            }
            return sb.ToString().Normalize(NormalizationForm.FormC).Replace('đ', 'd').Replace('Đ', 'D');
        }

        #endregion

        #region vCard & Tra Cứu Văn Bản

        /// <summary>
        /// Tạo nội dung vCard 3.0 chuẩn cho danh thiếp điện tử.
        /// </summary>
        public static string GenerateVCardPayload(string fullName, string phone, string email, string title, string company, string website)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("BEGIN:VCARD");
            sb.AppendLine("VERSION:3.0");
            if (!string.IsNullOrWhiteSpace(fullName)) sb.AppendLine("FN:" + fullName.Trim());
            if (!string.IsNullOrWhiteSpace(title)) sb.AppendLine("TITLE:" + title.Trim());
            if (!string.IsNullOrWhiteSpace(company)) sb.AppendLine("ORG:" + company.Trim());
            if (!string.IsNullOrWhiteSpace(phone)) sb.AppendLine("TEL;TYPE=CELL:" + phone.Trim());
            if (!string.IsNullOrWhiteSpace(email)) sb.AppendLine("EMAIL:" + email.Trim());
            if (!string.IsNullOrWhiteSpace(website)) sb.AppendLine("URL:" + website.Trim());
            sb.Append("END:VCARD");
            return sb.ToString();
        }

        /// <summary>
        /// Tạo nội dung tra cứu văn bản hành chính theo quy chuẩn NĐ 30/2020/NĐ-CP.
        /// </summary>
        public static string GenerateAdminDocPayload(string docNumber, string docDate, string issuingAuthority, string title, string verifyUrl)
        {
            if (!string.IsNullOrWhiteSpace(verifyUrl))
            {
                return verifyUrl.Trim();
            }

            StringBuilder sb = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(docNumber)) sb.AppendLine("Số: " + docNumber.Trim());
            if (!string.IsNullOrWhiteSpace(docDate)) sb.AppendLine("Ngày: " + docDate.Trim());
            if (!string.IsNullOrWhiteSpace(issuingAuthority)) sb.AppendLine("Cơ quan: " + issuingAuthority.Trim());
            if (!string.IsNullOrWhiteSpace(title)) sb.AppendLine("Trích yếu: " + title.Trim());
            return sb.ToString().TrimEnd();
        }

        /// <summary>
        /// Lưu mã QR Code thành tệp hình ảnh PNG ra máy tính.
        /// </summary>
        public static bool SaveQrCodeAsImage(string content, string filePath, int pixelsPerModule, Color? darkColor, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (string.IsNullOrWhiteSpace(content))
            {
                statusMessage = "Nội dung QR Code không được để trống.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(filePath))
            {
                statusMessage = "Đường dẫn tệp tin lưu không hợp lệ.";
                return false;
            }

            try
            {
                using (Bitmap bmp = GenerateQrBitmap(content, pixelsPerModule, darkColor, null))
                {
                    if (bmp == null)
                    {
                        statusMessage = "Không thể tạo mã QR Code.";
                        return false;
                    }

                    string dir = Path.GetDirectoryName(filePath);
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }

                    bmp.Save(filePath, ImageFormat.Png);
                    statusMessage = "Đã lưu mã QR Code thành công: " + Path.GetFileName(filePath);
                    return true;
                }
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi lưu ảnh QR Code: " + ex.Message;
                return false;
            }
        }

        #endregion
    }
}
