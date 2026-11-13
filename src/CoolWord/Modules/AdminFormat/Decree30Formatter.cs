using System;
using System.Text.RegularExpressions;
using CoolWord.Modules.NumberToText;
using CoolWord.Modules.TextUtilities;
using Word = Microsoft.Office.Interop.Word;

namespace CoolWord.Modules.AdminFormat
{
    /// <summary>
    /// Bộ máy chuẩn hóa thể thức và kỹ thuật trình bày văn bản hành chính
    /// theo đúng quy định tại Phụ lục I - Nghị định 30/2020/NĐ-CP của Chính phủ.
    /// (Được thiết kế và viết mới hoàn toàn độc lập cho CoolWord).
    /// </summary>
    public static class Decree30Formatter
    {
        public static float MmToPoints(double mm)
        {
            return (float)(mm * 72.0 / 25.4);
        }

        public static float CmToPoints(double cm)
        {
            return (float)(cm * 72.0 / 2.54);
        }

        /// <summary>
        /// Định dạng chuỗi địa danh và ngày tháng ban hành văn bản theo đúng Phụ lục I - Nghị định 30/2020/NĐ-CP:
        /// Thêm chữ số 0 phía trước các ngày nhỏ hơn 10 và các tháng 1, 2.
        /// </summary>
        public static string FormatAdminDateLine(string location, DateTime date)
        {
            string dayStr = date.Day < 10 ? "0" + date.Day : date.Day.ToString();
            string monthStr = (date.Month == 1 || date.Month == 2) ? "0" + date.Month : date.Month.ToString();
            string loc = string.IsNullOrWhiteSpace(location) ? "Hà Nội" : location.Trim();
            return string.Format("{0}, ngày {1} tháng {2} năm {3}", loc, dayStr, monthStr, date.Year);
        }

        /// <summary>
        /// Đọc ngày tháng năm hoàn toàn bằng chữ Tiếng Việt (dùng cho biên bản, hợp đồng, lời chứng).
        /// </summary>
        public static string FormatDateInVietnameseWords(DateTime date, string location = null)
        {
            NumberConvertOptions pureLower = new NumberConvertOptions
            {
                ReadingMode = NumberReadingMode.PureNumber,
                Dialect = VietnameseDialect.MixedCommon,
                Capitalization = TextCapitalization.LowerCase
            };

            string dayWords = VietnameseNumberConverter.Convert(date.Day.ToString(), pureLower);
            string monthWords = VietnameseNumberConverter.Convert(date.Month.ToString(), pureLower);
            string yearWords = VietnameseNumberConverter.Convert(date.Year.ToString(), pureLower);

            if (!string.IsNullOrWhiteSpace(location))
            {
                return string.Format("{0}, ngày {1} tháng {2} năm {3}", location.Trim(), dayWords, monthWords, yearWords);
            }

            return string.Format("Ngày {0} tháng {1} năm {2}", dayWords, monthWords, yearWords);
        }

        /// <summary>
        /// 1-Click: Chuẩn hóa toàn diện văn bản theo Nghị định 30/2020/NĐ-CP
        /// (Khổ A4 + Căn lề chuẩn + Phông Times New Roman 13/14pt + Đoạn văn + Dọn rác + Số trang).
        /// </summary>
        public static bool ApplyFullStandardOneClick(Word.Application app, AdminDocOptions opts, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            if (opts == null) opts = new AdminDocOptions();

            try
            {
                Word.Document doc = app.ActiveDocument;
                StartUndo(app, "CoolWord: Chuẩn hóa toàn diện NĐ 30/2020");
                try
                {
                    ApplyPageMargins(doc, opts);
                    int cleanedCount = CleanDocumentParagraphGarbage(doc);
                    int formattedParas = FormatBodyParagraphs(doc, opts);

                    if (opts.EnableDecree30PageNumbering)
                    {
                        ApplyDecree30PageNumbering(doc, opts);
                    }

                    statusMessage = string.Format(
                        "Đã chuẩn hóa NĐ 30/2020: Khổ A4 (Lề {0}-{1}-{2}-{3}mm), {4} đoạn văn ({5} {6}pt) & dọn {7} lỗi định dạng.",
                        opts.TopMarginMm, opts.BottomMarginMm, opts.LeftMarginMm, opts.RightMarginMm,
                        formattedParas, opts.FontName, opts.BodyFontSize, cleanedCount);
                    return true;
                }
                finally
                {
                    EndUndo(app);
                }
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi chuẩn hóa văn bản: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Thiết lập khổ giấy A4 và căn lề chuẩn Nghị định 30/2020/NĐ-CP.
        /// </summary>
        public static void ApplyPageMargins(Word.Document doc, AdminDocOptions opts)
        {
            if (doc == null || opts == null) return;

            Word.PageSetup ps = doc.PageSetup;
            ps.PaperSize = Word.WdPaperSize.wdPaperA4;
            ps.TopMargin = MmToPoints(opts.TopMarginMm);
            ps.BottomMargin = MmToPoints(opts.BottomMarginMm);
            ps.LeftMargin = MmToPoints(opts.LeftMarginMm);
            ps.RightMargin = MmToPoints(opts.RightMarginMm);
            ps.HeaderDistance = MmToPoints(10.0);
            ps.FooterDistance = MmToPoints(10.0);
        }

        /// <summary>
        /// Đánh số trang chuẩn Nghị định 30/2020/NĐ-CP:
        /// Đặt canh giữa ở lề trên (Header), chữ số Ả Rập, cỡ chữ 13-14pt Times New Roman, không hiển thị ở trang thứ nhất.
        /// </summary>
        public static void ApplyDecree30PageNumbering(Word.Document doc, AdminDocOptions opts)
        {
            if (doc == null || opts == null) return;

            try
            {
                foreach (Word.Section section in doc.Sections)
                {
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1; // True: Không đánh số ở trang đầu

                    Word.HeaderFooter primaryHeader = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    primaryHeader.Range.Text = string.Empty;

                    object alignment = Word.WdPageNumberAlignment.wdAlignPageNumberCenter;
                    object firstPage = false;
                    primaryHeader.PageNumbers.Add(ref alignment, ref firstPage);

                    primaryHeader.Range.Font.Name = opts.FontName;
                    primaryHeader.Range.Font.Size = 13.0f;
                    primaryHeader.Range.Font.Bold = 0;
                    primaryHeader.Range.Font.Italic = 0;
                }
            }
            catch { }
        }

        /// <summary>
        /// Chuẩn hóa phông chữ, cỡ chữ, giãn dòng và thụt lề dòng đầu cho toàn bộ đoạn văn nội dung (ngoài bảng).
        /// </summary>
        public static int FormatBodyParagraphs(Word.Document doc, AdminDocOptions opts)
        {
            if (doc == null || opts == null) return 0;

            int count = 0;
            float firstLinePt = CmToPoints(opts.FirstLineIndentCm);

            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                Word.Range r = para.Range;

                // Bỏ qua các đoạn nằm trong bảng biểu (Quốc hiệu, Nơi nhận - Chữ ký, bảng số liệu)
                bool inTable = false;
                try
                {
                    inTable = r.Tables.Count > 0;
                }
                catch { }

                if (inTable) continue;

                r.Font.Name = opts.FontName;
                r.Font.Size = opts.BodyFontSize;

                para.SpaceBefore = opts.SpaceBeforePt;
                para.SpaceAfter = opts.SpaceAfterPt;
                para.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                para.LineSpacing = opts.LineSpacingMultiple * 12.0f;

                // Nếu đoạn văn đang căn giữa (tiêu đề QUYẾT ĐỊNH, BÁO CÁO...) thì giữ nguyên căn giữa và không thụt đầu dòng
                if (para.Alignment == Word.WdParagraphAlignment.wdAlignParagraphCenter)
                {
                    para.FirstLineIndent = 0f;
                    para.LeftIndent = 0f;
                    para.RightIndent = 0f;
                }
                else
                {
                    para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    para.LeftIndent = 0f;
                    para.RightIndent = 0f;
                    para.FirstLineIndent = firstLinePt;
                }

                count++;
            }

            return count;
        }

        /// <summary>
        /// Dọn dẹp định dạng rác trong văn bản: xóa khoảng trắng thừa đầu/cuối đoạn, gộp dấu cách kép,
        /// xóa dấu Tab đầu dòng, chuẩn hóa dấu câu tiếng Việt và xóa dòng trống thừa liên tiếp.
        /// </summary>
        public static int CleanDocumentParagraphGarbage(Word.Document doc)
        {
            if (doc == null) return 0;

            int fixedCount = 0;
            bool previousWasEmpty = false;

            for (int i = doc.Paragraphs.Count; i >= 1; i--)
            {
                Word.Paragraph para = doc.Paragraphs[i];
                Word.Range r = para.Range;

                bool inTable = false;
                try
                {
                    inTable = r.Tables.Count > 0;
                }
                catch { }

                if (inTable)
                {
                    previousWasEmpty = false;
                    continue;
                }

                string raw = r.Text;
                if (string.IsNullOrEmpty(raw)) continue;

                string body = raw.TrimEnd('\r', '\n', '\a');
                if (string.IsNullOrWhiteSpace(body))
                {
                    if (previousWasEmpty && doc.Paragraphs.Count > 1)
                    {
                        try
                        {
                            r.Delete();
                            fixedCount++;
                        }
                        catch { }
                    }
                    previousWasEmpty = true;
                    continue;
                }

                previousWasEmpty = false;

                // Làm sạch khoảng trắng thừa, tab đầu dòng và chuẩn hóa dấu câu
                string cleaned = Regex.Replace(body, @"^[\s\t]+", "");
                cleaned = TextCleanModule.NormalizeVietnameseText(cleaned);

                if (!string.Equals(body, cleaned, StringComparison.Ordinal))
                {
                    Word.Range contentRange = r.Duplicate;
                    contentRange.End = contentRange.End - 1; // Giữ lại ký tự \r cuối đoạn để không mất định dạng đoạn
                    contentRange.Text = cleaned;
                    fixedCount++;
                }
            }

            return fixedCount;
        }

        /// <summary>
        /// Chèn khung Đầu Văn Bản chuẩn Nghị định 30/2020/NĐ-CP (Cơ quan ban hành, Quốc hiệu - Tiêu ngữ,
        /// Số ký hiệu, Địa danh ngày tháng, Tên loại văn bản & Trích yếu).
        /// </summary>
        public static bool InsertHeaderBlock(Word.Application app, AdminDocOptions opts, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            if (opts == null) opts = new AdminDocOptions();

            try
            {
                Word.Document doc = app.ActiveDocument;
                StartUndo(app, "CoolWord: Chèn Quốc hiệu & Tiêu ngữ NĐ 30/2020");
                try
                {
                    // Di chuyển lên đầu tài liệu để chèn khung Quốc hiệu - Tiêu ngữ
                    Word.Range startRange = doc.Range(0, 0);
                    startRange.InsertParagraphBefore();
                    Word.Range tableRange = doc.Range(0, 0);

                    Word.Table headerTable = doc.Tables.Add(tableRange, 2, 2);
                    headerTable.Borders.Enable = 0;
                    headerTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                    headerTable.PreferredWidth = 100f;

                    headerTable.Columns[1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                    headerTable.Columns[1].PreferredWidth = 42f;
                    headerTable.Columns[2].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                    headerTable.Columns[2].PreferredWidth = 58f;

                    // Ô (1, 1): Tên cơ quan chủ quản + Tên cơ quan ban hành + Đường kẻ ngang
                    Word.Cell cellOrg = headerTable.Cell(1, 1);
                    FormatHeaderOrgCell(cellOrg, opts);

                    // Ô (1, 2): Quốc hiệu + Tiêu ngữ + Đường kẻ ngang
                    Word.Cell cellQuocHieu = headerTable.Cell(1, 2);
                    FormatHeaderQuocHieuCell(cellQuocHieu, opts);

                    // Ô (2, 1): Số, ký hiệu văn bản
                    Word.Cell cellDocNo = headerTable.Cell(2, 1);
                    FormatSimpleTableCell(
                        cellDocNo,
                        string.IsNullOrWhiteSpace(opts.DocumentNumber) ? "Số:     /..." : opts.DocumentNumber.Trim(),
                        opts.FontName,
                        13.0f,
                        false,
                        false,
                        4.0f);

                    // Ô (2, 2): Địa danh và ngày tháng ban hành
                    Word.Cell cellDate = headerTable.Cell(2, 2);
                    string dateLine = FormatAdminDateLine(opts.LocationName, opts.IssueDate);
                    FormatSimpleTableCell(
                        cellDate,
                        dateLine,
                        opts.FontName,
                        opts.BodyFontSize,
                        false,
                        true,
                        4.0f);

                    // Chèn Tên loại văn bản & Trích yếu nội dung ngay bên dưới bảng
                    Word.Range afterTable = headerTable.Range.Duplicate;
                    object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                    afterTable.Collapse(ref collapseEnd);

                    if (!string.IsNullOrWhiteSpace(opts.DocumentTypeTitle))
                    {
                        string title = opts.DocumentTypeTitle.Trim().ToUpperInvariant();
                        string subject = (opts.DocumentSubject ?? string.Empty).Trim();

                        afterTable.InsertParagraphAfter();
                        afterTable.Collapse(ref collapseEnd);

                        afterTable.Text = title + "\r";
                        afterTable.Font.Name = opts.FontName;
                        afterTable.Font.Size = 14.0f;
                        afterTable.Font.Bold = 1;
                        afterTable.Font.Italic = 0;
                        afterTable.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        afterTable.ParagraphFormat.FirstLineIndent = 0f;
                        afterTable.ParagraphFormat.SpaceBefore = 12.0f;
                        afterTable.ParagraphFormat.SpaceAfter = 2.0f;

                        afterTable.Collapse(ref collapseEnd);

                        if (!string.IsNullOrEmpty(subject))
                        {
                            afterTable.Text = subject + "\r───────────────\r";
                            afterTable.Font.Name = opts.FontName;
                            afterTable.Font.Size = opts.BodyFontSize;
                            afterTable.Font.Bold = 1;
                            afterTable.Font.Italic = 0;
                            afterTable.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            afterTable.ParagraphFormat.FirstLineIndent = 0f;
                            afterTable.ParagraphFormat.SpaceBefore = 0f;
                            afterTable.ParagraphFormat.SpaceAfter = 6.0f;
                            afterTable.Collapse(ref collapseEnd);
                        }
                    }

                    statusMessage = "Đã chèn khung Quốc hiệu - Tiêu ngữ & Tiêu đề chuẩn Nghị định 30/2020/NĐ-CP.";
                    return true;
                }
                finally
                {
                    EndUndo(app);
                }
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi chèn khung Quốc hiệu: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Chèn khung Cuối Văn Bản chuẩn Nghị định 30/2020/NĐ-CP (Nơi nhận bên trái & Quyền hạn, Chức vụ, Chữ ký bên phải).
        /// </summary>
        public static bool InsertFooterSignatureBlock(Word.Application app, AdminDocOptions opts, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            if (opts == null) opts = new AdminDocOptions();

            try
            {
                Word.Document doc = app.ActiveDocument;
                StartUndo(app, "CoolWord: Chèn khung Nơi nhận & Chữ ký NĐ 30/2020");
                try
                {
                    Word.Range endRange = doc.Content.Duplicate;
                    object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                    endRange.Collapse(ref collapseEnd);
                    endRange.InsertParagraphAfter();
                    endRange.Collapse(ref collapseEnd);

                    Word.Table footerTable = doc.Tables.Add(endRange, 1, 2);
                    footerTable.Borders.Enable = 0;
                    footerTable.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                    footerTable.PreferredWidth = 100f;

                    footerTable.Columns[1].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                    footerTable.Columns[1].PreferredWidth = 45f;
                    footerTable.Columns[2].PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
                    footerTable.Columns[2].PreferredWidth = 55f;

                    // Ô (1, 1): Nơi nhận (12pt đậm nghiêng + danh sách 11pt đứng)
                    Word.Cell cellRecipients = footerTable.Cell(1, 1);
                    FormatRecipientsCell(cellRecipients, opts);

                    // Ô (1, 2): Quyền hạn, chức vụ người ký + Khoảng trống ký tên + Họ và tên
                    Word.Cell cellSigner = footerTable.Cell(1, 2);
                    FormatSignerCell(cellSigner, opts);

                    statusMessage = "Đã chèn khung Nơi nhận & Chữ ký chuẩn Nghị định 30/2020/NĐ-CP ở cuối văn bản.";
                    return true;
                }
                finally
                {
                    EndUndo(app);
                }
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi chèn khung Chữ ký: " + ex.Message;
                return false;
            }
        }

        private static void FormatHeaderOrgCell(Word.Cell cell, AdminDocOptions opts)
        {
            Word.Range r = cell.Range;
            r.End -= 1; // Bỏ ký tự kết thúc ô \r\a

            string parent = (opts.ParentOrgName ?? string.Empty).Trim().ToUpperInvariant();
            string issuing = string.IsNullOrWhiteSpace(opts.IssuingOrgName)
                ? "TÊN CƠ QUAN, TỔ CHỨC"
                : opts.IssuingOrgName.Trim().ToUpperInvariant();

            if (!string.IsNullOrEmpty(parent))
            {
                r.Text = parent + "\r" + issuing + "\r──────────";
            }
            else
            {
                r.Text = issuing + "\r──────────";
            }

            r.Font.Name = opts.FontName;
            r.Font.Size = 12.5f;
            r.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            r.ParagraphFormat.FirstLineIndent = 0f;
            r.ParagraphFormat.SpaceBefore = 0f;
            r.ParagraphFormat.SpaceAfter = 0f;
            r.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

            int pCount = cell.Range.Paragraphs.Count;
            if (!string.IsNullOrEmpty(parent) && pCount >= 3)
            {
                cell.Range.Paragraphs[1].Range.Font.Bold = 0;
                cell.Range.Paragraphs[2].Range.Font.Bold = 1;
                cell.Range.Paragraphs[3].Range.Font.Bold = 1;
                cell.Range.Paragraphs[3].Range.Font.Size = 10.0f;
            }
            else if (pCount >= 2)
            {
                cell.Range.Paragraphs[1].Range.Font.Bold = 1;
                cell.Range.Paragraphs[2].Range.Font.Bold = 1;
                cell.Range.Paragraphs[2].Range.Font.Size = 10.0f;
            }
        }

        private static void FormatHeaderQuocHieuCell(Word.Cell cell, AdminDocOptions opts)
        {
            Word.Range r = cell.Range;
            r.End -= 1;

            r.Text = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\rĐộc lập - Tự do - Hạnh phúc\r───────────────────────";
            r.Font.Name = opts.FontName;
            r.Font.Bold = 1;
            r.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            r.ParagraphFormat.FirstLineIndent = 0f;
            r.ParagraphFormat.SpaceBefore = 0f;
            r.ParagraphFormat.SpaceAfter = 0f;
            r.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

            if (cell.Range.Paragraphs.Count >= 3)
            {
                cell.Range.Paragraphs[1].Range.Font.Size = 12.5f;
                cell.Range.Paragraphs[2].Range.Font.Size = 13.5f;
                cell.Range.Paragraphs[3].Range.Font.Size = 10.0f;
            }
        }

        private static void FormatSimpleTableCell(
            Word.Cell cell,
            string text,
            string fontName,
            float fontSize,
            bool bold,
            bool italic,
            float spaceBefore)
        {
            Word.Range r = cell.Range;
            r.End -= 1;
            r.Text = text;
            r.Font.Name = fontName;
            r.Font.Size = fontSize;
            r.Font.Bold = bold ? 1 : 0;
            r.Font.Italic = italic ? 1 : 0;
            r.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            r.ParagraphFormat.FirstLineIndent = 0f;
            r.ParagraphFormat.SpaceBefore = spaceBefore;
            r.ParagraphFormat.SpaceAfter = 2.0f;
            r.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
        }

        private static void FormatRecipientsCell(Word.Cell cell, AdminDocOptions opts)
        {
            Word.Range r = cell.Range;
            r.End -= 1;

            string list = string.IsNullOrWhiteSpace(opts.RecipientsList)
                ? "- Như trên;\r- Lưu: VT."
                : opts.RecipientsList.Trim().Replace("\r\n", "\r").Replace("\n", "\r");

            r.Text = "Nơi nhận:\r" + list;
            r.Font.Name = opts.FontName;
            r.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            r.ParagraphFormat.FirstLineIndent = 0f;
            r.ParagraphFormat.SpaceBefore = 1.0f;
            r.ParagraphFormat.SpaceAfter = 1.0f;
            r.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

            int pCount = cell.Range.Paragraphs.Count;
            if (pCount >= 1)
            {
                Word.Range firstLine = cell.Range.Paragraphs[1].Range;
                firstLine.Font.Size = 12.0f;
                firstLine.Font.Bold = 1;
                firstLine.Font.Italic = 1;
            }

            for (int i = 2; i <= pCount; i++)
            {
                Word.Range itemRange = cell.Range.Paragraphs[i].Range;
                itemRange.Font.Size = 11.0f;
                itemRange.Font.Bold = 0;
                itemRange.Font.Italic = 0;
            }
        }

        private static void FormatSignerCell(Word.Cell cell, AdminDocOptions opts)
        {
            Word.Range r = cell.Range;
            r.End -= 1;

            string auth = (opts.AuthorityPrefix ?? string.Empty).Trim().ToUpperInvariant();
            string pos = string.IsNullOrWhiteSpace(opts.SignerPosition)
                ? "THỦ TRƯỞNG ĐƠN VỊ"
                : opts.SignerPosition.Trim().ToUpperInvariant();
            string name = string.IsNullOrWhiteSpace(opts.SignerFullName)
                ? "Nguyễn Văn A"
                : opts.SignerFullName.Trim();

            string headerBlock = string.IsNullOrEmpty(auth) ? pos : (auth + "\r" + pos);
            r.Text = headerBlock + "\r\r\r\r\r" + name;

            r.Font.Name = opts.FontName;
            r.Font.Size = opts.BodyFontSize;
            r.Font.Bold = 1;
            r.Font.Italic = 0;
            r.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            r.ParagraphFormat.FirstLineIndent = 0f;
            r.ParagraphFormat.SpaceBefore = 0f;
            r.ParagraphFormat.SpaceAfter = 0f;
            r.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
        }

        private static void StartUndo(Word.Application app, string name)
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

        private static void EndUndo(Word.Application app)
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
