using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace CoolWord.Modules.TableOfContents
{
    public static class TableOfContentsService
    {
        // Regex nhận diện tiêu đề cấp 1 (Heading 1)
        private static readonly Regex RegexHeading1 = new Regex(
            @"^\s*(CHƯƠNG|Chương|PHẦN|Phần|ĐIỀU|Điều)\s+[0-9IVXLCDMivxlcdm]+[\.\:\s\-].*|" +
            @"^\s*[I|V|X|L|C|D|M]+\.\s+\S+.*|" +
            @"^\s*(BẢNG CÂN ĐỐI KẾ TOÁN|BÁO CÁO KẾT QUẢ HOẠT ĐỘNG KINH DOANH|BÁO CÁO LƯU CHUYỂN TIỀN TỆ|THUYẾT MINH BÁO CÁO TÀI CHÍNH|DANH MỤC CHỨNG TỪ KẾ TOÁN|KẾT QUẢ KINH DOANH).*",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        // Regex nhận diện tiêu đề cấp 2 (Heading 2)
        private static readonly Regex RegexHeading2 = new Regex(
            @"^\s*\d+\.\s+[^\d\s].*|" +
            @"^\s*\d+\.\d+\s+\S+.*|" +
            @"^\s*(MỤC|Mục|KHOẢN|Khoản)\s+[0-9IVXLCDMivxlcdm]+[\.\:\s\-].*",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        // Regex nhận diện tiêu đề cấp 3 (Heading 3)
        private static readonly Regex RegexHeading3 = new Regex(
            @"^\s*\d+\.\d+\.\d+\s+\S+.*|" +
            @"^\s*[a-z]\)\s+\S+.*|" +
            @"^\s*(Điểm|Tiểu mục)\s+[a-z0-9]+[\.\:\s\-].*",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        /// <summary>
        /// Phân tích đoạn văn xem có phù hợp với cấp độ Heading nào không (1, 2, 3 hoặc 0 nếu không phải).
        /// </summary>
        public static int DetectHeadingLevel(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;
            string trimmed = text.Trim();

            // Nếu đoạn văn quá dài (> 200 ký tự) hoặc kết thúc bằng dấu chấm câu dài, nhiều khả năng là đoạn văn nội dung
            if (trimmed.Length > 200) return 0;

            if (RegexHeading1.IsMatch(trimmed)) return 1;
            if (RegexHeading2.IsMatch(trimmed)) return 2;
            if (RegexHeading3.IsMatch(trimmed)) return 3;

            return 0;
        }

        /// <summary>
        /// Chèn bảng mục lục tự động (TOC) vào tài liệu Word.
        /// </summary>
        public static bool InsertTableOfContents(
            Word.Document doc,
            int upperLevel,
            bool atStartOfDoc,
            bool includeHyperlinks,
            string leaderChar,
            out string statusMessage)
        {
            if (doc == null)
            {
                statusMessage = "Không có tài liệu Word nào đang mở.";
                return false;
            }

            try
            {
                int maxLevel = Math.Max(1, Math.Min(9, upperLevel));
                Word.Range targetRange;

                if (atStartOfDoc)
                {
                    targetRange = doc.Range(0, 0);
                    // Chèn dòng tiêu đề mục lục
                    targetRange.Text = "MỤC LỤC\r\n";
                    targetRange.Font.Name = "Times New Roman";
                    targetRange.Font.Size = 14;
                    targetRange.Font.Bold = 1;
                    targetRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    // Di chuyển con trỏ xuống dưới dòng tiêu đề
                    targetRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }
                else
                {
                    targetRange = doc.Application.Selection.Range;
                }

                object useHeadingStyles = true;
                object upperHeadingLevel = 1;
                object lowerHeadingLevel = maxLevel;
                object useFields = true;
                object tableID = 1;
                object rightAlignPageNumbers = true;
                object includePageNumbers = true;
                object useHyperlinks = includeHyperlinks;
                object hidePageNumbersInWeb = true;
                object useOutlineLevels = true;

                Word.TableOfContents toc = doc.TablesOfContents.Add(
                    targetRange,
                    ref useHeadingStyles,
                    ref upperHeadingLevel,
                    ref lowerHeadingLevel,
                    ref useFields,
                    ref tableID,
                    ref rightAlignPageNumbers,
                    ref includePageNumbers,
                    Type.Missing,
                    ref useHyperlinks,
                    ref hidePageNumbersInWeb,
                    ref useOutlineLevels);

                // Thiết lập kiểu dấu chấm dẫn dòng
                if (leaderChar == "hyphens")
                    toc.TabLeader = Word.WdTabLeader.wdTabLeaderDashes;
                else if (leaderChar == "lines")
                    toc.TabLeader = Word.WdTabLeader.wdTabLeaderLines;
                else if (leaderChar == "none")
                    toc.TabLeader = Word.WdTabLeader.wdTabLeaderSpaces;
                else
                    toc.TabLeader = Word.WdTabLeader.wdTabLeaderDots;

                int detectedHeadings = CountHeadings(doc, maxLevel);
                if (detectedHeadings == 0)
                {
                    statusMessage = string.Format(
                        "Đã chèn khung Mục Lục (1 - {0} cấp). Lưu ý: Tài liệu chưa có đề mục Heading nào. Hãy dùng nút 'Quét & Gán Heading Tự Động' để nạp nội dung.",
                        maxLevel);
                }
                else
                {
                    statusMessage = string.Format(
                        "Đã tạo Mục Lục tự động thành công ({0} cấp đề mục, phát hiện {1} đề mục).",
                        maxLevel, detectedHeadings);
                }
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi chèn mục lục: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Chèn Mục lục chuẩn cho Báo cáo tài chính & Kế toán doanh nghiệp.
        /// </summary>
        public static bool InsertAccountingFinancialToc(
            Word.Document doc,
            bool atStartOfDoc,
            out string statusMessage)
        {
            if (doc == null)
            {
                statusMessage = "Không có tài liệu Word nào đang mở.";
                return false;
            }

            try
            {
                // Trước tiên tự động quét gán các tiêu đề BCTC nếu có
                int assigned;
                string assignMsg;
                AutoAssignHeadings(doc, out assigned, out assignMsg);

                Word.Range targetRange;
                if (atStartOfDoc)
                {
                    targetRange = doc.Range(0, 0);
                    targetRange.Text = "MỤC LỤC BÁO CÁO TÀI CHÍNH & CHỨNG TỪ KẾ TOÁN\r\n";
                    targetRange.Font.Name = "Times New Roman";
                    targetRange.Font.Size = 14;
                    targetRange.Font.Bold = 1;
                    targetRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    targetRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }
                else
                {
                    targetRange = doc.Application.Selection.Range;
                }

                object useHeadingStyles = true;
                object upperHeadingLevel = 1;
                object lowerHeadingLevel = 3;
                object useFields = true;
                object tableID = 1;
                object rightAlignPageNumbers = true;
                object includePageNumbers = true;
                object useHyperlinks = true;
                object hidePageNumbersInWeb = true;
                object useOutlineLevels = true;

                Word.TableOfContents toc = doc.TablesOfContents.Add(
                    targetRange,
                    ref useHeadingStyles,
                    ref upperHeadingLevel,
                    ref lowerHeadingLevel,
                    ref useFields,
                    ref tableID,
                    ref rightAlignPageNumbers,
                    ref includePageNumbers,
                    Type.Missing,
                    ref useHyperlinks,
                    ref hidePageNumbersInWeb,
                    ref useOutlineLevels);

                toc.TabLeader = Word.WdTabLeader.wdTabLeaderDots;

                // Chuẩn hóa định dạng phông chữ Times New Roman theo thể thức NĐ 30/2020
                string formatMsg;
                FormatTableOfContentsDecree30(doc, out formatMsg);

                statusMessage = string.Format(
                    "Đã tạo Mục Lục Báo Cáo Tài Chính & Kế Toán thành công (Tự động nhận diện và gán {0} đề mục).",
                    assigned);
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi tạo Mục lục Báo cáo tài chính: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Quét toàn bộ tài liệu và tự động gán Style Heading 1, 2, 3 cho các dòng tiêu đề văn bản.
        /// </summary>
        public static bool AutoAssignHeadings(
            Word.Document doc,
            out int assignedCount,
            out string statusMessage)
        {
            assignedCount = 0;
            if (doc == null)
            {
                statusMessage = "Không có tài liệu Word nào đang mở.";
                return false;
            }

            try
            {
                int countH1 = 0;
                int countH2 = 0;
                int countH3 = 0;

                object styleH1 = Word.WdBuiltinStyle.wdStyleHeading1;
                object styleH2 = Word.WdBuiltinStyle.wdStyleHeading2;
                object styleH3 = Word.WdBuiltinStyle.wdStyleHeading3;

                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    try
                    {
                        string text = para.Range.Text;
                        if (string.IsNullOrWhiteSpace(text)) continue;

                        int level = DetectHeadingLevel(text);
                        if (level == 1)
                        {
                            para.set_Style(ref styleH1);
                            countH1++;
                        }
                        else if (level == 2)
                        {
                            para.set_Style(ref styleH2);
                            countH2++;
                        }
                        else if (level == 3)
                        {
                            para.set_Style(ref styleH3);
                            countH3++;
                        }
                    }
                    catch { }
                }

                assignedCount = countH1 + countH2 + countH3;

                // Tự động làm mới các bảng mục lục nếu có
                if (doc.TablesOfContents.Count > 0)
                {
                    foreach (Word.TableOfContents toc in doc.TablesOfContents)
                    {
                        try { toc.Update(); } catch { }
                    }
                }

                if (assignedCount > 0)
                {
                    statusMessage = string.Format(
                        "Đã tự động gán {0} đề mục Heading (Heading 1: {1}, Heading 2: {2}, Heading 3: {3}).",
                        assignedCount, countH1, countH2, countH3);
                }
                else
                {
                    statusMessage = "Không tìm thấy tiêu đề phù hợp (Chương, Phần, Điều, 1.1, 1.2...) để gán Heading.";
                }
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi quét gán đề mục Heading: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Cập nhật tất cả các bảng mục lục hiện có trong tài liệu.
        /// </summary>
        public static bool UpdateAllTablesOfContents(
            Word.Document doc,
            bool entireTable,
            out string statusMessage)
        {
            if (doc == null)
            {
                statusMessage = "Không có tài liệu Word nào đang mở.";
                return false;
            }

            try
            {
                int count = doc.TablesOfContents.Count;
                if (count == 0)
                {
                    statusMessage = "Tài liệu chưa có bảng mục lục nào để cập nhật.";
                    return false;
                }

                foreach (Word.TableOfContents toc in doc.TablesOfContents)
                {
                    try
                    {
                        if (entireTable)
                            toc.Update();
                        else
                            toc.UpdatePageNumbers();
                    }
                    catch { }
                }

                // Cập nhật cả danh mục bảng biểu và hình ảnh nếu có
                try
                {
                    foreach (Word.TableOfFigures tof in doc.TablesOfFigures)
                    {
                        try
                        {
                            if (entireTable) tof.Update();
                            else tof.UpdatePageNumbers();
                        }
                        catch { }
                    }
                }
                catch { }

                statusMessage = string.Format(
                    "Đã cập nhật thành công {0} bảng mục lục ({1}).",
                    count, entireTable ? "Cả nội dung & số trang" : "Chỉ số trang");
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi cập nhật mục lục: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Xóa bỏ tất cả các bảng mục lục hiện có trong tài liệu.
        /// </summary>
        public static bool DeleteAllTablesOfContents(Word.Document doc, out string statusMessage)
        {
            if (doc == null)
            {
                statusMessage = "Không có tài liệu Word nào đang mở.";
                return false;
            }

            try
            {
                int count = doc.TablesOfContents.Count;
                if (count == 0)
                {
                    statusMessage = "Tài liệu hiện không có bảng mục lục nào.";
                    return false;
                }

                while (doc.TablesOfContents.Count > 0)
                {
                    try
                    {
                        doc.TablesOfContents[1].Delete();
                    }
                    catch
                    {
                        break;
                    }
                }

                statusMessage = string.Format("Đã xóa bỏ thành công {0} bảng mục lục khỏi tài liệu.", count);
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi xóa mục lục: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Chèn danh mục Bảng Biểu (List of Tables) hoặc Hình Ảnh (List of Figures).
        /// </summary>
        public static bool InsertTableOfFigures(
            Word.Document doc,
            bool isTable,
            bool atStartOfDoc,
            out string statusMessage)
        {
            if (doc == null)
            {
                statusMessage = "Không có tài liệu Word nào đang mở.";
                return false;
            }

            try
            {
                Word.Range targetRange;
                string title = isTable ? "DANH MỤC BẢNG BIỂU\r\n" : "DANH MỤC HÌNH ẢNH / SƠ ĐỒ\r\n";

                if (atStartOfDoc)
                {
                    targetRange = doc.Range(0, 0);
                    targetRange.Text = title;
                    targetRange.Font.Name = "Times New Roman";
                    targetRange.Font.Size = 14;
                    targetRange.Font.Bold = 1;
                    targetRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    targetRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }
                else
                {
                    targetRange = doc.Application.Selection.Range;
                }

                string captionLabel = isTable ? "Bảng" : "Hình";
                object caption = captionLabel;
                object includeLabel = true;
                object useHeadingStyles = false;
                object rightAlignPageNumbers = true;
                object includePageNumbers = true;
                object useHyperlinks = true;
                object hidePageNumbersInWeb = true;

                Word.TableOfFigures tof = doc.TablesOfFigures.Add(
                    targetRange,
                    ref caption,
                    ref includeLabel,
                    ref useHeadingStyles,
                    Type.Missing,
                    Type.Missing,
                    ref rightAlignPageNumbers,
                    ref includePageNumbers,
                    Type.Missing,
                    ref useHyperlinks,
                    ref hidePageNumbersInWeb);

                tof.TabLeader = Word.WdTabLeader.wdTabLeaderDots;

                statusMessage = string.Format(
                    "Đã chèn {0} thành công (Nhãn chú thích: \"{1}\").",
                    isTable ? "Danh mục Bảng Biểu" : "Danh mục Hình Ảnh",
                    captionLabel);
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi chèn danh mục bảng biểu/hình ảnh: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Định dạng lại toàn bộ mục lục theo chuẩn thể thức văn bản hành chính NĐ 30/2020.
        /// </summary>
        public static bool FormatTableOfContentsDecree30(Word.Document doc, out string statusMessage)
        {
            if (doc == null)
            {
                statusMessage = "Không có tài liệu Word nào đang mở.";
                return false;
            }

            try
            {
                if (doc.TablesOfContents.Count == 0)
                {
                    statusMessage = "Tài liệu chưa có mục lục nào để định dạng.";
                    return false;
                }

                foreach (Word.TableOfContents toc in doc.TablesOfContents)
                {
                    try
                    {
                        Word.Range r = toc.Range;
                        r.Font.Name = "Times New Roman";
                        r.Font.Size = 13;
                        r.Font.ColorIndex = Word.WdColorIndex.wdAuto;
                        r.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                        r.ParagraphFormat.LineSpacing = doc.Application.LinesToPoints(1.2f);
                        r.ParagraphFormat.SpaceBefore = 2f;
                        r.ParagraphFormat.SpaceAfter = 2f;
                    }
                    catch { }
                }

                statusMessage = "Đã chuẩn hóa định dạng toàn bộ mục lục theo chuẩn Nghị định 30/2020 (Times New Roman 13pt, giãn dòng 1.2).";
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi định dạng mục lục: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Đếm số lượng đề mục Heading trong văn bản.
        /// </summary>
        public static int CountHeadings(Word.Document doc, int maxLevel)
        {
            if (doc == null) return 0;
            int count = 0;
            try
            {
                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    try
                    {
                        object styleObj = para.get_Style();
                        string styleName = styleObj != null ? styleObj.ToString() : string.Empty;

                        for (int lvl = 1; lvl <= maxLevel; lvl++)
                        {
                            if (styleName.IndexOf("Heading " + lvl, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                styleName.IndexOf("Tiêu đề " + lvl, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                count++;
                                break;
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return count;
        }

        /// <summary>
        /// Tạo danh sách các đề mục mẫu chuẩn cho Thuyết minh Báo cáo tài chính kế toán.
        /// </summary>
        public static List<string> GetAccountingFinancialHeadingsTemplate()
        {
            return new List<string>
            {
                "PHẦN I: BÁO CÁO TÀI CHÍNH NĂM",
                "1. Bảng cân đối kế toán (Mẫu số B01-DN)",
                "2. Báo cáo kết quả hoạt động kinh doanh (Mẫu số B02-DN)",
                "3. Báo cáo lưu chuyển tiền tệ (Mẫu số B03-DN)",
                "4. Thuyết minh báo cáo tài chính (Mẫu số B09-DN)",
                "PHẦN II: HỒ SƠ & DANH MỤC CHỨNG TỪ KẾ TOÁN",
                "1. Bảng tổng hợp phiếu thu - phiếu chi quỹ tiền mặt",
                "2. Bảng kê hóa đơn chứng từ hàng hóa dịch vụ mua vào & bán ra",
                "3. Sổ nhật ký chung và Sổ cái các tài khoản tài chính",
                "4. Bảng tính và phân bổ khấu hao tài sản cố định (TSCĐ)",
                "5. Bảng thanh toán tiền lương và các khoản trích theo lương (BHXH, BHYT, BHTN)",
                "6. Biên bản đối chiếu công nợ phải thu - phải trả khách hàng",
                "7. Biên bản kiểm kê quỹ tiền mặt và hàng tồn kho cuối kỳ",
                "PHẦN III: PHỤ LỤC & GIẢI TRÌNH KIỂM TOÁN",
                "1. Chính sách kế toán áp dụng tại doanh nghiệp",
                "2. Sự kiện phát sinh sau ngày kết thúc kỳ kế toán năm"
            };
        }
    }
}
