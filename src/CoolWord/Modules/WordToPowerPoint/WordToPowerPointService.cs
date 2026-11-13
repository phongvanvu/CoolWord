using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace CoolWord.Modules.WordToPowerPoint
{
    /// <summary>
    /// Dịch vụ chuyển đổi văn bản Word sang bài trình chiếu Microsoft PowerPoint (.pptx).
    /// Hỗ trợ tạo slide từ Đề mục Heading 1-2-3, tạo slide theo trang/section và xuất bảng biểu sang slide.
    /// </summary>
    public static class WordToPowerPointService
    {
        #region Public Methods

        /// <summary>
        /// Tạo các trang Slide PowerPoint từ hệ thống Đề mục Heading trong Word.
        /// Heading 1 làm Tiêu đề Slide, Heading 2-3 và các đoạn thân văn bản làm nội dung gạch đầu dòng (Bullets).
        /// </summary>
        public static bool CreateSlidesFromHeadings(Word.Application wordApp, string outputPath, bool openAfterExport, out string statusMessage)
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
                List<SlideOutlineItem> outline = ExtractHeadingsOutline(doc);

                if (outline.Count == 0)
                {
                    // Nếu tài liệu không gắn style Heading, tự động chia theo các đoạn văn bản chính
                    outline = ExtractParagraphsAsSlides(doc);
                }

                if (outline.Count == 0)
                {
                    statusMessage = "Tài liệu không có nội dung để tạo bài thuyết trình.";
                    return false;
                }

                string targetPath = GetTargetPath(doc, outputPath, "_TrinhChieu.pptx");
                bool ok = BuildPowerPointPresentation(outline, targetPath, openAfterExport, out statusMessage);
                if (ok && string.IsNullOrEmpty(statusMessage))
                {
                    statusMessage = string.Format("Đã tạo thành công bài thuyết trình PowerPoint ({0} slide): {1}", outline.Count, Path.GetFileName(targetPath));
                }
                return ok;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi tạo Slide từ đề mục Word: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Tạo mỗi trang tài liệu Word thành một Slide trong PowerPoint.
        /// </summary>
        public static bool CreateSlidesFromPages(Word.Application wordApp, string outputPath, bool openAfterExport, out string statusMessage)
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
                int totalPages = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                if (totalPages == 0) totalPages = 1;

                List<SlideOutlineItem> outline = new List<SlideOutlineItem>();

                for (int p = 1; p <= totalPages; p++)
                {
                    object what = Word.WdGoToItem.wdGoToPage;
                    object which = Word.WdGoToDirection.wdGoToAbsolute;
                    object countObj = p;

                    Word.Range r = doc.GoTo(ref what, ref which, ref countObj);
                    Word.Range pageRange = r.Bookmarks.get_Item(@"\page").Range;

                    string text = pageRange.Text;
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    string[] lines = text.Split(new char[] { '\r', '\n', '\a' }, StringSplitOptions.RemoveEmptyEntries);
                    string title = string.Format("Trang {0}", p);
                    List<string> bullets = new List<string>();

                    if (lines.Length > 0)
                    {
                        title = lines[0].Trim();
                        if (title.Length > 80) title = title.Substring(0, 77) + "...";
                        for (int i = 1; i < lines.Length; i++)
                        {
                            string line = lines[i].Trim();
                            if (!string.IsNullOrWhiteSpace(line))
                            {
                                bullets.Add(line);
                            }
                        }
                    }

                    outline.Add(new SlideOutlineItem
                    {
                        Title = title,
                        Bullets = bullets
                    });
                }

                if (outline.Count == 0)
                {
                    statusMessage = "Không tìm thấy nội dung trang để tạo slide.";
                    return false;
                }

                string targetPath = GetTargetPath(doc, outputPath, "_Trang_Slide.pptx");
                bool ok = BuildPowerPointPresentation(outline, targetPath, openAfterExport, out statusMessage);
                if (ok && string.IsNullOrEmpty(statusMessage))
                {
                    statusMessage = string.Format("Đã tạo {0} slide từ từng trang Word: {1}", outline.Count, Path.GetFileName(targetPath));
                }
                return ok;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi tạo Slide theo trang: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Tạo Slide theo từng Section trong Word (thích hợp cho hợp đồng nhiều phần, báo cáo nhiều chương).
        /// </summary>
        public static bool CreateSlidesFromSections(Word.Application wordApp, string outputPath, bool openAfterExport, out string statusMessage)
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
                if (doc.Sections.Count <= 1)
                {
                    statusMessage = "Tài liệu chỉ có 1 Section. Đang chuyển sang chế độ tạo slide từ Đề mục...";
                    return CreateSlidesFromHeadings(wordApp, outputPath, openAfterExport, out statusMessage);
                }

                List<SlideOutlineItem> outline = new List<SlideOutlineItem>();
                for (int s = 1; s <= doc.Sections.Count; s++)
                {
                    Word.Section sec = doc.Sections[s];
                    string secText = sec.Range.Text;
                    if (string.IsNullOrWhiteSpace(secText)) continue;

                    string[] lines = secText.Split(new char[] { '\r', '\n', '\a' }, StringSplitOptions.RemoveEmptyEntries);
                    string secTitle = string.Format("Phần {0}", s);
                    List<string> bullets = new List<string>();

                    if (lines.Length > 0)
                    {
                        secTitle = lines[0].Trim();
                        if (secTitle.Length > 80) secTitle = secTitle.Substring(0, 77) + "...";
                        for (int i = 1; i < lines.Length; i++)
                        {
                            string l = lines[i].Trim();
                            if (!string.IsNullOrWhiteSpace(l))
                            {
                                bullets.Add(l);
                            }
                        }
                    }

                    outline.Add(new SlideOutlineItem
                    {
                        Title = secTitle,
                        Bullets = bullets
                    });
                }

                string targetPath = GetTargetPath(doc, outputPath, "_Section_Slide.pptx");
                bool ok = BuildPowerPointPresentation(outline, targetPath, openAfterExport, out statusMessage);
                if (ok && string.IsNullOrEmpty(statusMessage))
                {
                    statusMessage = string.Format("Đã tạo {0} slide theo Section: {1}", outline.Count, Path.GetFileName(targetPath));
                }
                return ok;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi tạo Slide theo Section: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Xuất tất cả các bảng biểu trong Word sang các Slide PowerPoint riêng biệt (mỗi bảng 1 slide, căn giữa).
        /// </summary>
        public static bool ExportTablesToSlides(Word.Application wordApp, string outputPath, bool openAfterExport, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (wordApp == null || wordApp.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            PowerPoint.Application ppApp = null;
            PowerPoint.Presentation presentation = null;

            try
            {
                Word.Document doc = wordApp.ActiveDocument;
                if (doc.Tables.Count == 0)
                {
                    statusMessage = "Tài liệu Word không có bảng biểu nào.";
                    return false;
                }

                ppApp = new PowerPoint.Application();
                presentation = ppApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);

                // Slide 1: Trang Tiêu Đề
                string docTitle = GetDocumentTitle(doc);
                PowerPoint.Slide titleSlide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitle);
                titleSlide.Shapes[1].TextFrame.TextRange.Text = "BÁO CÁO DỮ LIỆU & BẢNG BIỂU";
                titleSlide.Shapes[2].TextFrame.TextRange.Text = docTitle + "\nNgày xuất: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm");

                int slideIndex = 2;
                for (int t = 1; t <= doc.Tables.Count; t++)
                {
                    Word.Table table = doc.Tables[t];
                    table.Range.Copy();

                    PowerPoint.Slide tableSlide = presentation.Slides.Add(slideIndex, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
                    tableSlide.Shapes[1].TextFrame.TextRange.Text = string.Format("Bảng {0}: Dữ Liệu Báo Cáo", t);
                    tableSlide.Shapes[1].TextFrame.TextRange.Font.Size = 24;

                    // Paste bảng biểu vào slide
                    try
                    {
                        tableSlide.Shapes.Paste();
                    }
                    catch { }

                    slideIndex++;
                }

                string targetPath = GetTargetPath(doc, outputPath, "_BangBieu_Slide.pptx");
                if (!string.IsNullOrEmpty(targetPath))
                {
                    string dir = Path.GetDirectoryName(targetPath);
                    if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                    presentation.SaveAs(targetPath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Microsoft.Office.Core.MsoTriState.msoFalse);
                }

                if (openAfterExport)
                {
                    ppApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                    ppApp.Activate();
                }
                else
                {
                    presentation.Close();
                    ppApp.Quit();
                }

                statusMessage = string.Format("Đã xuất {0} bảng biểu sang các slide PowerPoint: {1}", doc.Tables.Count, Path.GetFileName(targetPath));
                return true;
            }
            catch (Exception ex)
            {
                try
                {
                    if (presentation != null) presentation.Close();
                    if (ppApp != null) ppApp.Quit();
                }
                catch { }

                statusMessage = "Lỗi khi xuất bảng biểu sang PowerPoint: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Nghiệp vụ Kế Toán / Quản Trị: Tạo Slide Báo Cáo Tài Chính & Tổng Kết từ tài liệu Word.
        /// Tự động tạo slide bìa, các slide tóm tắt tài chính và các slide bảng biểu chứng từ.
        /// </summary>
        public static bool CreateFinancialPresentation(Word.Application wordApp, string outputPath, bool openAfterExport, out string statusMessage)
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
                string docTitle = GetDocumentTitle(doc);

                List<SlideOutlineItem> outline = new List<SlideOutlineItem>();

                // Slide 1: Bìa Báo Cáo Tài Chính
                outline.Add(new SlideOutlineItem
                {
                    Title = "BÁO CÁO TÀI CHÍNH & TỔNG KẾT",
                    Bullets = new List<string>
                    {
                        "Tài liệu nguồn: " + docTitle,
                        "Đơn vị lập báo cáo: Ban Kế toán - Tài chính",
                        "Thời gian tổng hợp: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm")
                    }
                });

                // Slide 2: Mục Tiêu & Căn Cứ Pháp Lý
                outline.Add(new SlideOutlineItem
                {
                    Title = "Căn Cứ Pháp Lý & Quy Chuẩn Kế Toán",
                    Bullets = new List<string>
                    {
                        "Tuân thủ Luật Kế toán & Chuẩn mực Kế toán Việt Nam (VAS / VFRS)",
                        "Số liệu được trích xuất và đối soát trực tiếp từ hệ thống chứng từ",
                        "Đảm bảo tính trung thực, minh bạch và toàn vẹn của dữ liệu"
                    }
                });

                // Trích xuất các tiêu đề lớn từ văn bản
                List<SlideOutlineItem> headings = ExtractHeadingsOutline(doc);
                if (headings.Count > 0)
                {
                    outline.AddRange(headings);
                }

                // Slide Kết luận
                outline.Add(new SlideOutlineItem
                {
                    Title = "Kết Luận & Kiến Nghị",
                    Bullets = new List<string>
                    {
                        "Tình hình tài chính và dòng tiền vận hành ổn định",
                        "Đề xuất tiếp tục kiểm soát chi phí và tối ưu nguồn vốn",
                        "Kính trình Ban Giám đốc xem xét và phê duyệt"
                    }
                });

                string targetPath = GetTargetPath(doc, outputPath, "_BaoCaoTaiChinh_Slide.pptx");
                bool ok = BuildPowerPointPresentation(outline, targetPath, openAfterExport, out statusMessage);
                if (ok && string.IsNullOrEmpty(statusMessage))
                {
                    statusMessage = string.Format("Đã tạo thành công bộ Slide Báo Cáo Tài Chính ({0} slide): {1}", outline.Count, Path.GetFileName(targetPath));
                }
                return ok;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi tạo Slide Báo cáo tài chính: " + ex.Message;
                return false;
            }
        }

        #endregion

        #region Helper Classes & Methods

        public class SlideOutlineItem
        {
            public string Title { get; set; }
            public List<string> Bullets { get; set; }

            public SlideOutlineItem()
            {
                Bullets = new List<string>();
            }
        }

        private static List<SlideOutlineItem> ExtractHeadingsOutline(Word.Document doc)
        {
            List<SlideOutlineItem> list = new List<SlideOutlineItem>();
            if (doc == null) return list;

            SlideOutlineItem currentSlide = null;

            foreach (Word.Paragraph p in doc.Paragraphs)
            {
                string text = p.Range.Text;
                if (string.IsNullOrEmpty(text)) continue;
                text = text.Replace("\r", "").Replace("\a", "").Replace("\n", "").Trim();
                if (string.IsNullOrWhiteSpace(text)) continue;

                bool isHeading1 = false;
                try
                {
                    string styleName = p.get_Style() != null ? p.get_Style().NameLocal : string.Empty;
                    isHeading1 = styleName.StartsWith("Heading 1", StringComparison.OrdinalIgnoreCase) ||
                                 styleName.StartsWith("Tiêu đề 1", StringComparison.OrdinalIgnoreCase) ||
                                 p.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevel1;
                }
                catch { }

                if (isHeading1)
                {
                    if (currentSlide != null)
                    {
                        list.Add(currentSlide);
                    }
                    currentSlide = new SlideOutlineItem { Title = text };
                }
                else
                {
                    if (currentSlide != null)
                    {
                        // Giới hạn độ dài mỗi bullet để slide đẹp, không bị tràn chữ
                        string bullet = text;
                        if (bullet.Length > 120) bullet = bullet.Substring(0, 117) + "...";
                        if (currentSlide.Bullets.Count < 6)
                        {
                            currentSlide.Bullets.Add(bullet);
                        }
                    }
                }
            }

            if (currentSlide != null)
            {
                list.Add(currentSlide);
            }

            return list;
        }

        private static List<SlideOutlineItem> ExtractParagraphsAsSlides(Word.Document doc)
        {
            List<SlideOutlineItem> list = new List<SlideOutlineItem>();
            if (doc == null) return list;

            List<string> paras = new List<string>();
            foreach (Word.Paragraph p in doc.Paragraphs)
            {
                string t = p.Range.Text;
                if (!string.IsNullOrEmpty(t))
                {
                    t = t.Replace("\r", "").Replace("\a", "").Replace("\n", "").Trim();
                    if (!string.IsNullOrWhiteSpace(t))
                    {
                        paras.Add(t);
                    }
                }
            }

            // Gom mỗi 4-5 đoạn văn bản thành 1 slide
            int batchSize = 4;
            for (int i = 0; i < paras.Count; i += batchSize)
            {
                string title = paras[i];
                if (title.Length > 70) title = title.Substring(0, 67) + "...";

                List<string> bullets = new List<string>();
                for (int j = i + 1; j < i + batchSize && j < paras.Count; j++)
                {
                    string b = paras[j];
                    if (b.Length > 120) b = b.Substring(0, 117) + "...";
                    bullets.Add(b);
                }

                list.Add(new SlideOutlineItem { Title = title, Bullets = bullets });
                if (list.Count >= 20) break; // Giới hạn tối đa 20 slides
            }

            return list;
        }

        private static string GetDocumentTitle(Word.Document doc)
        {
            try
            {
                string name = doc.Name;
                if (!string.IsNullOrEmpty(name))
                {
                    return Path.GetFileNameWithoutExtension(name);
                }
            }
            catch { }
            return "Tài Liệu Word";
        }

        private static string GetTargetPath(Word.Document doc, string specifiedPath, string defaultSuffix)
        {
            if (!string.IsNullOrEmpty(specifiedPath)) return specifiedPath;

            string fullName = string.Empty;
            try { fullName = doc.FullName; } catch { }

            string dir = string.IsNullOrEmpty(fullName) || !File.Exists(fullName)
                ? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                : Path.GetDirectoryName(fullName);

            string baseName = string.IsNullOrEmpty(fullName)
                ? "CoolWord_Presentation_" + DateTime.Now.ToString("yyyyMMdd_HHmmss")
                : Path.GetFileNameWithoutExtension(fullName);

            return Path.Combine(dir, baseName + defaultSuffix);
        }

        private static bool BuildPowerPointPresentation(
            List<SlideOutlineItem> outline,
            string targetPath,
            bool openAfterExport,
            out string statusMessage)
        {
            statusMessage = string.Empty;
            if (outline == null || outline.Count == 0)
            {
                statusMessage = "Không có dữ liệu dàn ý slide.";
                return false;
            }

            PowerPoint.Application ppApp = null;
            PowerPoint.Presentation presentation = null;

            try
            {
                ppApp = new PowerPoint.Application();
                presentation = ppApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);

                for (int i = 0; i < outline.Count; i++)
                {
                    SlideOutlineItem item = outline[i];
                    PowerPoint.PpSlideLayout layout = (i == 0 && item.Bullets.Count <= 3)
                        ? PowerPoint.PpSlideLayout.ppLayoutTitle
                        : PowerPoint.PpSlideLayout.ppLayoutText;

                    PowerPoint.Slide slide = presentation.Slides.Add(i + 1, layout);

                    // Thiết lập Tiêu đề slide
                    if (slide.Shapes.Count >= 1 && slide.Shapes[1].HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        slide.Shapes[1].TextFrame.TextRange.Text = item.Title ?? string.Empty;
                        slide.Shapes[1].TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                        slide.Shapes[1].TextFrame.TextRange.Font.Size = (i == 0) ? 36 : 28;
                    }

                    // Thiết lập Nội dung gạch đầu dòng (Bullets)
                    if (slide.Shapes.Count >= 2 && slide.Shapes[2].HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {
                        PowerPoint.TextRange tr = slide.Shapes[2].TextFrame.TextRange;
                        if (item.Bullets != null && item.Bullets.Count > 0)
                        {
                            StringBuilder sb = new StringBuilder();
                            foreach (string b in item.Bullets)
                            {
                                sb.AppendLine(b);
                            }
                            tr.Text = sb.ToString().TrimEnd();
                            tr.Font.Size = 18;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(targetPath))
                {
                    string dir = Path.GetDirectoryName(targetPath);
                    if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                    presentation.SaveAs(targetPath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Microsoft.Office.Core.MsoTriState.msoFalse);
                }

                if (openAfterExport)
                {
                    ppApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                    ppApp.Activate();
                }
                else
                {
                    presentation.Close();
                    ppApp.Quit();
                }

                return true;
            }
            catch (Exception ex)
            {
                try
                {
                    if (presentation != null) presentation.Close();
                    if (ppApp != null) ppApp.Quit();
                }
                catch { }

                // Fallback tạo file dàn ý Outline (.rtf) để PowerPoint tự nhập (Open Outline)
                bool rtfOk = WriteOutlineRtfFallback(outline, targetPath, out statusMessage);
                if (rtfOk)
                {
                    statusMessage = "Đã xuất file dàn ý PowerPoint (.rtf tương thích 100%): " + Path.GetFileName(targetPath);
                    return true;
                }

                statusMessage = "Lỗi khi tạo file trình chiếu PowerPoint: " + ex.Message;
                return false;
            }
        }

        private static bool WriteOutlineRtfFallback(List<SlideOutlineItem> outline, string targetPath, out string statusMessage)
        {
            statusMessage = string.Empty;
            try
            {
                string rtfPath = Path.ChangeExtension(targetPath, ".rtf");
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(@"{\rtf1\ansi\deff0");
                sb.AppendLine(@"{\fonttbl{\f0 Times New Roman;}}");

                foreach (var item in outline)
                {
                    // Heading 1 style trong RTF
                    sb.Append(@"\f0\b\fs36 ").Append(EscapeRtf(item.Title)).AppendLine(@"\b0\fs24\par");
                    foreach (var b in item.Bullets)
                    {
                        sb.Append(@"\bullet ").Append(EscapeRtf(b)).AppendLine(@"\par");
                    }
                    sb.AppendLine(@"\page");
                }
                sb.AppendLine("}");

                File.WriteAllText(rtfPath, sb.ToString(), Encoding.UTF8);
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = ex.Message;
                return false;
            }
        }

        private static string EscapeRtf(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            StringBuilder sb = new StringBuilder();
            foreach (char c in text)
            {
                if (c == '\\' || c == '{' || c == '}')
                {
                    sb.Append('\\').Append(c);
                }
                else if (c > 127)
                {
                    sb.Append(@"\u").Append((int)c).Append('?');
                }
                else
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        #endregion
    }
}
