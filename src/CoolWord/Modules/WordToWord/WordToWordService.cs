using System;
using System.IO;
using System.Windows.Forms;
using CoolWord.Core;
using Word = Microsoft.Office.Interop.Word;

namespace CoolWord.Modules.WordToWord
{
    public static class WordToWordService
    {
        private static object Missing = Type.Missing;

        /// <summary>
        /// Chuyển đổi mã phông cũ (TCVN3 / VNI) sang Unicode chuẩn và đổi phông chữ sang Times New Roman.
        /// </summary>
        public static bool ConvertEncodingToUnicode(Word.Application app, bool selectionOnly, string sourceEncoding, out string statusMessage)
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
                Word.Range range = selectionOnly ? app.Selection.Range : doc.Content;

                if (range == null || range.Start == range.End)
                {
                    range = doc.Content;
                }

                string raw = range.Text;
                if (string.IsNullOrEmpty(raw))
                {
                    statusMessage = "Nội dung văn bản trống.";
                    return false;
                }

                string mode = string.IsNullOrEmpty(sourceEncoding) ? "AUTO" : sourceEncoding.ToUpperInvariant();
                string converted = raw;

                if (mode == "TCVN3")
                {
                    converted = VietnameseEncodingConverter.ConvertTcvn3ToUnicode(raw);
                }
                else if (mode == "VNI")
                {
                    converted = VietnameseEncodingConverter.ConvertVniToUnicode(raw);
                }
                else // AUTO
                {
                    converted = VietnameseEncodingConverter.ConvertTcvn3ToUnicode(raw);
                    converted = VietnameseEncodingConverter.ConvertVniToUnicode(converted);
                }

                StartUndoRecord(app, "CoolWord: Chuyển mã sang Unicode");
                try
                {
                    range.Text = converted;
                    // Tự động chuyển phông chữ sang Times New Roman nếu đang dùng phông TCVN3 hoặc VNI
                    string currentFont = range.Font.Name;
                    if (VietnameseEncodingConverter.IsTcvn3Font(currentFont) || VietnameseEncodingConverter.IsVniFont(currentFont))
                    {
                        range.Font.Name = "Times New Roman";
                    }
                }
                finally
                {
                    EndUndoRecord(app);
                }

                statusMessage = string.Format("Đã chuyển đổi mã thành công sang Unicode chuẩn ({0}).", selectionOnly ? "vùng chọn" : "toàn văn bản");
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi chuyển mã văn bản: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Nâng cấp file .doc cũ sang định dạng .docx OpenXML chuẩn mới của Word.
        /// </summary>
        public static bool SaveAsDocx(Word.Application app, out string statusMessage)
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
                string fullName = string.Empty;
                try { fullName = doc.FullName; } catch { }

                string dir = string.IsNullOrEmpty(fullName) || !File.Exists(fullName)
                    ? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                    : Path.GetDirectoryName(fullName);

                string baseName = string.IsNullOrEmpty(fullName)
                    ? "CoolWord_Document_" + DateTime.Now.ToString("yyyyMMdd_HHmmss")
                    : Path.GetFileNameWithoutExtension(fullName);

                string newPath = Path.Combine(dir, baseName + "_New.docx");
                if (string.Equals(Path.GetExtension(fullName), ".doc", StringComparison.OrdinalIgnoreCase))
                {
                    newPath = Path.Combine(dir, baseName + ".docx");
                }

                object format = Word.WdSaveFormat.wdFormatXMLDocument;
                doc.SaveAs2(newPath, ref format);

                statusMessage = "Đã lưu sang định dạng chuẩn .docx thành công: " + Path.GetFileName(newPath);
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi lưu file .docx: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Lưu bản sao tài liệu dưới định dạng .doc (Word 97-2003) để tương thích hệ thống cũ.
        /// </summary>
        public static bool SaveAsDoc97(Word.Application app, out string statusMessage)
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
                string fullName = string.Empty;
                try { fullName = doc.FullName; } catch { }

                string dir = string.IsNullOrEmpty(fullName) || !File.Exists(fullName)
                    ? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                    : Path.GetDirectoryName(fullName);

                string baseName = string.IsNullOrEmpty(fullName)
                    ? "CoolWord_Document_" + DateTime.Now.ToString("yyyyMMdd_HHmmss")
                    : Path.GetFileNameWithoutExtension(fullName);

                string newPath = Path.Combine(dir, baseName + "_Word97.doc");
                object format = Word.WdSaveFormat.wdFormatDocument97;
                doc.SaveAs2(newPath, ref format);

                statusMessage = "Đã lưu bản sao tương thích .doc (Word 97-2003): " + Path.GetFileName(newPath);
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi lưu file .doc: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Gộp nhiều file Word thành 1 tài liệu duy nhất (kèm ngắt trang giữa các file).
        /// </summary>
        public static bool MergeWordDocuments(Word.Application app, string[] filePaths, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null)
            {
                statusMessage = "Không tìm thấy phiên bản Word đang chạy.";
                return false;
            }

            string[] files = filePaths;
            if (files == null || files.Length == 0)
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Title = "Chọn các file Word cần gộp (.docx, .doc)";
                    ofd.Filter = "Tài liệu Word (*.docx;*.doc)|*.docx;*.doc|Tất cả tệp (*.*)|*.*";
                    ofd.Multiselect = true;
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        files = ofd.FileNames;
                    }
                }
            }

            if (files == null || files.Length == 0)
            {
                statusMessage = "Chưa chọn file để gộp.";
                return false;
            }

            try
            {
                Word.Document doc;
                if (app.Documents.Count == 0)
                {
                    doc = app.Documents.Add();
                }
                else
                {
                    doc = app.ActiveDocument;
                }

                StartUndoRecord(app, "CoolWord: Gộp file Word");
                try
                {
                    int mergedCount = 0;
                    foreach (string file in files)
                    {
                        if (!File.Exists(file)) continue;

                        Word.Range endRange = doc.Content.Duplicate;
                        object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                        endRange.Collapse(ref collapseEnd);

                        if (mergedCount > 0 || doc.Content.Text.Length > 2)
                        {
                            object breakPage = Word.WdBreakType.wdPageBreak;
                            endRange.InsertBreak(ref breakPage);
                            endRange.Collapse(ref collapseEnd);
                        }

                        endRange.InsertFile(file);
                        mergedCount++;
                    }

                    statusMessage = string.Format("Đã gộp thành công {0} file Word vào tài liệu.", mergedCount);
                    return true;
                }
                finally
                {
                    EndUndoRecord(app);
                }
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi gộp file Word: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Tách tài liệu Word thành các file riêng biệt theo từng Section.
        /// </summary>
        public static bool SplitDocumentBySections(Word.Application app, string outputDir, out string statusMessage)
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
                if (doc.Sections.Count <= 1)
                {
                    statusMessage = "Tài liệu chỉ có 1 Section, không thể tách theo Section.";
                    return false;
                }

                string dir = outputDir;
                if (string.IsNullOrEmpty(dir))
                {
                    string fullName = string.Empty;
                    try { fullName = doc.FullName; } catch { }

                    dir = string.IsNullOrEmpty(fullName) || !File.Exists(fullName)
                        ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "CoolWord_SplitSections")
                        : Path.Combine(Path.GetDirectoryName(fullName), Path.GetFileNameWithoutExtension(fullName) + "_Sections");
                }

                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                int count = 0;
                int total = doc.Sections.Count;

                for (int i = 1; i <= total; i++)
                {
                    Word.Section sec = doc.Sections[i];
                    sec.Range.Copy();

                    Word.Document newDoc = app.Documents.Add();
                    newDoc.Content.Paste();

                    string partName = string.Format("Phan_{0:00}.docx", i);
                    string destPath = Path.Combine(dir, partName);

                    object format = Word.WdSaveFormat.wdFormatXMLDocument;
                    newDoc.SaveAs2(destPath, ref format);
                    ((Word._Document)newDoc).Close(false);

                    count++;
                }

                statusMessage = string.Format("Đã tách thành công {0} file Word theo Section tại thư mục: {1}", count, Path.GetFileName(dir));
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi tách Section: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Tách tài liệu Word thành các file riêng biệt theo từng Trang.
        /// </summary>
        public static bool SplitDocumentByPages(Word.Application app, string outputDir, out string statusMessage)
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
                int totalPages = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);

                if (totalPages <= 1)
                {
                    statusMessage = "Tài liệu chỉ có 1 trang, không cần tách.";
                    return false;
                }

                string dir = outputDir;
                if (string.IsNullOrEmpty(dir))
                {
                    string fullName = string.Empty;
                    try { fullName = doc.FullName; } catch { }

                    dir = string.IsNullOrEmpty(fullName) || !File.Exists(fullName)
                        ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "CoolWord_SplitPages")
                        : Path.Combine(Path.GetDirectoryName(fullName), Path.GetFileNameWithoutExtension(fullName) + "_Pages");
                }

                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                int count = 0;
                for (int p = 1; p <= totalPages; p++)
                {
                    object what = Word.WdGoToItem.wdGoToPage;
                    object which = Word.WdGoToDirection.wdGoToAbsolute;
                    object countObj = p;

                    Word.Range r = doc.GoTo(ref what, ref which, ref countObj);
                    Word.Range pageRange = r.Bookmarks.get_Item(@"\page").Range;

                    pageRange.Copy();
                    Word.Document newDoc = app.Documents.Add();
                    newDoc.Content.Paste();

                    string pageName = string.Format("Trang_{0:000}.docx", p);
                    string destPath = Path.Combine(dir, pageName);

                    object format = Word.WdSaveFormat.wdFormatXMLDocument;
                    newDoc.SaveAs2(destPath, ref format);
                    ((Word._Document)newDoc).Close(false);

                    count++;
                }

                statusMessage = string.Format("Đã tách thành công {0} trang thành các file Word riêng tại: {1}", count, Path.GetFileName(dir));
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi tách trang Word: " + ex.Message;
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
