using System;
using System.Collections.Generic;
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

        /// <summary>
        /// Tạo bản sao lưu an toàn cho danh sách file .docm trước khi thực hiện chuyển đổi.
        /// Thư mục sao lưu được đặt tên theo định dạng _Backup_DOCM_yyyyMMdd_HHmmss.
        /// </summary>
        public static bool CreateSafeBackup(string folderPath, IEnumerable<string> files, out string backupDir, out string backupMessage)
        {
            backupDir = string.Empty;
            backupMessage = string.Empty;

            if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
            {
                backupMessage = "Thư mục nguồn không hợp lệ hoặc không tồn tại.";
                return false;
            }

            if (files == null)
            {
                backupMessage = "Danh sách tệp tin cần sao lưu trống.";
                return false;
            }

            try
            {
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                backupDir = Path.Combine(folderPath, "_Backup_DOCM_" + timestamp);
                if (!Directory.Exists(backupDir))
                {
                    Directory.CreateDirectory(backupDir);
                }

                int copiedCount = 0;
                string normalizedBase = Path.GetFullPath(folderPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;

                foreach (string file in files)
                {
                    if (string.IsNullOrEmpty(file) || !File.Exists(file)) continue;

                    string fullPath = Path.GetFullPath(file);
                    string targetBackupFile;

                    if (fullPath.StartsWith(normalizedBase, StringComparison.OrdinalIgnoreCase))
                    {
                        // Giữ nguyên cấu trúc thư mục con nếu file nằm trong subfolder
                        string relative = fullPath.Substring(normalizedBase.Length);
                        targetBackupFile = Path.Combine(backupDir, relative);
                        string subDir = Path.GetDirectoryName(targetBackupFile);
                        if (!string.IsNullOrEmpty(subDir) && !Directory.Exists(subDir))
                        {
                            Directory.CreateDirectory(subDir);
                        }
                    }
                    else
                    {
                        targetBackupFile = Path.Combine(backupDir, Path.GetFileName(file));
                    }

                    File.Copy(fullPath, targetBackupFile, true);
                    copiedCount++;
                }

                if (copiedCount == 0)
                {
                    backupMessage = "Không có tệp tin nào được sao lưu.";
                    return false;
                }

                backupMessage = string.Format("Đã sao lưu an toàn {0} tệp tin tại thư mục: {1}", copiedCount, Path.GetFileName(backupDir));
                return true;
            }
            catch (Exception ex)
            {
                backupMessage = "Lỗi trong quá trình tạo bản sao lưu an toàn: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Chuyển đổi toàn bộ các file Word Macro (.docm) trong một thư mục sang định dạng chuẩn .docx (OpenXML).
        /// Hỗ trợ tùy chọn sao lưu an toàn toàn bộ file gốc trước khi chuyển đổi.
        /// </summary>
        public static bool ConvertDocmFolderToDocx(
            Word.Application app,
            string folderPath,
            bool includeSubdirectories,
            bool safeBackup,
            out string statusMessage)
        {
            statusMessage = string.Empty;

            if (string.IsNullOrEmpty(folderPath))
            {
                statusMessage = "Chưa chọn thư mục cần chuyển đổi.";
                return false;
            }

            if (!Directory.Exists(folderPath))
            {
                statusMessage = "Thư mục không tồn tại: " + folderPath;
                return false;
            }

            string targetFolder = folderPath;

            // Quét danh sách file .docm
            SearchOption searchOption = includeSubdirectories ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            string[] rawFiles;
            try
            {
                rawFiles = Directory.GetFiles(targetFolder, "*.docm", searchOption);
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi quét thư mục: " + ex.Message;
                return false;
            }

            List<string> docmFiles = new List<string>();
            foreach (string file in rawFiles)
            {
                string name = Path.GetFileName(file);
                // Bỏ qua các file tạm/khóa do Word tự tạo khi mở (bắt đầu bằng ~$)
                if (name.StartsWith("~$", StringComparison.OrdinalIgnoreCase)) continue;
                docmFiles.Add(file);
            }

            if (docmFiles.Count == 0)
            {
                statusMessage = "Không tìm thấy file .docm nào trong thư mục: " + Path.GetFileName(targetFolder);
                return false;
            }

            // Thực hiện sao lưu an toàn nếu được yêu cầu
            string backupDir = string.Empty;
            if (safeBackup)
            {
                string backupMsg;
                bool backedUp = CreateSafeBackup(targetFolder, docmFiles, out backupDir, out backupMsg);
                if (!backedUp)
                {
                    statusMessage = "Dừng chuyển đổi vì không thể tạo bản sao lưu an toàn: " + backupMsg;
                    return false;
                }
            }

            bool tempAppCreated = false;
            Word.Application wordApp = app;

            try
            {
                if (wordApp == null)
                {
                    try
                    {
                        wordApp = new Word.Application();
                        wordApp.Visible = false;
                        wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                        tempAppCreated = true;
                    }
                    catch (Exception ex)
                    {
                        statusMessage = "Không thể khởi động Microsoft Word để chuyển đổi: " + ex.Message;
                        return false;
                    }
                }

                // Tắt cảnh báo khi mở file có macro
                Word.WdAlertLevel originalAlerts = Word.WdAlertLevel.wdAlertsNone;
                try
                {
                    originalAlerts = wordApp.DisplayAlerts;
                    wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                }
                catch { }

                int convertedCount = 0;
                List<string> errorList = new List<string>();

                foreach (string file in docmFiles)
                {
                    Word.Document doc = null;
                    try
                    {
                        object fileName = file;
                        object confirmConversions = false;
                        object readOnly = true;
                        object isVisible = false;
                        object addToRecentFiles = false;

                        doc = wordApp.Documents.Open(
                            ref fileName,
                            ref confirmConversions,
                            ref readOnly,
                            ref addToRecentFiles,
                            ref Missing, ref Missing, ref Missing, ref Missing,
                            ref Missing, ref Missing, ref Missing,
                            ref isVisible,
                            ref Missing, ref Missing, ref Missing, ref Missing);

                        if (doc == null)
                        {
                            errorList.Add(Path.GetFileName(file) + " (Không thể mở)");
                            continue;
                        }

                        string targetDocxPath = Path.ChangeExtension(file, ".docx");
                        object targetPath = targetDocxPath;
                        object format = Word.WdSaveFormat.wdFormatXMLDocument; // 12: wdFormatXMLDocument (.docx tiêu chuẩn, không macro)

                        doc.SaveAs2(ref targetPath, ref format);
                        convertedCount++;
                    }
                    catch (Exception ex)
                    {
                        errorList.Add(Path.GetFileName(file) + " (" + ex.Message + ")");
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            try
                            {
                                object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                                ((Word._Document)doc).Close(ref saveChanges);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                            }
                            catch { }
                            doc = null;
                        }
                    }
                }

                try
                {
                    wordApp.DisplayAlerts = originalAlerts;
                }
                catch { }

                string backupInfo = safeBackup && !string.IsNullOrEmpty(backupDir)
                    ? string.Format(" Đã sao lưu an toàn tại thư mục '{0}'.", Path.GetFileName(backupDir))
                    : string.Empty;

                if (convertedCount == docmFiles.Count)
                {
                    statusMessage = string.Format("Đã chuyển đổi thành công toàn bộ {0}/{1} file .docm sang .docx.{2}",
                        convertedCount, docmFiles.Count, backupInfo);
                    return true;
                }
                else if (convertedCount > 0)
                {
                    statusMessage = string.Format("Đã chuyển đổi thành công {0}/{1} file .docm sang .docx (Lỗi {2} file: {3}).{4}",
                        convertedCount, docmFiles.Count, errorList.Count, string.Join(", ", errorList.ToArray()), backupInfo);
                    return true;
                }
                else
                {
                    statusMessage = string.Format("Không thể chuyển đổi file nào ({0} lỗi): {1}.{2}",
                        errorList.Count, string.Join("; ", errorList.ToArray()), backupInfo);
                    return false;
                }
            }
            finally
            {
                if (tempAppCreated && wordApp != null)
                {
                    try
                    {
                        object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                        ((Word._Application)wordApp).Quit(ref saveChanges);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                    }
                    catch { }
                }
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
