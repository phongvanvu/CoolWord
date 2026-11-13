using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using CoolWord.Modules.AdminFormat;
using CoolWord.Modules.NameSplit;
using CoolWord.Modules.NumberToText;
using CoolWord.Modules.Pdf;
using CoolWord.Modules.Printing;
using CoolWord.Modules.QuickInput;
using CoolWord.Modules.WordToWord;
using CoolWord.Modules.WordToExcel;
using CoolWord.Modules.WordToPowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace CoolWord.Core
{
    public class WordDocumentService : IWordDocumentService
    {
        private Word.Application _wordApp;

        public event EventHandler<WordSelectionEventArgs> SelectionChanged;

        public WordDocumentService(object wordApp = null)
        {
            Attach(wordApp);
        }

        public void Attach(object wordApp)
        {
            if (_wordApp != null)
            {
                try
                {
                    _wordApp.WindowSelectionChange -= OnWindowSelectionChange;
                }
                catch { }
            }

            _wordApp = wordApp as Word.Application;

            if (_wordApp != null)
            {
                try
                {
                    _wordApp.WindowSelectionChange += OnWindowSelectionChange;
                }
                catch { }
            }
        }

        private Word.Application GetActiveWordApp()
        {
            if (_wordApp != null)
            {
                return _wordApp;
            }

            try
            {
                object activeObj = Marshal.GetActiveObject("Word.Application");
                _wordApp = activeObj as Word.Application;
                if (_wordApp != null)
                {
                    _wordApp.WindowSelectionChange += OnWindowSelectionChange;
                }
            }
            catch
            {
                _wordApp = null;
            }

            return _wordApp;
        }

        public bool IsConnected
        {
            get
            {
                Word.Application app = GetActiveWordApp();
                if (app == null) return false;
                try
                {
                    return app.Documents.Count > 0;
                }
                catch
                {
                    return false;
                }
            }
        }

        private void OnWindowSelectionChange(Word.Selection sel)
        {
            try
            {
                EventHandler<WordSelectionEventArgs> handler = SelectionChanged;
                if (handler != null && sel != null)
                {
                    string text = CleanWordRangeText(sel.Text);
                    handler(this, new WordSelectionEventArgs { SelectedText = text });
                }
            }
            catch { }
        }

        public string GetSelectedText()
        {
            Word.Application app = GetActiveWordApp();
            if (app == null) return string.Empty;

            try
            {
                if (app.Documents.Count == 0 || app.Selection == null)
                    return string.Empty;

                Word.Range range = GetTrimmedSelectionOrWordUnderCursor(app.Selection);
                if (range == null)
                    return string.Empty;

                return CleanWordRangeText(range.Text);
            }
            catch
            {
                return string.Empty;
            }
        }

        public bool ConvertSelectedNumber(NumberConvertOptions options, out string convertedWords, out string statusMessage)
        {
            convertedWords = string.Empty;
            statusMessage = string.Empty;

            Word.Application app = GetActiveWordApp();
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            if (options == null)
            {
                options = SettingsService.Current.ToNumberConvertOptions();
            }

            try
            {
                Word.Range targetRange = GetTrimmedSelectionOrWordUnderCursor(app.Selection);
                if (targetRange == null)
                {
                    statusMessage = "Vui lòng bôi đen một con số trong văn bản Word.";
                    return false;
                }

                string rawText = CleanWordRangeText(targetRange.Text);
                if (string.IsNullOrWhiteSpace(rawText))
                {
                    statusMessage = "Vui lòng bôi đen một con số trong văn bản Word.";
                    return false;
                }

                ParsedNumberInfo parsed = NumberExtractor.Extract(rawText, options.ReadingMode);
                if (!parsed.IsValid)
                {
                    statusMessage = "Vùng đang chọn không chứa số hợp lệ: \"" + Truncate(rawText, 30) + "\"";
                    return false;
                }

                convertedWords = VietnameseNumberConverter.ConvertParsed(parsed, options);
                if (string.IsNullOrEmpty(convertedWords))
                {
                    statusMessage = "Không thể chuyển đổi số đang chọn.";
                    return false;
                }

                StartUndoRecord(app, "CoolWord: Đổi số ra chữ");
                try
                {
                    ApplyConversionToRange(targetRange, rawText, convertedWords, options);
                }
                finally
                {
                    EndUndoRecord(app);
                }

                SettingsService.Current.AddHistory(parsed.MatchedText, convertedWords);
                SettingsService.Save();

                statusMessage = "Đã đổi: " + parsed.MatchedText + " → " + convertedWords;
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi thao tác với Word: " + ex.Message;
                return false;
            }
        }

        public int ConvertAllNumbersInSelection(NumberConvertOptions options, out string statusMessage)
        {
            statusMessage = string.Empty;
            Word.Application app = GetActiveWordApp();
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return 0;
            }

            if (options == null)
            {
                options = SettingsService.Current.ToNumberConvertOptions();
            }

            try
            {
                Word.Selection sel = app.Selection;
                if (sel == null || sel.Range == null || sel.Range.Start == sel.Range.End)
                {
                    statusMessage = "Vui lòng bôi đen vùng văn bản chứa các con số cần đổi.";
                    return 0;
                }

                Word.Range range = GetTrimmedRange(sel.Range.Duplicate);
                string text = range.Text;
                if (string.IsNullOrWhiteSpace(text))
                {
                    statusMessage = "Vùng chọn trống.";
                    return 0;
                }

                int count = 0;
                Regex batchRegex = new Regex(@"(?<!\w)[-+]?\d(?:[\d.,]*\d)?(?!\w)");
                string replaced = batchRegex.Replace(text, match =>
                {
                    string token = match.Value;
                    ParsedNumberInfo parsed = NumberExtractor.Extract(token, options.ReadingMode);
                    if (!parsed.IsValid) return token;

                    string words = VietnameseNumberConverter.ConvertParsed(parsed, options);
                    if (string.IsNullOrEmpty(words)) return token;

                    count++;
                    return VietnameseNumberConverter.FormatForWordInsertion(token, words, options);
                });

                if (count > 0)
                {
                    StartUndoRecord(app, "CoolWord: Đổi hàng loạt số ra chữ");
                    try
                    {
                        range.Text = replaced;
                    }
                    finally
                    {
                        EndUndoRecord(app);
                    }
                    statusMessage = string.Format("Đã chuyển đổi thành công {0} con số trong vùng chọn.", count);
                }
                else
                {
                    statusMessage = "Không tìm thấy con số nào trong vùng bôi đen.";
                }

                return count;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi chuyển đổi hàng loạt: " + ex.Message;
                return 0;
            }
        }

        public bool InsertConvertedText(string originalInput, string convertedWords, NumberConvertOptions options, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (string.IsNullOrWhiteSpace(convertedWords))
            {
                statusMessage = "Chưa có nội dung chữ để chèn.";
                return false;
            }

            Word.Application app = GetActiveWordApp();
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở để chèn.";
                return false;
            }

            if (options == null)
            {
                options = SettingsService.Current.ToNumberConvertOptions();
            }

            try
            {
                Word.Selection sel = app.Selection;
                Word.Range range = GetTrimmedRange(sel.Range.Duplicate);

                StartUndoRecord(app, "CoolWord: Chèn số bằng chữ");
                try
                {
                    bool hasSelection = range.End > range.Start && !string.IsNullOrWhiteSpace(CleanWordRangeText(range.Text));
                    if (hasSelection)
                    {
                        string selectedRaw = CleanWordRangeText(range.Text);
                        ApplyConversionToRange(range, selectedRaw, convertedWords, options);
                    }
                    else
                    {
                        // Nếu con trỏ đang đứng trống (không bôi đen), chèn trực tiếp kết quả tại vị trí con trỏ
                        range.Text = convertedWords;
                        if (options.ItalicizeConvertedText)
                        {
                            range.Italic = 1;
                        }
                        object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                        range.Collapse(ref collapseEnd);
                        range.Select();
                    }
                }
                finally
                {
                    EndUndoRecord(app);
                }

                statusMessage = "Đã chèn vào văn bản Word: " + convertedWords;
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi chèn vào Word: " + ex.Message;
                return false;
            }
        }

        public bool TransformSelectedText(Func<string, string> transformer, string actionName, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (transformer == null) return false;

            Word.Application app = GetActiveWordApp();
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
                    statusMessage = "Vui lòng bôi đen đoạn văn bản cần xử lý.";
                    return false;
                }

                Word.Range range = GetTrimmedRange(sel.Range.Duplicate);
                string raw = CleanWordRangeText(range.Text);
                if (string.IsNullOrEmpty(raw))
                {
                    statusMessage = "Vùng văn bản đang chọn trống.";
                    return false;
                }

                string transformed = transformer(raw);
                StartUndoRecord(app, "CoolWord: " + actionName);
                try
                {
                    range.Text = transformed;
                    range.Select();
                }
                finally
                {
                    EndUndoRecord(app);
                }

                statusMessage = "Đã thực hiện: " + actionName;
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi xử lý văn bản: " + ex.Message;
                return false;
            }
        }

        public bool ApplyDecree30FullStandard(AdminDocOptions options, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return Decree30Formatter.ApplyFullStandardOneClick(app, options, out statusMessage);
        }

        public bool ApplyDecree30MarginsOnly(AdminDocOptions options, out string statusMessage)
        {
            statusMessage = string.Empty;
            Word.Application app = GetActiveWordApp();
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            if (options == null) options = new AdminDocOptions();
            try
            {
                StartUndoRecord(app, "CoolWord: Căn lề chuẩn NĐ 30/2020");
                try
                {
                    Decree30Formatter.ApplyPageMargins(app.ActiveDocument, options);
                    if (options.EnableDecree30PageNumbering)
                    {
                        Decree30Formatter.ApplyDecree30PageNumbering(app.ActiveDocument, options);
                    }
                }
                finally
                {
                    EndUndoRecord(app);
                }
                statusMessage = string.Format("Đã căn lề chuẩn A4 NĐ 30/2020: Trên {0}mm - Dưới {1}mm - Trái {2}mm - Phải {3}mm.",
                    options.TopMarginMm, options.BottomMarginMm, options.LeftMarginMm, options.RightMarginMm);
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi căn lề: " + ex.Message;
                return false;
            }
        }

        public bool InsertDecree30Header(AdminDocOptions options, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return Decree30Formatter.InsertHeaderBlock(app, options, out statusMessage);
        }

        public bool InsertDecree30FooterSignature(AdminDocOptions options, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return Decree30Formatter.InsertFooterSignatureBlock(app, options, out statusMessage);
        }

        public bool CleanFullDocumentGarbage(out string statusMessage)
        {
            statusMessage = string.Empty;
            Word.Application app = GetActiveWordApp();
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                int cleaned;
                StartUndoRecord(app, "CoolWord: Dọn dẹp định dạng rác toàn văn bản");
                try
                {
                    cleaned = Decree30Formatter.CleanDocumentParagraphGarbage(app.ActiveDocument);
                }
                finally
                {
                    EndUndoRecord(app);
                }

                statusMessage = string.Format("Đã dọn dẹp xong định dạng rác trên {0} đoạn văn trong tài liệu.", cleaned);
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi dọn dẹp văn bản: " + ex.Message;
                return false;
            }
        }

        private static void ApplyConversionToRange(Word.Range targetRange, string originalRaw, string convertedWords, NumberConvertOptions options)
        {
            switch (options.InsertMode)
            {
                case WordInsertMode.ReplaceSelection:
                    targetRange.Text = convertedWords;
                    if (options.ItalicizeConvertedText)
                    {
                        targetRange.Italic = 1;
                    }
                    break;

                case WordInsertMode.AppendInParentheses:
                case WordInsertMode.AppendWithLabelInParentheses:
                    {
                        string prefix = options.InsertMode == WordInsertMode.AppendWithLabelInParentheses
                            ? "Bằng chữ: "
                            : string.Empty;

                        Word.Range insertRange = targetRange.Duplicate;
                        object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                        insertRange.Collapse(ref collapseEnd);

                        string suffixText = " (" + prefix + convertedWords + ")";
                        insertRange.Text = suffixText;

                        if (options.ItalicizeConvertedText)
                        {
                            // Chỉ in nghiêng phần chữ nằm bên trong dấu ngoặc
                            Word.Range innerRange = insertRange.Duplicate;
                            innerRange.Start = insertRange.Start + 2; // Bỏ qua " ("
                            innerRange.End = insertRange.End - 1;     // Bỏ qua ")"
                            innerRange.Italic = 1;
                        }

                        insertRange.Collapse(ref collapseEnd);
                        insertRange.Select();
                    }
                    break;

                case WordInsertMode.AppendOnNewLine:
                    {
                        Word.Range insertRange = targetRange.Duplicate;
                        object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                        insertRange.Collapse(ref collapseEnd);

                        string lineText = "\rBằng chữ: " + convertedWords;
                        insertRange.Text = lineText;
                        if (options.ItalicizeConvertedText)
                        {
                            insertRange.Italic = 1;
                        }
                        insertRange.Collapse(ref collapseEnd);
                        insertRange.Select();
                    }
                    break;
            }
        }

        private static Word.Range GetTrimmedSelectionOrWordUnderCursor(Word.Selection sel)
        {
            if (sel == null || sel.Range == null) return null;

            Word.Range range = sel.Range.Duplicate;
            if (range.Start == range.End)
            {
                // Tự động mở rộng vùng chọn nếu con trỏ đang đặt ngay trong hoặc sát cạnh một con số
                try
                {
                    Word.Document doc = sel.Document;
                    int docEnd = doc.Content.End;
                    int start = range.Start;
                    int end = range.End;

                    while (start > 0)
                    {
                        Word.Range prev = doc.Range(start - 1, start);
                        string ch = prev.Text;
                        if (!string.IsNullOrEmpty(ch) && IsNumberTokenChar(ch[0]))
                        {
                            start--;
                        }
                        else break;
                    }

                    while (end < docEnd)
                    {
                        Word.Range next = doc.Range(end, end + 1);
                        string ch = next.Text;
                        if (!string.IsNullOrEmpty(ch) && IsNumberTokenChar(ch[0]))
                        {
                            end++;
                        }
                        else break;
                    }

                    if (end > start)
                    {
                        range.SetRange(start, end);
                    }
                }
                catch { }
            }

            return GetTrimmedRange(range);
        }

        private static bool IsNumberTokenChar(char c)
        {
            return char.IsDigit(c) || c == '.' || c == ',' || c == '-' || c == '$' || c == 'đ' || c == 'Đ';
        }

        private static Word.Range GetTrimmedRange(Word.Range range)
        {
            if (range == null) return null;

            try
            {
                // Thu gọn các ký tự điều khiển cuối đoạn (\r, \a, \n) và khoảng trắng thừa ở hai đầu Range
                while (range.End > range.Start)
                {
                    string currentText = range.Text;
                    if (string.IsNullOrEmpty(currentText)) break;

                    char lastChar = currentText[currentText.Length - 1];
                    if (lastChar == '\r' || lastChar == '\n' || lastChar == '\a' || lastChar == '\v' || lastChar == '\f' || char.IsWhiteSpace(lastChar))
                    {
                        range.End -= 1;
                    }
                    else
                    {
                        break;
                    }
                }

                while (range.Start < range.End)
                {
                    string currentText = range.Text;
                    if (string.IsNullOrEmpty(currentText)) break;

                    char firstChar = currentText[0];
                    if (char.IsWhiteSpace(firstChar))
                    {
                        range.Start += 1;
                    }
                    else
                    {
                        break;
                    }
                }
            }
            catch { }

            return range;
        }

        private static string CleanWordRangeText(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            return text.Replace("\r\a", "")
                       .Replace("\a", "")
                       .Replace("\r", " ")
                       .Replace("\v", " ")
                       .Trim();
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

        private static string Truncate(string value, int maxLen)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= maxLen) return value;
            return value.Substring(0, maxLen) + "...";
        }

        #region Tác vụ In ấn & Xuất bản

        public bool PrintCurrentPage(out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return PrintService.PrintCurrentPage(app, out statusMessage);
        }

        public bool PrintSelection(out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return PrintService.PrintSelection(app, out statusMessage);
        }

        public bool PrintManualDuplex(out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return PrintService.PrintManualDuplex(app, out statusMessage);
        }

        public bool PrintOddPagesOnly(out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return PrintService.PrintOddPagesOnly(app, out statusMessage);
        }

        public bool PrintEvenPagesOnly(out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return PrintService.PrintEvenPagesOnly(app, out statusMessage);
        }

        public bool PrintTwoPagesPerSheet(out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return PrintService.PrintTwoPagesPerSheet(app, out statusMessage);
        }

        public bool ExportQuickPdf(string targetFilePath, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return PrintService.ExportQuickPdf(app, targetFilePath, out statusMessage);
        }

        public bool PrintMultiCopyAccountingVoucher(int copies, string voucherType, bool fitTwoPerSheet, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return PrintService.PrintMultiCopyAccountingVoucher(app, copies, voucherType, fitTwoPerSheet, out statusMessage);
        }

        #endregion

        #region Tác vụ Tiện ích PDF & Kế toán

        public bool ExportCustomPdf(PdfOptions options, string targetFilePath, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return PdfService.ExportPdf(app, options, targetFilePath, out statusMessage);
        }

        public bool ExportAccountingPdfA(string targetFilePath, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return PdfService.ExportAccountingPdfA(app, targetFilePath, out statusMessage);
        }

        public bool ExportSelectionToPdf(string targetFilePath, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return PdfService.ExportSelectionToPdf(app, targetFilePath, out statusMessage);
        }

        public bool ExportCurrentPageToPdf(string targetFilePath, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return PdfService.ExportCurrentPageToPdf(app, targetFilePath, out statusMessage);
        }

        public bool AddAccountingWatermark(string watermarkText, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return PdfService.AddAccountingWatermark(app, watermarkText, out statusMessage);
        }

        public bool RemoveAccountingWatermark(out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return PdfService.RemoveAccountingWatermark(app, out statusMessage);
        }

        #endregion

        #region Tác vụ Nhập Liệu Nhanh

        public bool InsertTextSnippet(string text, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return QuickInputService.InsertSnippetIntoWord(app, text, out statusMessage);
        }

        public bool SaveCurrentSelectionAsSnippet(string title, string category, string shortcut, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return QuickInputService.SaveSelectionAsCustomSnippet(app, title, category, shortcut, out statusMessage);
        }

        #endregion

        #region Tác vụ Tách Họ - Tên

        public bool SplitTableColumn2(out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return NameSplitService.SplitTableColumn2(app, out statusMessage);
        }

        public bool SplitTableColumn3(out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return NameSplitService.SplitTableColumn3(app, out statusMessage);
        }

        public bool MergeTableColumns(out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return NameSplitService.MergeTableColumns(app, out statusMessage);
        }

        public bool SplitSelectedTextLines(string delimiter, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return NameSplitService.SplitSelectedTextLines(app, delimiter, out statusMessage);
        }

        public bool NormalizeSelectedNames(out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return NameSplitService.NormalizeSelectedNames(app, out statusMessage);
        }

        #endregion

        #region Tác vụ Word ➔ Word

        public bool ConvertEncodingToUnicode(bool selectionOnly, string sourceEncoding, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return WordToWordService.ConvertEncodingToUnicode(app, selectionOnly, sourceEncoding, out statusMessage);
        }

        public bool SaveAsDocx(out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return WordToWordService.SaveAsDocx(app, out statusMessage);
        }

        public bool SaveAsDoc97(out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return WordToWordService.SaveAsDoc97(app, out statusMessage);
        }

        public bool MergeWordDocuments(string[] filePaths, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return WordToWordService.MergeWordDocuments(app, filePaths, out statusMessage);
        }

        public bool SplitDocumentBySections(string outputDir, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return WordToWordService.SplitDocumentBySections(app, outputDir, out statusMessage);
        }

        public bool SplitDocumentByPages(string outputDir, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return WordToWordService.SplitDocumentByPages(app, outputDir, out statusMessage);
        }

        #endregion

        #region Tác vụ Word ➔ Excel

        public bool ExportAllTablesToExcel(string outputPath, bool openAfterExport, bool separateSheets, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return WordToExcelService.ExportAllTablesToExcel(app, outputPath, openAfterExport, separateSheets, out statusMessage);
        }

        public bool ExportSelectedTableToExcel(string outputPath, bool openAfterExport, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return WordToExcelService.ExportSelectedTableToExcel(app, outputPath, openAfterExport, out statusMessage);
        }

        public bool ExportAccountingTableToExcel(string outputPath, bool openAfterExport, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return WordToExcelService.ExportAccountingTableToExcel(app, outputPath, openAfterExport, out statusMessage);
        }

        public bool ExportParagraphsToExcel(string outputPath, bool openAfterExport, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return WordToExcelService.ExportParagraphsToExcel(app, outputPath, openAfterExport, out statusMessage);
        }

        #endregion

        #region Tác vụ Word ➔ PowerPoint

        public bool CreateSlidesFromHeadings(string outputPath, bool openAfterExport, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return WordToPowerPointService.CreateSlidesFromHeadings(app, outputPath, openAfterExport, out statusMessage);
        }

        public bool CreateSlidesFromPages(string outputPath, bool openAfterExport, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return WordToPowerPointService.CreateSlidesFromPages(app, outputPath, openAfterExport, out statusMessage);
        }

        public bool CreateSlidesFromSections(string outputPath, bool openAfterExport, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return WordToPowerPointService.CreateSlidesFromSections(app, outputPath, openAfterExport, out statusMessage);
        }

        public bool ExportTablesToSlides(string outputPath, bool openAfterExport, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return WordToPowerPointService.ExportTablesToSlides(app, outputPath, openAfterExport, out statusMessage);
        }

        public bool CreateFinancialPresentation(string outputPath, bool openAfterExport, out string statusMessage)
        {
            Word.Application app = GetActiveWordApp();
            return WordToPowerPointService.CreateFinancialPresentation(app, outputPath, openAfterExport, out statusMessage);
        }

        #endregion
    }
}
