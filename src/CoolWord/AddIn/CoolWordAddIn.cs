using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using CoolWord.Core;
using CoolWord.Modules.AdminFormat;
using CoolWord.Modules.NameSplit;
using CoolWord.Modules.NumberToText;
using CoolWord.Modules.Pdf;
using CoolWord.Modules.QuickInput;
using CoolWord.Modules.TextUtilities;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace CoolWord.AddIn
{
    [ComVisible(true)]
    [Guid("4D7E9A11-8B2C-4F6D-9E3A-1C2B3D4E5F6A")]
    [ProgId("CoolWord.AddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class CoolWordAddIn : IDTExtensibility2, Office.IRibbonExtensibility, Office.ICustomTaskPaneConsumer
    {
        private Word.Application _wordApp;
        private WordDocumentService _wordService;
        private Office.IRibbonUI _ribbon;
        private Office.ICTPFactory _ctpFactory;
        private Office.CustomTaskPane _customTaskPane;
        private TaskPaneHost _taskPaneHost;

        static CoolWordAddIn()
        {
            AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
        }

        private static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                AssemblyName requestedName = new AssemblyName(args.Name);
                string baseDir = Path.GetDirectoryName(typeof(CoolWordAddIn).Assembly.Location);
                if (!string.IsNullOrEmpty(baseDir))
                {
                    string candidate = Path.Combine(baseDir, requestedName.Name + ".dll");
                    if (File.Exists(candidate))
                    {
                        return Assembly.LoadFrom(candidate);
                    }
                }
            }
            catch { }
            return null;
        }

        #region IDTExtensibility2

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _wordApp = Application as Word.Application;
                _wordService = new WordDocumentService(_wordApp);
                _wordService.SelectionChanged += (s, e) =>
                {
                    if (_ribbon != null)
                    {
                        try { _ribbon.Invalidate(); } catch { }
                    }
                };

                SettingsService.SettingsChanged += (s, e) =>
                {
                    if (_ribbon != null)
                    {
                        try { _ribbon.Invalidate(); } catch { }
                    }
                };
            }
            catch { }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if (_customTaskPane != null)
                {
                    _customTaskPane.Visible = false;
                    _customTaskPane = null;
                }
                if (_wordService != null)
                {
                    _wordService.Attach(null);
                    _wordService = null;
                }
                _wordApp = null;
            }
            catch { }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                Assembly asm = typeof(CoolWordAddIn).Assembly;
                using (Stream stream = asm.GetManifestResourceStream("CoolWord.AddIn.Ribbon.xml"))
                {
                    if (stream != null)
                    {
                        using (StreamReader reader = new StreamReader(stream))
                        {
                            return reader.ReadToEnd();
                        }
                    }
                }

                string dir = Path.GetDirectoryName(asm.Location);
                string fileCandidate = Path.Combine(dir ?? "", "Ribbon.xml");
                if (File.Exists(fileCandidate))
                {
                    return File.ReadAllText(fileCandidate);
                }
            }
            catch { }

            return string.Empty;
        }

        #endregion

        #region ICustomTaskPaneConsumer

        public void CTPFactoryAvailable(Office.ICTPFactory CTPFactoryInst)
        {
            _ctpFactory = CTPFactoryInst;
        }

        private void EnsureTaskPaneCreated()
        {
            if (_customTaskPane != null || _ctpFactory == null)
                return;

            try
            {
                _customTaskPane = _ctpFactory.CreateCTP(
                    "CoolWord.TaskPaneHost",
                    "CoolWord - Công Cụ Tiếng Việt",
                    Type.Missing);

                _customTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                _customTaskPane.Width = 380;
                _customTaskPane.VisibleStateChange += OnTaskPaneVisibleStateChange;

                _taskPaneHost = _customTaskPane.ContentControl as TaskPaneHost;
                if (_taskPaneHost != null)
                {
                    _taskPaneHost.InitializeControl(_wordService);
                }
            }
            catch (Exception ex)
            {
                NotifyStatus("Không thể khởi tạo bảng công cụ CoolWord: " + ex.Message);
            }
        }

        private void OnTaskPaneVisibleStateChange(Office.CustomTaskPane CustomTaskPaneInst)
        {
            if (_ribbon != null)
            {
                try { _ribbon.InvalidateControl("btnToggleTaskPane"); } catch { }
            }
        }

        #endregion

        #region Ribbon & Context Menu Callbacks

        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public string GetContextQuickLabel(Office.IRibbonControl control)
        {
            try
            {
                if (_wordService == null)
                    return "CoolWord: Đổi số ra chữ Tiếng Việt";

                string selected = _wordService.GetSelectedText();
                if (string.IsNullOrWhiteSpace(selected))
                    return "CoolWord: Đổi số ra chữ Tiếng Việt";

                NumberConvertOptions options = SettingsService.Current.ToNumberConvertOptions();
                ParsedNumberInfo parsed = NumberExtractor.Extract(selected, options.ReadingMode);
                if (!parsed.IsValid)
                    return "CoolWord: Đổi số ra chữ Tiếng Việt";

                string words = VietnameseNumberConverter.ConvertParsed(parsed, options);
                if (string.IsNullOrEmpty(words))
                    return "CoolWord: Đổi số ra chữ Tiếng Việt";

                string displayWords = words.Length > 52 ? words.Substring(0, 49) + "..." : words;
                return string.Format("Đổi số: \"{0}\"", displayWords);
            }
            catch
            {
                return "CoolWord: Đổi số ra chữ Tiếng Việt";
            }
        }

        public void OnConvertQuickClick(Office.IRibbonControl control)
        {
            ExecuteConvertWithMode(null);
        }

        public void OnConvertAppendParenClick(Office.IRibbonControl control)
        {
            ExecuteConvertWithMode(WordInsertMode.AppendInParentheses);
        }

        public void OnConvertReplaceClick(Office.IRibbonControl control)
        {
            ExecuteConvertWithMode(WordInsertMode.ReplaceSelection);
        }

        public void OnConvertAppendLabelClick(Office.IRibbonControl control)
        {
            ExecuteConvertWithMode(WordInsertMode.AppendWithLabelInParentheses);
        }

        public void OnConvertNewLineClick(Office.IRibbonControl control)
        {
            ExecuteConvertWithMode(WordInsertMode.AppendOnNewLine);
        }

        public void OnConvertBatchClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            NumberConvertOptions options = SettingsService.Current.ToNumberConvertOptions();
            _wordService.ConvertAllNumbersInSelection(options, out status);
            NotifyStatus(status);
        }

        public void OnCopyConvertedClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string selected = _wordService.GetSelectedText();
            NumberConvertOptions options = SettingsService.Current.ToNumberConvertOptions();
            ParsedNumberInfo parsed = NumberExtractor.Extract(selected, options.ReadingMode);
            if (!parsed.IsValid)
            {
                NotifyStatus("Vui lòng bôi đen một con số hợp lệ để sao chép kết quả đọc số.");
                return;
            }

            string words = VietnameseNumberConverter.ConvertParsed(parsed, options);
            try
            {
                Clipboard.SetText(words);
                SettingsService.Current.AddHistory(parsed.MatchedText, words);
                SettingsService.Save();
                NotifyStatus("Đã sao chép vào Clipboard: " + words);
            }
            catch (Exception ex)
            {
                NotifyStatus("Lỗi sao chép: " + ex.Message);
            }
        }

        public void OnSetUnitVnd(Office.IRibbonControl control)
        {
            SettingsService.Current.ReadingMode = (int)NumberReadingMode.CurrencyVnd;
            SettingsService.Save();
            NotifyStatus("Đã chuyển sang chế độ đọc: Tiền Việt Nam (đồng)");
        }

        public void OnSetUnitUsd(Office.IRibbonControl control)
        {
            SettingsService.Current.ReadingMode = (int)NumberReadingMode.CurrencyUsd;
            SettingsService.Save();
            NotifyStatus("Đã chuyển sang chế độ đọc: Ngoại tệ USD (đô la Mỹ / xu)");
        }

        public void OnSetUnitPure(Office.IRibbonControl control)
        {
            SettingsService.Current.ReadingMode = (int)NumberReadingMode.PureNumber;
            SettingsService.Save();
            NotifyStatus("Đã chuyển sang chế độ đọc: Số thuần túy");
        }

        public void OnSetDialectCommon(Office.IRibbonControl control)
        {
            SettingsService.Current.Dialect = (int)VietnameseDialect.MixedCommon;
            SettingsService.Save();
            NotifyStatus("Đã chọn phương ngữ: Phổ thông hành chính (nghìn, lẻ, tư)");
        }

        public void OnSetDialectNorth(Office.IRibbonControl control)
        {
            SettingsService.Current.Dialect = (int)VietnameseDialect.StandardNorth;
            SettingsService.Save();
            NotifyStatus("Đã chọn phương ngữ: Chuẩn Bắc Bộ (nghìn, linh, tư)");
        }

        public void OnSetDialectSouth(Office.IRibbonControl control)
        {
            SettingsService.Current.Dialect = (int)VietnameseDialect.StandardSouth;
            SettingsService.Save();
            NotifyStatus("Đã chọn phương ngữ: Chuẩn Nam Bộ (ngàn, lẻ, bốn)");
        }

        public bool GetItalicPressed(Office.IRibbonControl control)
        {
            return SettingsService.Current.ItalicizeConvertedText;
        }

        public void OnToggleItalic(Office.IRibbonControl control, bool pressed)
        {
            SettingsService.Current.ItalicizeConvertedText = pressed;
            SettingsService.Save();
            NotifyStatus(pressed ? "Đã bật in nghiêng chữ khi chèn." : "Đã tắt in nghiêng chữ khi chèn.");
        }

        public void OnCleanTextClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.TransformSelectedText(TextCleanModule.NormalizeVietnameseText, "Chuẩn hóa dấu câu & khoảng trắng", out status);
            NotifyStatus(status);
        }

        public void OnCaseTitleClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.TransformSelectedText(TextCleanModule.ToTitleCaseVi, "Viết Hoa Đầu Mỗi Từ", out status);
            NotifyStatus(status);
        }

        public void OnCaseSentenceClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.TransformSelectedText(TextCleanModule.ToSentenceCaseVi, "Viết hoa đầu câu", out status);
            NotifyStatus(status);
        }

        public void OnCaseUpperClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            CultureInfo vi = new CultureInfo("vi-VN");
            _wordService.TransformSelectedText(t => t.ToUpper(vi), "Đổi sang CHỮ IN HOA", out status);
            NotifyStatus(status);
        }

        public void OnCaseLowerClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            CultureInfo vi = new CultureInfo("vi-VN");
            _wordService.TransformSelectedText(t => t.ToLower(vi), "Đổi sang chữ thường", out status);
            NotifyStatus(status);
        }

        public bool GetTaskPanePressed(Office.IRibbonControl control)
        {
            return _customTaskPane != null && _customTaskPane.Visible;
        }

        public void OnToggleTaskPane(Office.IRibbonControl control, bool pressed)
        {
            EnsureTaskPaneCreated();
            if (_customTaskPane != null)
            {
                _customTaskPane.Visible = pressed;
            }
        }

        public void OnOpenTaskPaneFromContext(Office.IRibbonControl control)
        {
            EnsureTaskPaneCreated();
            if (_customTaskPane != null)
            {
                _customTaskPane.Visible = true;
            }
            if (_ribbon != null)
            {
                try { _ribbon.InvalidateControl("btnToggleTaskPane"); } catch { }
            }
        }

        public void OnDecree30OneClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.ApplyDecree30FullStandard(new AdminDocOptions(), out status);
            NotifyStatus(status);
        }

        public void OnInsertQuocHieuClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.InsertDecree30Header(new AdminDocOptions(), out status);
            NotifyStatus(status);
        }

        public void OnInsertNoiNhanChuKyClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.InsertDecree30FooterSignature(new AdminDocOptions(), out status);
            NotifyStatus(status);
        }

        public void OnDecree30MarginsOnlyClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.ApplyDecree30MarginsOnly(new AdminDocOptions(), out status);
            NotifyStatus(status);
        }

        public void OnCleanDocGarbageClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.CleanFullDocumentGarbage(out status);
            NotifyStatus(status);
        }

        public void OnOpenDecree30PaneClick(Office.IRibbonControl control)
        {
            EnsureTaskPaneCreated();
            if (_customTaskPane != null)
            {
                _customTaskPane.Visible = true;
            }
            if (_taskPaneHost != null && _taskPaneHost.WpfControl != null)
            {
                _taskPaneHost.WpfControl.SelectModule("admin-decree30");
            }
            if (_ribbon != null)
            {
                try { _ribbon.InvalidateControl("btnToggleTaskPane"); } catch { }
            }
        }

        #region Callbacks Nhóm In Ấn & Xuất Bản

        public void OnPrintCurrentPageClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.PrintCurrentPage(out status);
            NotifyStatus(status);
        }

        public void OnPrintManualDuplexClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.PrintManualDuplex(out status);
            NotifyStatus(status);
        }

        public void OnPrintOddPagesClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.PrintOddPagesOnly(out status);
            NotifyStatus(status);
        }

        public void OnPrintEvenPagesClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.PrintEvenPagesOnly(out status);
            NotifyStatus(status);
        }

        public void OnPrintSelectionClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.PrintSelection(out status);
            NotifyStatus(status);
        }

        public void OnPrintTwoPagesPerSheetClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.PrintTwoPagesPerSheet(out status);
            NotifyStatus(status);
        }

        public void OnExportPdfClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.ExportQuickPdf(null, out status);
            NotifyStatus(status);
        }

        #endregion

        #region Callbacks Nhóm Tiện Ích & Xuất Bản PDF

        public void OnExportAccountingPdfAClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.ExportAccountingPdfA(null, out status);
            NotifyStatus(status);
        }

        public void OnExportQuickPdfClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.ExportQuickPdf(null, out status);
            NotifyStatus(status);
        }

        public void OnExportSelectionPdfClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.ExportSelectionToPdf(null, out status);
            NotifyStatus(status);
        }

        public void OnExportCurrentPagePdfClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.ExportCurrentPageToPdf(null, out status);
            NotifyStatus(status);
        }

        public void OnExportOptimizedPdfClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            PdfOptions opts = new PdfOptions
            {
                OptimizeForMinimumSize = true,
                OpenAfterExport = true
            };
            string status;
            _wordService.ExportCustomPdf(opts, null, out status);
            NotifyStatus(status);
        }

        public void OnWatermarkThanhToanClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.AddAccountingWatermark("ĐÃ THANH TOÁN", out status);
            NotifyStatus(status);
        }

        public void OnWatermarkKiemToanClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.AddAccountingWatermark("ĐÃ KIỂM TOÁN", out status);
            NotifyStatus(status);
        }

        public void OnWatermarkBanSaoClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.AddAccountingWatermark("BẢN SAO KẾ TOÁN", out status);
            NotifyStatus(status);
        }

        public void OnWatermarkNoiBoClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.AddAccountingWatermark("LƯU HÀNH NỘI BỘ", out status);
            NotifyStatus(status);
        }

        public void OnRemoveWatermarkClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.RemoveAccountingWatermark(out status);
            NotifyStatus(status);
        }

        public void OnOpenPdfPaneClick(Office.IRibbonControl control)
        {
            EnsureTaskPaneCreated();
            if (_customTaskPane != null)
            {
                _customTaskPane.Visible = true;
            }
            if (_taskPaneHost != null && _taskPaneHost.WpfControl != null)
            {
                _taskPaneHost.WpfControl.SelectModule("pdf-utilities");
            }
            if (_ribbon != null)
            {
                try { _ribbon.InvalidateControl("btnToggleTaskPane"); } catch { }
            }
        }

        #endregion

        #region Callbacks Nhóm Nhập Liệu Nhanh

        public void OnSaveSelectionAsSnippetClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            bool ok = _wordService.SaveCurrentSelectionAsSnippet(null, "Cá nhân", null, out status);
            NotifyStatus(status);
            if (ok)
            {
                EnsureTaskPaneCreated();
                if (_customTaskPane != null)
                {
                    _customTaskPane.Visible = true;
                }
                if (_taskPaneHost != null && _taskPaneHost.WpfControl != null)
                {
                    _taskPaneHost.WpfControl.SelectModule("quick-input");
                }
            }
        }

        public void OnInsertSnippetLuatDn(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-luat-dn");
        }

        public void OnInsertSnippetLuatLd(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-luat-ld");
        }

        public void OnInsertSnippetNd30(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-nd30");
        }

        public void OnInsertSnippetKinhGui(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-kinh-gui");
        }

        public void OnInsertSnippetHieuLucQd(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-hieu-luc-qd");
        }

        public void OnInsertSnippetDnTamUng(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-dn-tam-ung");
        }

        public void OnInsertSnippetDnThanhToan(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-dn-thanh-toan");
        }

        public void OnInsertSnippetThongTinNh(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-tt-nh");
        }

        public void OnInsertSnippetDieuKhoanTt(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-dk-tt");
        }

        public void OnInsertSnippetCamDoanKt(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-cam-ket-kt");
        }

        public void OnInsertSnippetBaoMat(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-bao-mat");
        }

        public void OnInsertSnippetBatKhaKhang(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-bat-kha-khang");
        }

        public void OnInsertSnippetTranhChap(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-tranh-chap");
        }

        public void OnInsertSnippetBoxUnchecked(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-box-unchecked");
        }

        public void OnInsertSnippetBoxChecked(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-box-checked");
        }

        public void OnInsertSnippetArrow(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-arrow");
        }

        public void OnInsertSnippetStar(Office.IRibbonControl control)
        {
            InsertSnippetById("snip-star");
        }

        public void OnOpenQuickInputPaneClick(Office.IRibbonControl control)
        {
            EnsureTaskPaneCreated();
            if (_customTaskPane != null)
            {
                _customTaskPane.Visible = true;
            }
            if (_taskPaneHost != null && _taskPaneHost.WpfControl != null)
            {
                _taskPaneHost.WpfControl.SelectModule("quick-input");
            }
            if (_ribbon != null)
            {
                try { _ribbon.InvalidateControl("btnToggleTaskPane"); } catch { }
            }
        }

        private void InsertSnippetById(string snippetId)
        {
            if (_wordService == null || string.IsNullOrEmpty(snippetId)) return;

            QuickSnippet snip = QuickInputService.GetAllSnippets()
                .FirstOrDefault(s => string.Equals(s.Id, snippetId, StringComparison.OrdinalIgnoreCase));
            if (snip != null)
            {
                string status;
                _wordService.InsertTextSnippet(snip.Content, out status);
                NotifyStatus(status);
            }
        }

        #endregion

        #region Callbacks Nhóm Tách Họ - Tên

        public void OnSplitNameTable2Click(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.SplitTableColumn2(out status);
            NotifyStatus(status);
        }

        public void OnSplitNameTable3Click(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.SplitTableColumn3(out status);
            NotifyStatus(status);
        }

        public void OnMergeNameColumnsClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.MergeTableColumns(out status);
            NotifyStatus(status);
        }

        public void OnSplitNameSelectionTextClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.SplitSelectedTextLines("\t", out status);
            NotifyStatus(status);
        }

        public void OnNormalizeNameSelectionClick(Office.IRibbonControl control)
        {
            if (_wordService == null) return;
            string status;
            _wordService.NormalizeSelectedNames(out status);
            NotifyStatus(status);
        }

        public void OnOpenNameSplitPaneClick(Office.IRibbonControl control)
        {
            EnsureTaskPaneCreated();
            if (_customTaskPane != null)
            {
                _customTaskPane.Visible = true;
            }
            if (_taskPaneHost != null && _taskPaneHost.WpfControl != null)
            {
                _taskPaneHost.WpfControl.SelectModule("name-split");
            }
            if (_ribbon != null)
            {
                try { _ribbon.InvalidateControl("btnToggleTaskPane"); } catch { }
            }
        }

        #endregion

        private void ExecuteConvertWithMode(WordInsertMode? overrideInsertMode)
        {
            if (_wordService == null) return;

            NumberConvertOptions options = SettingsService.Current.ToNumberConvertOptions();
            if (overrideInsertMode.HasValue)
            {
                options.InsertMode = overrideInsertMode.Value;
                SettingsService.Current.InsertMode = (int)overrideInsertMode.Value;
                SettingsService.Save();
            }

            string converted;
            string status;
            _wordService.ConvertSelectedNumber(options, out converted, out status);
            NotifyStatus(status);
        }

        private void NotifyStatus(string message)
        {
            if (string.IsNullOrWhiteSpace(message)) return;

            try
            {
                if (_wordApp != null)
                {
                    _wordApp.StatusBar = "CoolWord: " + message;
                }
            }
            catch { }

            try
            {
                if (_taskPaneHost != null && _taskPaneHost.WpfControl != null)
                {
                    _taskPaneHost.WpfControl.UpdateStatus(message);
                }
            }
            catch { }
        }

        #endregion
    }
}
