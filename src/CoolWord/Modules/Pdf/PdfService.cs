using System;
using System.IO;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace CoolWord.Modules.Pdf
{
    /// <summary>
    /// Dịch vụ xử lý xuất bản và thao tác tài liệu PDF chuyên nghiệp cho CoolWord.
    /// Bao gồm xuất PDF chuẩn lưu trữ kế toán & thuế PDF/A (ISO 19005-1), xuất nhanh theo phạm vi,
    /// và hệ thống đóng dấu chìm kế toán (Watermark).
    /// </summary>
    public static class PdfService
    {
        public const string WatermarkShapeName = "CoolWordWatermarkShape";

        /// <summary>
        /// Tạo đường dẫn xuất file PDF mặc định dựa theo file Word đang mở.
        /// </summary>
        public static string ResolveDefaultPdfPath(Word.Document doc, string suffix = null)
        {
            string baseDir = null;
            string baseName = null;

            if (doc != null)
            {
                try
                {
                    string fullName = doc.FullName;
                    if (!string.IsNullOrEmpty(fullName) && File.Exists(fullName))
                    {
                        baseDir = Path.GetDirectoryName(fullName);
                        baseName = Path.GetFileNameWithoutExtension(fullName);
                    }
                }
                catch { }

                if (string.IsNullOrEmpty(baseName))
                {
                    try
                    {
                        baseName = Path.GetFileNameWithoutExtension(doc.Name);
                    }
                    catch { }
                }
            }

            if (string.IsNullOrEmpty(baseDir))
            {
                baseDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            }

            if (string.IsNullOrEmpty(baseName) || baseName.StartsWith("Document", StringComparison.OrdinalIgnoreCase))
            {
                baseName = "CoolWord_Document_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
            }

            string append = string.IsNullOrEmpty(suffix) ? string.Empty : "_" + suffix.Trim();
            string finalName = baseName + append + ".pdf";
            return Path.Combine(baseDir, finalName);
        }

        /// <summary>
        /// Xuất file PDF với các tùy chọn đầy đủ.
        /// </summary>
        public static bool ExportPdf(Word.Application app, PdfOptions options, string targetPath, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            if (options == null)
            {
                options = new PdfOptions();
            }

            try
            {
                Word.Document doc = app.ActiveDocument;
                string destFile = targetPath;
                if (string.IsNullOrWhiteSpace(destFile))
                {
                    string suffix = options.UsePdfA ? "PDFA" : null;
                    destFile = ResolveDefaultPdfPath(doc, suffix);
                }

                string dir = Path.GetDirectoryName(destFile);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                // Đóng dấu chìm tạm thời nếu có yêu cầu trong options
                bool appliedTempWatermark = false;
                if (!string.IsNullOrWhiteSpace(options.WatermarkText))
                {
                    appliedTempWatermark = AddAccountingWatermarkInternal(doc, options.WatermarkText);
                }

                Word.WdExportRange exportRange;
                switch (options.RangeMode)
                {
                    case PdfExportRange.CurrentPage:
                        exportRange = Word.WdExportRange.wdExportCurrentPage;
                        break;
                    case PdfExportRange.Selection:
                        exportRange = Word.WdExportRange.wdExportSelection;
                        break;
                    default:
                        exportRange = Word.WdExportRange.wdExportAllDocument;
                        break;
                }

                Word.WdExportOptimizeFor optimizeFor = options.OptimizeForMinimumSize
                    ? Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen
                    : Word.WdExportOptimizeFor.wdExportOptimizeForPrint;

                Word.WdExportCreateBookmarks bookmarks = options.CreateBookmarksFromHeadings
                    ? Word.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks
                    : Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks;

                try
                {
                    doc.ExportAsFixedFormat(
                        OutputFileName: destFile,
                        ExportFormat: Word.WdExportFormat.wdExportFormatPDF,
                        OpenAfterExport: options.OpenAfterExport,
                        OptimizeFor: optimizeFor,
                        Range: exportRange,
                        Item: Word.WdExportItem.wdExportDocumentContent,
                        IncludeDocProps: true,
                        KeepIRM: true,
                        CreateBookmarks: bookmarks,
                        DocStructureTags: true,
                        BitmapMissingFonts: true,
                        UseISO19005_1: options.UsePdfA
                    );
                }
                finally
                {
                    // Nếu là đóng dấu chìm chỉ cho lần xuất này, gỡ ra để không sửa đổi tài liệu gốc
                    if (appliedTempWatermark && !string.IsNullOrWhiteSpace(options.WatermarkText))
                    {
                        RemoveAccountingWatermarkInternal(doc);
                    }
                }

                string stdLabel = options.UsePdfA ? "chuẩn lưu trữ PDF/A (ISO 19005-1)" : "chuẩn in ấn";
                statusMessage = string.Format("Đã xuất PDF {0} thành công: {1}", stdLabel, Path.GetFileName(destFile));
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi xuất PDF: " + ex.Message;
                return false;
            }
        }

        #region Tính năng Kế toán: Xuất Báo Cáo & Sổ Kế Toán Chuẩn Lưu Trữ PDF/A (ISO 19005-1)

        /// <summary>
        /// NGHIỆP VỤ KẾ TOÁN: Xuất Báo Cáo Tài Chính / Sổ Kế Toán chuẩn lưu trữ điện tử PDF/A (ISO 19005-1).
        /// Đảm bảo tính pháp lý theo Luật Kế toán và quy định của Tổng cục Thuế về lưu trữ chứng từ điện tử lâu dài (10-20 năm),
        /// tự động nhúng đầy đủ bảng mã ký tự, phông chữ và siêu dữ liệu toàn vẹn.
        /// </summary>
        public static bool ExportAccountingPdfA(Word.Application app, string targetPath, out string statusMessage)
        {
            PdfOptions opts = PdfOptions.CreateAccountingPdfAStandard();
            return ExportPdf(app, opts, targetPath, out statusMessage);
        }

        #endregion

        #region Các Tác Vụ Xuất Nhanh Theo Phạm Vi

        /// <summary>
        /// Chỉ xuất vùng văn bản hoặc bảng biểu đang bôi đen ra file PDF riêng biệt.
        /// </summary>
        public static bool ExportSelectionToPdf(Word.Application app, string targetPath, out string statusMessage)
        {
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            Word.Selection sel = app.Selection;
            if (sel == null || sel.Range == null || sel.Range.Start == sel.Range.End)
            {
                statusMessage = "Vui lòng bôi đen đoạn văn bản hoặc bảng biểu muốn xuất PDF.";
                return false;
            }

            PdfOptions opts = new PdfOptions
            {
                RangeMode = PdfExportRange.Selection,
                OpenAfterExport = true
            };
            return ExportPdf(app, opts, targetPath, out statusMessage);
        }

        /// <summary>
        /// Chỉ xuất trang hiện tại (nơi con trỏ chuột đang đứng) ra file PDF riêng biệt.
        /// </summary>
        public static bool ExportCurrentPageToPdf(Word.Application app, string targetPath, out string statusMessage)
        {
            PdfOptions opts = new PdfOptions
            {
                RangeMode = PdfExportRange.CurrentPage,
                OpenAfterExport = true
            };
            return ExportPdf(app, opts, targetPath, out statusMessage);
        }

        #endregion

        #region Hệ Thống Đóng Dấu Chìm Kế Toán (Watermark)

        /// <summary>
        /// Đóng dấu chìm kế toán / văn bản (Ví dụ: "ĐÃ KIỂM TOÁN", "BẢN SAO KẾ TOÁN", "ĐÃ THANH TOÁN", "LƯU HÀNH NỘI BỘ").
        /// Dấu chìm được chèn vào lớp Header của mọi Section, xoay 315 độ, màu xám nhạt trong suốt, không ảnh hưởng văn bản gốc.
        /// </summary>
        public static bool AddAccountingWatermark(Word.Application app, string watermarkText, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(watermarkText))
            {
                statusMessage = "Nội dung dấu chìm không được để trống.";
                return false;
            }

            try
            {
                Word.Document doc = app.ActiveDocument;
                bool ok = AddAccountingWatermarkInternal(doc, watermarkText.Trim());
                if (ok)
                {
                    statusMessage = string.Format("Đã đóng dấu chìm kế toán \"{0}\" cho toàn bộ văn bản.", watermarkText.Trim());
                    return true;
                }
                statusMessage = "Không thể chèn dấu chìm vào tài liệu.";
                return false;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi chèn dấu chìm: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Gỡ bỏ toàn bộ dấu chìm đã đóng trên văn bản Word.
        /// </summary>
        public static bool RemoveAccountingWatermark(Word.Application app, out string statusMessage)
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
                int removed = RemoveAccountingWatermarkInternal(doc);
                statusMessage = string.Format("Đã gỡ bỏ {0} dấu chìm khỏi tài liệu.", removed);
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi gỡ dấu chìm: " + ex.Message;
                return false;
            }
        }

        private static bool AddAccountingWatermarkInternal(Word.Document doc, string watermarkText)
        {
            if (doc == null || string.IsNullOrWhiteSpace(watermarkText)) return false;

            // Xóa dấu chìm cũ nếu có
            RemoveAccountingWatermarkInternal(doc);

            try
            {
                foreach (Word.Section section in doc.Sections)
                {
                    Word.HeaderFooter header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    Word.Shape shape = header.Shapes.AddTextEffect(
                        PresetTextEffect: Office.MsoPresetTextEffect.msoTextEffect1,
                        Text: watermarkText,
                        FontName: "Times New Roman",
                        FontSize: 54.0f,
                        FontBold: Office.MsoTriState.msoTrue,
                        FontItalic: Office.MsoTriState.msoFalse,
                        Left: 0f,
                        Top: 0f
                    );

                    shape.Name = WatermarkShapeName;
                    shape.Rotation = 315.0f;
                    shape.Fill.Visible = Office.MsoTriState.msoTrue;
                    shape.Fill.Solid();
                    shape.Fill.ForeColor.RGB = 0x00C0C0C0; // Màu xám nhạt (BGR: 192, 192, 192)
                    shape.Fill.Transparency = 0.55f;
                    shape.Line.Visible = Office.MsoTriState.msoFalse;

                    shape.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    shape.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                    shape.Left = (float)Word.WdShapePosition.wdShapeCenter;
                    shape.Top = (float)Word.WdShapePosition.wdShapeCenter;
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static int RemoveAccountingWatermarkInternal(Word.Document doc)
        {
            if (doc == null) return 0;
            int count = 0;

            try
            {
                foreach (Word.Section section in doc.Sections)
                {
                    foreach (Word.HeaderFooter header in section.Headers)
                    {
                        if (header == null) continue;
                        try
                        {
                            for (int i = header.Shapes.Count; i >= 1; i--)
                            {
                                Word.Shape shape = header.Shapes[i];
                                if (shape != null && (shape.Name == WatermarkShapeName || (shape.Name != null && shape.Name.StartsWith("CoolWordWatermark"))))
                                {
                                    shape.Delete();
                                    count++;
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }

            return count;
        }

        #endregion
    }
}
