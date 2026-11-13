using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace CoolWord.Modules.ImageTools
{
    /// <summary>
    /// Dịch vụ xử lý hình ảnh chuyên nghiệp trong Microsoft Word:
    /// Trích xuất toàn bộ ảnh chất lượng gốc, nén ảnh, chèn ảnh hàng loạt và căn chỉnh.
    /// Nghiệp vụ kế toán / hành chính: Chèn con dấu đỏ & chữ ký số tự động tách nền trong suốt.
    /// </summary>
    public static class ImageToolsService
    {
        private static readonly object Missing = Type.Missing;

        #region Public Methods

        /// <summary>
        /// Trích xuất toàn bộ hình ảnh trong tài liệu Word ra một thư mục riêng (định dạng PNG gốc).
        /// </summary>
        public static bool ExportAllImages(Word.Application wordApp, string outputDir, out string statusMessage)
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
                int inlineCount = doc.InlineShapes.Count;
                int shapeCount = doc.Shapes.Count;

                if (inlineCount == 0 && shapeCount == 0)
                {
                    statusMessage = "Tài liệu Word không chứa hình ảnh nào.";
                    return false;
                }

                string dir = outputDir;
                if (string.IsNullOrEmpty(dir))
                {
                    string fullName = string.Empty;
                    try { fullName = doc.FullName; } catch { }

                    dir = string.IsNullOrEmpty(fullName) || !File.Exists(fullName)
                        ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "CoolWord_Images")
                        : Path.Combine(Path.GetDirectoryName(fullName), Path.GetFileNameWithoutExtension(fullName) + "_Images");
                }

                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                int exportedCount = 0;

                // 1. Duyệt InlineShapes (Ảnh chèn inline trong dòng văn bản)
                for (int i = 1; i <= inlineCount; i++)
                {
                    try
                    {
                        Word.InlineShape shape = doc.InlineShapes[i];
                        if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture ||
                            shape.Type == Word.WdInlineShapeType.wdInlineShapeLinkedPicture)
                        {
                            shape.Range.CopyAsPicture();
                            if (Clipboard.ContainsImage())
                            {
                                using (Image img = Clipboard.GetImage())
                                {
                                    if (img != null)
                                    {
                                        string filePath = Path.Combine(dir, string.Format("Anh_{0:000}.png", exportedCount + 1));
                                        img.Save(filePath, ImageFormat.Png);
                                        exportedCount++;
                                    }
                                }
                            }
                        }
                    }
                    catch { }
                }

                // 2. Duyệt Shapes (Ảnh nổi, ảnh trôi nổi trong văn bản)
                for (int j = 1; j <= shapeCount; j++)
                {
                    try
                    {
                        Word.Shape sh = doc.Shapes[j];
                        if (sh.Type == Microsoft.Office.Core.MsoShapeType.msoPicture ||
                            sh.Type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture)
                        {
                            sh.Select(Missing);
                            wordApp.Selection.CopyAsPicture();
                            if (Clipboard.ContainsImage())
                            {
                                using (Image img = Clipboard.GetImage())
                                {
                                    if (img != null)
                                    {
                                        string filePath = Path.Combine(dir, string.Format("Anh_{0:000}.png", exportedCount + 1));
                                        img.Save(filePath, ImageFormat.Png);
                                        exportedCount++;
                                    }
                                }
                            }
                        }
                    }
                    catch { }
                }

                statusMessage = string.Format("Đã trích xuất thành công {0} hình ảnh ra thư mục: {1}", exportedCount, Path.GetFileName(dir));
                return exportedCount > 0;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi trích xuất ảnh: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Nghiệp Vụ Kế Toán & Hành Chính: Chèn Con Dấu Đỏ & Chữ Ký Số vào tài liệu.
        /// Tự động tách trong suốt màu nền trắng để con dấu đè lên chữ ký chuẩn văn bản hành chính Việt Nam.
        /// </summary>
        public static bool InsertOfficialStampOrSignature(
            Word.Application wordApp,
            string stampImagePath,
            bool inFrontOfText,
            float widthPoints,
            out string statusMessage)
        {
            statusMessage = string.Empty;
            if (wordApp == null || wordApp.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            string imgPath = stampImagePath;
            if (string.IsNullOrEmpty(imgPath) || !File.Exists(imgPath))
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Title = "Chọn file ảnh Con dấu đỏ hoặc Chữ ký số (.png, .jpg)";
                    ofd.Filter = "Tệp hình ảnh (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg|Tất cả tệp (*.*)|*.*";
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        imgPath = ofd.FileName;
                    }
                }
            }

            if (string.IsNullOrEmpty(imgPath) || !File.Exists(imgPath))
            {
                statusMessage = "Chưa chọn file ảnh con dấu / chữ ký.";
                return false;
            }

            try
            {
                Word.Document doc = wordApp.ActiveDocument;
                Word.Selection sel = wordApp.Selection;
                Word.Range targetRange = (sel != null) ? sel.Range : doc.Content;

                float stampWidth = (widthPoints > 0) ? widthPoints : 120f; // Mặc định rộng khoảng 4.2 cm

                Word.Shape shape = doc.Shapes.AddPicture(
                    imgPath,
                    false,
                    true,
                    targetRange.Start,
                    targetRange.End,
                    stampWidth,
                    -1,
                    targetRange);

                // Thiết lập hiển thị nổi hoặc chìm
                shape.WrapFormat.Type = Word.WdWrapType.wdWrapNone;
                if (inFrontOfText)
                {
                    shape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringInFrontOfText);
                }
                else
                {
                    shape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendBehindText);
                }

                // Tự động làm trong suốt màu nền trắng của con dấu để không bị che khuất chữ ký
                try
                {
                    shape.PictureFormat.TransparentBackground = Microsoft.Office.Core.MsoTriState.msoTrue;
                    shape.PictureFormat.TransparencyColor = ColorTranslator.ToOle(Color.White);
                }
                catch { }

                statusMessage = "Đã chèn con dấu đỏ / chữ ký số vào văn bản thành công.";
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi chèn con dấu / chữ ký: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Chèn hàng loạt nhiều hình ảnh vào tài liệu Word theo kích thước chuẩn thống nhất.
        /// </summary>
        public static bool BatchInsertImages(
            Word.Application wordApp,
            string[] imageFiles,
            float widthPoints,
            out string statusMessage)
        {
            statusMessage = string.Empty;
            if (wordApp == null || wordApp.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            string[] files = imageFiles;
            if (files == null || files.Length == 0)
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Title = "Chọn các file ảnh cần chèn vào tài liệu";
                    ofd.Filter = "Tệp hình ảnh (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp|Tất cả tệp (*.*)|*.*";
                    ofd.Multiselect = true;
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        files = ofd.FileNames;
                    }
                }
            }

            if (files == null || files.Length == 0)
            {
                statusMessage = "Chưa chọn file ảnh để chèn.";
                return false;
            }

            try
            {
                Word.Document doc = wordApp.ActiveDocument;
                Word.Selection sel = wordApp.Selection;
                Word.Range insertRange = sel.Range;

                int inserted = 0;
                float targetWidth = widthPoints > 0 ? widthPoints : 350f;

                foreach (string file in files)
                {
                    if (!File.Exists(file)) continue;

                    Word.InlineShape shape = insertRange.InlineShapes.AddPicture(file, false, true, Missing);
                    if (shape != null)
                    {
                        // Giữ nguyên tỷ lệ khung hình
                        if (shape.Width > targetWidth)
                        {
                            float ratio = shape.Height / shape.Width;
                            shape.Width = targetWidth;
                            shape.Height = targetWidth * ratio;
                        }

                        shape.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                        insertRange.Collapse(ref collapseEnd);
                        insertRange.InsertParagraphAfter();
                        insertRange.Collapse(ref collapseEnd);
                        inserted++;
                    }
                }

                statusMessage = string.Format("Đã chèn thành công {0} hình ảnh vào tài liệu.", inserted);
                return inserted > 0;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi chèn ảnh hàng loạt: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Căn giữa tất cả các hình ảnh Inline trong toàn bộ tài liệu Word.
        /// </summary>
        public static bool CenterAllImages(Word.Application wordApp, out string statusMessage)
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
                int count = 0;
                foreach (Word.InlineShape shape in doc.InlineShapes)
                {
                    if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture ||
                        shape.Type == Word.WdInlineShapeType.wdInlineShapeLinkedPicture)
                    {
                        shape.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }

                statusMessage = string.Format("Đã căn giữa {0} hình ảnh trong tài liệu.", count);
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi căn giữa ảnh: " + ex.Message;
                return false;
            }
        }

        #endregion
    }
}
