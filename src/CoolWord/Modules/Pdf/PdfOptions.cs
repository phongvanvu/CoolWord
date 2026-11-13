using System;

namespace CoolWord.Modules.Pdf
{
    public enum PdfExportRange
    {
        AllDocument = 0,
        CurrentPage = 1,
        Selection = 2
    }

    /// <summary>
    /// Các tùy chọn xuất tài liệu sang định dạng PDF.
    /// Hỗ trợ chuẩn ISO 19005-1 (PDF/A) phục vụ lưu trữ kế toán, quyết toán thuế và pháp lý lâu dài.
    /// </summary>
    public class PdfOptions
    {
        /// <summary>
        /// Chuẩn lưu trữ PDF/A (ISO 19005-1).
        /// Bắt buộc cho việc lưu trữ hồ sơ kế toán, báo cáo tài chính, chứng từ điện tử theo quy định kế toán & thuế.
        /// Khi bật, Word sẽ nhúng toàn bộ phông chữ và siêu dữ liệu chuẩn hóa vào file PDF.
        /// </summary>
        public bool UsePdfA { get; set; }

        /// <summary>
        /// Tự động mở file PDF sau khi xuất thành công.
        /// </summary>
        public bool OpenAfterExport { get; set; }

        /// <summary>
        /// Tối ưu hóa dung lượng (OnScreen/Web) để gửi email nhẹ hơn thay vì tối ưu in ấn.
        /// </summary>
        public bool OptimizeForMinimumSize { get; set; }

        /// <summary>
        /// Tạo Bookmarks từ các Heading trong văn bản Word.
        /// </summary>
        public bool CreateBookmarksFromHeadings { get; set; }

        /// <summary>
        /// Phạm vi xuất PDF (Toàn bộ, Trang hiện tại, hoặc Vùng chọn).
        /// </summary>
        public PdfExportRange RangeMode { get; set; }

        /// <summary>
        /// Dấu chìm kế toán / văn thư (Watermark) đính kèm khi xuất (nếu có).
        /// Ví dụ: "ĐÃ KIỂM TOÁN", "BẢN SAO KẾ TOÁN", "ĐÃ THANH TOÁN", "LƯU HÀNH NỘI BỘ", "MẬT".
        /// </summary>
        public string WatermarkText { get; set; }

        public PdfOptions()
        {
            UsePdfA = false;
            OpenAfterExport = true;
            OptimizeForMinimumSize = false;
            CreateBookmarksFromHeadings = true;
            RangeMode = PdfExportRange.AllDocument;
            WatermarkText = string.Empty;
        }

        public static PdfOptions CreateAccountingPdfAStandard()
        {
            return new PdfOptions
            {
                UsePdfA = true,
                OpenAfterExport = true,
                OptimizeForMinimumSize = false,
                CreateBookmarksFromHeadings = true,
                RangeMode = PdfExportRange.AllDocument
            };
        }
    }
}
