using System;

namespace CoolWord.Modules.AdminFormat
{
    public class AdminDocOptions
    {
        // 1. Căn lề trang chuẩn Nghị định 30/2020/NĐ-CP (đơn vị: mm)
        public double TopMarginMm { get; set; }
        public double BottomMarginMm { get; set; }
        public double LeftMarginMm { get; set; }
        public double RightMarginMm { get; set; }

        // 2. Phông chữ & Đoạn văn nội dung
        public string FontName { get; set; }
        public float BodyFontSize { get; set; }
        public double FirstLineIndentCm { get; set; }
        public float SpaceBeforePt { get; set; }
        public float SpaceAfterPt { get; set; }
        public float LineSpacingMultiple { get; set; }
        public bool EnableDecree30PageNumbering { get; set; }

        // 3. Thông tin Khung Đầu Văn Bản (Quốc hiệu - Tiêu ngữ - Cơ quan - Số ký hiệu - Ngày tháng)
        public string ParentOrgName { get; set; }
        public string IssuingOrgName { get; set; }
        public string DocumentNumber { get; set; }
        public string LocationName { get; set; }
        public DateTime IssueDate { get; set; }
        public string DocumentTypeTitle { get; set; }
        public string DocumentSubject { get; set; }

        // 4. Thông tin Khung Cuối Văn Bản (Nơi nhận & Chữ ký)
        public string RecipientsList { get; set; }
        public string AuthorityPrefix { get; set; }
        public string SignerPosition { get; set; }
        public string SignerFullName { get; set; }

        public AdminDocOptions()
        {
            // Mặc định chuẩn Phụ lục I - Nghị định 30/2020/NĐ-CP:
            // Lề trên 20-25mm, Lề dưới 20-25mm, Lề trái 30-35mm, Lề phải 15-20mm
            TopMarginMm = 20.0;
            BottomMarginMm = 20.0;
            LeftMarginMm = 30.0;
            RightMarginMm = 15.0;

            FontName = "Times New Roman";
            BodyFontSize = 14.0f;
            FirstLineIndentCm = 1.27;
            SpaceBeforePt = 6.0f;
            SpaceAfterPt = 0.0f;
            LineSpacingMultiple = 1.15f;
            EnableDecree30PageNumbering = true;

            ParentOrgName = "UBND THÀNH PHỐ HÀ NỘI";
            IssuingOrgName = "SỞ THÔNG TIN VÀ TRUYỀN THÔNG";
            DocumentNumber = "Số:     /TTr-STTTT";
            LocationName = "Hà Nội";
            IssueDate = DateTime.Now;
            DocumentTypeTitle = "TỜ TRÌNH";
            DocumentSubject = "Về việc phê duyệt kế hoạch chuyển đổi số năm " + DateTime.Now.Year;

            RecipientsList = "- Như trên;\r- Ban Giám đốc;\r- Lưu: VT, KH.";
            AuthorityPrefix = "";
            SignerPosition = "GIÁM ĐỐC";
            SignerFullName = "Nguyễn Văn A";
        }
    }
}
