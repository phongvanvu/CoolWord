using System;

namespace CoolWord.Modules.NameSplit
{
    /// <summary>
    /// Đối tượng chứa kết quả phân tích cấu trúc họ và tên tiếng Việt.
    /// </summary>
    public class ParsedVietnameseName
    {
        public string RawInput { get; set; }

        /// <summary>
        /// Họ và tên đầy đủ đã được chuẩn hóa (viết hoa chữ cái đầu, xóa khoảng cách thừa).
        /// </summary>
        public string FullName { get; set; }

        /// <summary>
        /// Họ và tên đệm (dùng khi tách thành 2 cột: Họ đệm & Tên).
        /// Ví dụ: "Nguyễn Văn An" -> "Nguyễn Văn"
        /// </summary>
        public string HoDem { get; set; }

        /// <summary>
        /// Họ (dùng khi tách thành 3 cột: Họ, Đệm & Tên).
        /// Ví dụ: "Nguyễn Văn An" -> "Nguyễn"
        /// </summary>
        public string Ho { get; set; }

        /// <summary>
        /// Tên đệm / tên lót (dùng khi tách thành 3 cột).
        /// Ví dụ: "Trần Thị Bích Ngọc" -> "Thị Bích"
        /// </summary>
        public string Dem { get; set; }

        /// <summary>
        /// Tên chính (từ cuối cùng dùng để sắp xếp danh sách A-Z).
        /// Ví dụ: "Nguyễn Văn An" -> "An"
        /// </summary>
        public string Ten { get; set; }

        public ParsedVietnameseName()
        {
            RawInput = string.Empty;
            FullName = string.Empty;
            HoDem = string.Empty;
            Ho = string.Empty;
            Dem = string.Empty;
            Ten = string.Empty;
        }

        public override string ToString()
        {
            return string.Format("{0} | {1}", HoDem, Ten);
        }
    }
}
