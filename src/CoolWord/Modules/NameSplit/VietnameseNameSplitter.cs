using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace CoolWord.Modules.NameSplit
{
    /// <summary>
    /// Thuật toán phân tích, chuẩn hóa và tách họ tên người Việt Nam
    /// chuẩn âm vị và ngữ pháp tiếng Việt.
    /// </summary>
    public static class VietnameseNameSplitter
    {
        private static readonly HashSet<string> CompoundSurnames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Tôn Thất", "Tôn Nữ", "Công Tằng Tôn Nữ", "Nguyễn Phước", "Nguyễn Phúc",
            "Nguyễn Đình", "Nguyễn Hữu", "Nguyễn Văn", "Nguyễn Thị", "Phan Huy",
            "Đặng Vũ", "Hoàng Kim", "Lê Phạm", "Võ Nguyễn", "Trần Đăng", "Bùi Đình"
        };

        /// <summary>
        /// Chuẩn hóa họ tên: Xóa khoảng trắng thừa, viết hoa chữ cái đầu mỗi từ chuẩn tiếng Việt,
        /// giữ nguyên định dạng các họ tên dân tộc thiểu số có tiền tố đặc biệt (H'Hen, Y Moan, K'...).
        /// </summary>
        public static string NormalizeName(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;

            // Xóa ký tự điều khiển và khoảng cách thừa
            string clean = Regex.Replace(raw.Trim(), @"[\r\n\t\a]+", " ");
            clean = Regex.Replace(clean, @"\s+", " ");

            CultureInfo vi = new CultureInfo("vi-VN");
            string titleCase = vi.TextInfo.ToTitleCase(clean.ToLower(vi));

            // Sửa lại các tiền tố đặc biệt dân tộc thiểu số: H'hen -> H'Hen, K'sor -> K'Sor, Y' -> Y'
            titleCase = Regex.Replace(titleCase, @"\b(H|K|Y|M|N)'([a-zà-ỹ])", m => m.Groups[1].Value + "'" + m.Groups[2].Value.ToUpper(vi), RegexOptions.IgnoreCase);

            return titleCase;
        }

        /// <summary>
        /// Phân tích và tách chuỗi họ và tên thành các thành phần: Họ, Tên đệm, Tên, và Họ đệm.
        /// </summary>
        public static ParsedVietnameseName Split(string raw)
        {
            ParsedVietnameseName result = new ParsedVietnameseName
            {
                RawInput = raw ?? string.Empty
            };

            if (string.IsNullOrWhiteSpace(raw))
            {
                return result;
            }

            string normalized = NormalizeName(raw);
            result.FullName = normalized;

            string[] tokens = normalized.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length == 0)
            {
                return result;
            }

            if (tokens.Length == 1)
            {
                result.Ten = tokens[0];
                result.Ho = string.Empty;
                result.Dem = string.Empty;
                result.HoDem = string.Empty;
                return result;
            }

            if (tokens.Length == 2)
            {
                result.Ho = tokens[0];
                result.Dem = string.Empty;
                result.Ten = tokens[1];
                result.HoDem = tokens[0];
                return result;
            }

            // tokens.Length >= 3:
            // Tên luôn là từ cuối cùng (chuẩn xếp danh sách A-Z của người Việt)
            result.Ten = tokens[tokens.Length - 1];

            // Họ và tên đệm: tất cả các từ ngoại trừ từ cuối cùng
            result.HoDem = string.Join(" ", tokens.Take(tokens.Length - 1));

            // Phân biệt Họ và Tên đệm (3 cột):
            // Kiểm tra trường hợp họ ghép 2 từ (nếu có từ 4 từ trở lên và 2 từ đầu là họ ghép thông dụng)
            if (tokens.Length >= 4)
            {
                string firstTwo = tokens[0] + " " + tokens[1];
                if (CompoundSurnames.Contains(firstTwo))
                {
                    result.Ho = firstTwo;
                    result.Dem = string.Join(" ", tokens.Skip(2).Take(tokens.Length - 3));
                    return result;
                }
            }

            // Chuẩn mặc định: từ đầu tiên là Họ, các từ ở giữa là Tên đệm
            result.Ho = tokens[0];
            result.Dem = string.Join(" ", tokens.Skip(1).Take(tokens.Length - 2));

            return result;
        }

        /// <summary>
        /// Gộp Họ đệm và Tên thành một chuỗi Họ và tên hoàn chỉnh.
        /// </summary>
        public static string Merge(string hoDem, string ten)
        {
            string p1 = NormalizeName(hoDem);
            string p2 = NormalizeName(ten);

            if (string.IsNullOrEmpty(p1)) return p2;
            if (string.IsNullOrEmpty(p2)) return p1;

            return p1 + " " + p2;
        }
    }
}
