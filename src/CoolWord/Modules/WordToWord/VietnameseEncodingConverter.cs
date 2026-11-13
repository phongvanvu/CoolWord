using System;
using System.Collections.Generic;
using System.Text;

namespace CoolWord.Modules.WordToWord
{
    /// <summary>
    /// Bộ máy chuyển đổi bảng mã tiếng Việt TCVN3 (ABC), VNI-Windows sang Unicode dựng sẵn.
    /// Hoạt động 100% offline, nhanh và chính xác cho các tài liệu Word cũ.
    /// </summary>
    public static class VietnameseEncodingConverter
    {
        #region TCVN3 (ABC) Character Map

        // Bảng ký tự TCVN3 (ABC) chữ thường và chữ hoa
        private static readonly Dictionary<char, string> Tcvn3ToUnicodeMap = new Dictionary<char, string>
        {
            // a, ă, â
            { '\u00B8', "á" }, { '\u00B5', "à" }, { '\u00B6', "ả" }, { '\u00B7', "ã" }, { '\u00B9', "ạ" },
            { '\u00A8', "ă" },
            { '\u00BE', "ắ" }, { '\u00BB', "ằ" }, { '\u00BC', "ẳ" }, { '\u00BD', "ẵ" }, { '\u00C6', "ặ" },
            { '\u00A9', "â" },
            { '\u00CA', "ấ" }, { '\u00C7', "ầ" }, { '\u00C8', "ẩ" }, { '\u00C9', "ẫ" }, { '\u00CB', "ậ" },

            // e, ê
            { '\u00D0', "é" }, { '\u00CC', "è" }, { '\u00CE', "ẻ" }, { '\u00CF', "ẽ" }, { '\u00D1', "ẹ" },
            { '\u00AA', "ê" },
            { '\u00D5', "ế" }, { '\u00D2', "ề" }, { '\u00D3', "ể" }, { '\u00D4', "ễ" }, { '\u00D6', "ệ" },

            // i
            { '\u00DD', "í" }, { '\u00D7', "ì" }, { '\u00D8', "ỉ" }, { '\u00DC', "ĩ" }, { '\u00DE', "ị" },

            // o, ô, ơ
            { '\u00E3', "ó" }, { '\u00DF', "ò" }, { '\u00E1', "ỏ" }, { '\u00E2', "õ" }, { '\u00E4', "ọ" },
            { '\u00AB', "ô" },
            { '\u00E9', "ố" }, { '\u00E5', "ồ" }, { '\u00E7', "ổ" }, { '\u00E8', "ỗ" }, { '\u00EA', "ộ" },
            { '\u00AC', "ơ" },
            { '\u00EE', "ớ" }, { '\u00EB', "ờ" }, { '\u00EC', "ở" }, { '\u00ED', "ỡ" }, { '\u00EF', "ợ" },

            // u, ư
            { '\u00F3', "ú" }, { '\u00F0', "ù" }, { '\u00F1', "ủ" }, { '\u00F2', "ũ" }, { '\u00F4', "ụ" },
            { '\u00AD', "ư" },
            { '\u00F8', "ứ" }, { '\u00F5', "ừ" }, { '\u00F6', "ử" }, { '\u00F7', "ữ" }, { '\u00F9', "ự" },

            // y
            { '\u00FD', "ý" }, { '\u00FA', "ỳ" }, { '\u00FB', "ỷ" }, { '\u00FC', "ỹ" }, { '\u00FE', "ỵ" },

            // đ, Đ
            { '\u00AE', "đ" }, { '\u00A7', "Đ" },

            // Chữ hoa gốc TCVN3
            { '\u00A1', "Ă" }, { '\u00A2', "Â" }, { '\u00A3', "Ê" }, { '\u00A4', "Ô" }, { '\u00A5', "Ơ" }, { '\u00A6', "Ư" }
        };

        #endregion

        #region VNI-Windows String Replacements

        // Bảng ký tự VNI-Windows (chuỗi 2 byte và ký tự đặc biệt)
        private static readonly string[] VniSources = new string[]
        {
            "aù", "aø", "aû", "aõ", "aï",
            "aé", "aè", "aú", "aü", "aë", "aê",
            "aá", "aà", "aå", "aã", "aä", "aâ",
            "eù", "eø", "eû", "eõ", "eï",
            "eá", "eà", "eå", "eã", "eä", "eâ",
            "où", "oø", "oû", "oõ", "oï",
            "oá", "oà", "oå", "oã", "oä", "oâ",
            "ôù", "ôø", "ôû", "ôõ", "ôï", "ô",
            "uù", "uø", "uû", "uõ", "uï",
            "öù", "öø", "öû", "öõ", "öï", "ö",
            "yù", "yø", "yû", "yõ", "î",
            "í", "ì", "æ", "ó", "ò",
            "ñ", "Ñ",
            // Chữ hoa VNI
            "AÙ", "AØ", "AÛ", "AÕ", "AÏ",
            "AÉ", "AÈ", "AÚ", "AÜ", "AË", "AÊ",
            "AÁ", "AÀ", "AÅ", "AÃ", "AÄ", "AÂ",
            "EÙ", "EØ", "EÛ", "EÕ", "EÏ",
            "EÁ", "EÀ", "EÅ", "EÃ", "EÄ", "EÂ",
            "OÙ", "OØ", "OÛ", "OÕ", "OÏ",
            "OÁ", "OÀ", "OÅ", "OÃ", "OÄ", "OÂ",
            "ÔÙ", "ÔØ", "ÔÛ", "ÔÕ", "ÔÏ", "Ô",
            "UÙ", "UØ", "UÛ", "UÕ", "UÏ",
            "ÖÙ", "ÖØ", "ÖÛ", "ÖÕ", "ÖÏ", "Ö",
            "YÙ", "YØ", "YÛ", "YÕ", "Î",
            "Í", "Ì", "Æ", "Ó", "Ò"
        };

        private static readonly string[] VniTargets = new string[]
        {
            "á", "à", "ả", "ã", "ạ",
            "ắ", "ằ", "ẳ", "ẵ", "ặ", "ă",
            "ấ", "ầ", "ẩ", "ẫ", "ậ", "â",
            "é", "è", "ẻ", "ẽ", "ẹ",
            "ế", "ề", "ể", "ễ", "ệ", "ê",
            "ó", "ò", "ỏ", "õ", "ọ",
            "ố", "ồ", "ổ", "ỗ", "ộ", "ô",
            "ớ", "ờ", "ở", "ỡ", "ợ", "ơ",
            "ú", "ù", "ủ", "ũ", "ụ",
            "ứ", "ừ", "ử", "ữ", "ự", "ư",
            "ý", "ỳ", "ỷ", "ỹ", "ỵ",
            "í", "ì", "ỉ", "ĩ", "ị",
            "đ", "Đ",
            // Chữ hoa tương ứng
            "Á", "À", "Ả", "Ã", "Ạ",
            "Ắ", "Ằ", "Ẳ", "Ẵ", "Ặ", "Ă",
            "Ấ", "Ầ", "Ẩ", "Ẫ", "Ậ", "Â",
            "É", "È", "Ẻ", "Ẽ", "Ẹ",
            "Ế", "Ề", "Ể", "Ễ", "Ệ", "Ê",
            "Ó", "Ò", "Ỏ", "Õ", "Ọ",
            "Ố", "Ồ", "Ổ", "Ỗ", "Ộ", "Ô",
            "Ớ", "Ờ", "Ở", "Ỡ", "Ợ", "Ơ",
            "Ú", "Ù", "Ủ", "Ũ", "Ụ",
            "Ứ", "Ừ", "Ử", "Ữ", "Ự", "Ư",
            "Ý", "Ỳ", "Ỷ", "Ỹ", "Ỵ",
            "Í", "Ì", "Ỉ", "Ĩ", "Ị"
        };

        #endregion

        /// <summary>
        /// Chuyển đổi chuỗi văn bản từ mã TCVN3 (ABC) sang Unicode chuẩn.
        /// </summary>
        public static string ConvertTcvn3ToUnicode(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;

            StringBuilder sb = new StringBuilder(input.Length);
            foreach (char c in input)
            {
                string u;
                if (Tcvn3ToUnicodeMap.TryGetValue(c, out u))
                {
                    sb.Append(u);
                }
                else
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Chuyển đổi chuỗi văn bản từ mã VNI-Windows sang Unicode chuẩn.
        /// </summary>
        public static string ConvertVniToUnicode(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;

            string res = input;
            for (int i = 0; i < VniSources.Length; i++)
            {
                res = res.Replace(VniSources[i], VniTargets[i]);
            }
            return res;
        }

        /// <summary>
        /// Kiểm tra tên phông chữ có phải là phông mã TCVN3 (.VnTime, .VnArial, ...) không.
        /// </summary>
        public static bool IsTcvn3Font(string fontName)
        {
            if (string.IsNullOrWhiteSpace(fontName)) return false;
            string fn = fontName.Trim();
            return fn.StartsWith(".Vn", StringComparison.OrdinalIgnoreCase) ||
                   fn.StartsWith(".vn", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Kiểm tra tên phông chữ có phải là phông mã VNI (VNI-Times, VNI-Helve, ...) không.
        /// </summary>
        public static bool IsVniFont(string fontName)
        {
            if (string.IsNullOrWhiteSpace(fontName)) return false;
            string fn = fontName.Trim();
            return fn.StartsWith("VNI-", StringComparison.OrdinalIgnoreCase) ||
                   fn.StartsWith("Vni-", StringComparison.OrdinalIgnoreCase);
        }
    }
}
