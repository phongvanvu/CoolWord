using System;
using System.Text;
using System.Text.RegularExpressions;

namespace CoolWord.Modules.NumberToText
{
    public class ParsedNumberInfo
    {
        public bool IsValid { get; set; }
        public bool IsNegative { get; set; }
        public string IntegerDigits { get; set; }
        public string FractionDigits { get; set; }
        public string MatchedText { get; set; }
        public int MatchStartIndex { get; set; }
        public int MatchLength { get; set; }
        public NumberReadingMode? DetectedCurrency { get; set; }

        public bool HasFraction
        {
            get { return !string.IsNullOrEmpty(FractionDigits) && !IsAllZeros(FractionDigits); }
        }

        private static bool IsAllZeros(string s)
        {
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] != '0') return false;
            }
            return true;
        }
    }

    public static class NumberExtractor
    {
        // Regex tìm cụm số (bao gồm dấu âm, dấu chấm/phẩy/khoảng trắng phân cách, và ký hiệu tiền tệ đi kèm)
        private static readonly Regex NumberCandidateRegex = new Regex(
            @"(?<prefix>\$|USD|VNĐ|VND)?\s*(?<sign>[-+])?\s*(?<num>\d(?:[\d.,\s]*\d)?)\s*(?<suffix>đồng|vnđ|vnd|đ|usd|\$)?",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static ParsedNumberInfo Extract(string rawInput, NumberReadingMode currentMode = NumberReadingMode.CurrencyVnd)
        {
            if (string.IsNullOrWhiteSpace(rawInput))
            {
                return new ParsedNumberInfo { IsValid = false };
            }

            Match match = NumberCandidateRegex.Match(rawInput);
            if (!match.Success)
            {
                return new ParsedNumberInfo { IsValid = false };
            }

            string prefix = match.Groups["prefix"].Value.Trim();
            string sign = match.Groups["sign"].Value.Trim();
            string numRaw = match.Groups["num"].Value.Trim();
            string suffix = match.Groups["suffix"].Value.Trim();

            NumberReadingMode? detectedCurrency = null;
            string currencyToken = (prefix + " " + suffix).ToLowerInvariant();
            if (currencyToken.Contains("usd") || currencyToken.Contains("$"))
            {
                detectedCurrency = NumberReadingMode.CurrencyUsd;
            }
            else if (currencyToken.Contains("vnđ") || currencyToken.Contains("vnd") || currencyToken.Contains("đồng") || currencyToken.Contains("đ"))
            {
                detectedCurrency = NumberReadingMode.CurrencyVnd;
            }

            bool isNegative = sign == "-";
            NumberReadingMode effectiveMode = detectedCurrency ?? currentMode;

            string intPart;
            string fracPart;
            if (!TrySplitIntegerAndFraction(numRaw, effectiveMode, out intPart, out fracPart))
            {
                return new ParsedNumberInfo { IsValid = false };
            }

            intPart = TrimLeadingZeros(intPart);
            fracPart = TrimTrailingZeros(fracPart);

            return new ParsedNumberInfo
            {
                IsValid = true,
                IsNegative = isNegative && (intPart != "0" || !string.IsNullOrEmpty(fracPart)),
                IntegerDigits = intPart,
                FractionDigits = fracPart,
                MatchedText = match.Value.Trim(),
                MatchStartIndex = match.Index,
                MatchLength = match.Length,
                DetectedCurrency = detectedCurrency
            };
        }

        private static bool TrySplitIntegerAndFraction(string numRaw, NumberReadingMode mode, out string intDigits, out string fracDigits)
        {
            intDigits = string.Empty;
            fracDigits = string.Empty;

            // Xóa khoảng trắng phân cách nghìn (VD: "1 250 000" -> "1250000")
            string cleaned = Regex.Replace(numRaw, @"\s+", "");
            if (string.IsNullOrEmpty(cleaned))
                return false;

            int dotCount = CountChar(cleaned, '.');
            int commaCount = CountChar(cleaned, ',');

            if (dotCount == 0 && commaCount == 0)
            {
                intDigits = KeepOnlyDigits(cleaned);
                return !string.IsNullOrEmpty(intDigits);
            }

            if (dotCount > 0 && commaCount > 0)
            {
                // Có cả dấu chấm và dấu phẩy: dấu nào xuất hiện sau cùng là dấu thập phân
                int lastDot = cleaned.LastIndexOf('.');
                int lastComma = cleaned.LastIndexOf(',');
                int decIndex = Math.Max(lastDot, lastComma);

                string left = cleaned.Substring(0, decIndex);
                string right = cleaned.Substring(decIndex + 1);

                intDigits = KeepOnlyDigits(left);
                fracDigits = KeepOnlyDigits(right);
                return !string.IsNullOrEmpty(intDigits);
            }

            char sep = dotCount > 0 ? '.' : ',';
            int sepCount = dotCount > 0 ? dotCount : commaCount;

            if (sepCount > 1)
            {
                // Nhiều dấu phân cách giống nhau (VD: 1.250.000 hoặc 1,250,000) -> chắc chắn là phân cách hàng nghìn
                intDigits = KeepOnlyDigits(cleaned);
                return !string.IsNullOrEmpty(intDigits);
            }

            // Chỉ có đúng 1 dấu chấm hoặc 1 dấu phẩy
            int idx = cleaned.IndexOf(sep);
            string leftPart = cleaned.Substring(0, idx);
            string rightPart = cleaned.Substring(idx + 1);

            bool isThousandSeparator = false;
            if (rightPart.Length == 3 && leftPart.Length >= 1 && leftPart.Length <= 3 && leftPart != "0")
            {
                if (sep == '.')
                {
                    // Trong tiếng Việt, 150.000 hay 1.250 hầu như luôn là dấu phân cách hàng nghìn
                    isThousandSeparator = true;
                }
                else if (sep == ',')
                {
                    // Với dấu phẩy: nếu đang ở chế độ tiền tệ VNĐ/USD hoặc kết thúc bằng 000 (VD: 150,000) -> phân cách nghìn
                    if (mode == NumberReadingMode.CurrencyVnd || mode == NumberReadingMode.CurrencyUsd || rightPart == "000")
                    {
                        isThousandSeparator = true;
                    }
                }
            }

            if (isThousandSeparator)
            {
                intDigits = KeepOnlyDigits(leftPart + rightPart);
                fracDigits = string.Empty;
            }
            else
            {
                intDigits = KeepOnlyDigits(leftPart);
                fracDigits = KeepOnlyDigits(rightPart);
            }

            return !string.IsNullOrEmpty(intDigits);
        }

        private static int CountChar(string s, char c)
        {
            int count = 0;
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == c) count++;
            }
            return count;
        }

        private static string KeepOnlyDigits(string s)
        {
            StringBuilder sb = new StringBuilder(s.Length);
            for (int i = 0; i < s.Length; i++)
            {
                if (char.IsDigit(s[i]))
                {
                    sb.Append(s[i]);
                }
            }
            return sb.ToString();
        }

        private static string TrimLeadingZeros(string s)
        {
            if (string.IsNullOrEmpty(s)) return "0";
            int i = 0;
            while (i < s.Length - 1 && s[i] == '0')
            {
                i++;
            }
            return s.Substring(i);
        }

        private static string TrimTrailingZeros(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            int i = s.Length - 1;
            while (i >= 0 && s[i] == '0')
            {
                i--;
            }
            return i < 0 ? string.Empty : s.Substring(0, i + 1);
        }
    }
}
