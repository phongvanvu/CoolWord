using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace CoolWord.Modules.NumberToText
{
    public static class VietnameseNumberConverter
    {
        private static readonly string[] DigitWords = new[]
        {
            "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"
        };

        /// <summary>
        /// Chuyển đổi chuỗi chứa số (hoặc tiền tệ) sang chữ Tiếng Việt theo tùy chọn.
        /// </summary>
        public static string Convert(string rawInput, NumberConvertOptions options = null)
        {
            if (options == null)
            {
                options = new NumberConvertOptions();
            }

            ParsedNumberInfo parsed = NumberExtractor.Extract(rawInput, options.ReadingMode);
            if (!parsed.IsValid)
            {
                return string.Empty;
            }

            return ConvertParsed(parsed, options);
        }

        /// <summary>
        /// Chuyển đổi đối tượng số đã phân tích sang chữ Tiếng Việt.
        /// </summary>
        public static string ConvertParsed(ParsedNumberInfo parsed, NumberConvertOptions options = null)
        {
            if (parsed == null || !parsed.IsValid)
            {
                return string.Empty;
            }

            if (options == null)
            {
                options = new NumberConvertOptions();
            }

            NumberReadingMode effectiveMode = options.ReadingMode;
            if (options.AutoDetectCurrencyFromSelection && parsed.DetectedCurrency.HasValue)
            {
                effectiveMode = parsed.DetectedCurrency.Value;
            }

            string thousandWord = options.GetThousandWord();
            string zeroTensWord = options.GetZeroTensWord();
            bool useTu = options.UseTuForFour();

            string intWords = ReadIntegerDigits(parsed.IntegerDigits, thousandWord, zeroTensWord, useTu);
            StringBuilder sb = new StringBuilder();

            if (parsed.IsNegative)
            {
                sb.Append("âm ");
            }

            sb.Append(intWords);

            switch (effectiveMode)
            {
                case NumberReadingMode.PureNumber:
                    if (parsed.HasFraction)
                    {
                        sb.Append(" phẩy ");
                        sb.Append(ReadDecimalFraction(parsed.FractionDigits, thousandWord, zeroTensWord, useTu));
                    }
                    break;

                case NumberReadingMode.CurrencyVnd:
                    if (parsed.HasFraction)
                    {
                        sb.Append(" phẩy ");
                        sb.Append(ReadDecimalFraction(parsed.FractionDigits, thousandWord, zeroTensWord, useTu));
                        sb.Append(" đồng");
                    }
                    else
                    {
                        sb.Append(" đồng");
                        if (options.AppendChan)
                        {
                            sb.Append(" chẵn");
                        }
                    }
                    break;

                case NumberReadingMode.CurrencyUsd:
                    sb.Append(" đô la Mỹ");
                    if (parsed.HasFraction)
                    {
                        string centsDigits = NormalizeSubUnitDigits(parsed.FractionDigits);
                        string centsWords = ReadIntegerDigits(centsDigits, thousandWord, zeroTensWord, useTu);
                        sb.Append(" và ");
                        sb.Append(centsWords);
                        sb.Append(" xu");
                    }
                    else if (options.AppendChan)
                    {
                        sb.Append(" chẵn");
                    }
                    break;

                case NumberReadingMode.CustomCurrency:
                    string major = string.IsNullOrWhiteSpace(options.CustomMajorUnit) ? "đồng" : options.CustomMajorUnit.Trim();
                    string minor = string.IsNullOrWhiteSpace(options.CustomMinorUnit) ? "xu" : options.CustomMinorUnit.Trim();

                    if (parsed.HasFraction)
                    {
                        if (!string.IsNullOrEmpty(options.CustomMinorUnit))
                        {
                            string minorDigits = NormalizeSubUnitDigits(parsed.FractionDigits);
                            string minorWords = ReadIntegerDigits(minorDigits, thousandWord, zeroTensWord, useTu);
                            sb.Append(" ");
                            sb.Append(major);
                            sb.Append(" và ");
                            sb.Append(minorWords);
                            sb.Append(" ");
                            sb.Append(minor);
                        }
                        else
                        {
                            sb.Append(" phẩy ");
                            sb.Append(ReadDecimalFraction(parsed.FractionDigits, thousandWord, zeroTensWord, useTu));
                            sb.Append(" ");
                            sb.Append(major);
                        }
                    }
                    else
                    {
                        sb.Append(" ");
                        sb.Append(major);
                        if (options.AppendChan)
                        {
                            sb.Append(" chẵn");
                        }
                    }
                    break;
            }

            string result = NormalizeSpaces(sb.ToString());
            result = ApplyCapitalization(result, options.Capitalization);

            if (options.AppendDotSlash && effectiveMode != NumberReadingMode.PureNumber)
            {
                result += "./.";
            }

            return result;
        }

        /// <summary>
        /// Tạo chuỗi văn bản hoàn chỉnh để thay thế hoặc chèn vào tài liệu Word dựa trên chế độ chèn.
        /// </summary>
        public static string FormatForWordInsertion(string originalSelection, string convertedWords, NumberConvertOptions options)
        {
            if (string.IsNullOrEmpty(convertedWords))
            {
                return originalSelection;
            }

            if (options == null)
            {
                options = new NumberConvertOptions();
            }

            string trimmedOriginal = (originalSelection ?? string.Empty).TrimEnd();

            switch (options.InsertMode)
            {
                case WordInsertMode.ReplaceSelection:
                    return convertedWords;

                case WordInsertMode.AppendInParentheses:
                    return trimmedOriginal + " (" + convertedWords + ")";

                case WordInsertMode.AppendWithLabelInParentheses:
                    return trimmedOriginal + " (Bằng chữ: " + convertedWords + ")";

                case WordInsertMode.AppendOnNewLine:
                    return trimmedOriginal + Environment.NewLine + "Bằng chữ: " + convertedWords;

                default:
                    return convertedWords;
            }
        }

        private static string ReadIntegerDigits(string digits, string thousandWord, string zeroTensWord, bool useTu)
        {
            if (string.IsNullOrEmpty(digits) || digits == "0")
            {
                return "không";
            }

            int rem = digits.Length % 3;
            string padded = rem == 0 ? digits : digits.PadLeft(digits.Length + (3 - rem), '0');
            int groupCount = padded.Length / 3;

            List<string> parts = new List<string>();

            for (int g = 0; g < groupCount; g++)
            {
                int scaleIndex = (groupCount - 1) - g;
                int h = padded[g * 3] - '0';
                int t = padded[g * 3 + 1] - '0';
                int u = padded[g * 3 + 2] - '0';

                bool hasHigherNonZero = g > 0;
                bool isZeroGroup = (h == 0 && t == 0 && u == 0);

                if (!isZeroGroup)
                {
                    string groupWords = ReadThreeDigitGroup(h, t, u, hasHigherNonZero, zeroTensWord, useTu);
                    parts.Add(groupWords);

                    string scaleUnit = GetScaleUnit(scaleIndex, thousandWord);
                    if (!string.IsNullOrEmpty(scaleUnit))
                    {
                        parts.Add(scaleUnit);
                    }
                }
                else
                {
                    // Nếu nhóm 3 chữ số bằng 000 nhưng nằm ở ranh giới hàng "tỷ" (scaleIndex chia hết cho 3, >= 3)
                    // thì vẫn đọc chữ "tỷ" (ví dụ: 1.000.000.000.000 -> một nghìn tỷ)
                    if (scaleIndex >= 3 && scaleIndex % 3 == 0)
                    {
                        parts.Add("tỷ");
                    }
                }
            }

            return string.Join(" ", parts);
        }

        private static string ReadThreeDigitGroup(int h, int t, int u, bool readZeroHundred, string zeroTensWord, bool useTu)
        {
            List<string> words = new List<string>(4);

            // Hàng trăm
            if (h > 0 || readZeroHundred)
            {
                words.Add(DigitWords[h]);
                words.Add("trăm");
            }

            // Hàng chục
            if (t == 0)
            {
                if (u > 0 && (h > 0 || readZeroHundred))
                {
                    words.Add(zeroTensWord);
                }
            }
            else if (t == 1)
            {
                words.Add("mười");
            }
            else
            {
                words.Add(DigitWords[t]);
                words.Add("mươi");
            }

            // Hàng đơn vị
            if (u > 0)
            {
                if (u == 1)
                {
                    words.Add(t >= 2 ? "mốt" : "một");
                }
                else if (u == 4)
                {
                    words.Add((t >= 2 && useTu) ? "tư" : "bốn");
                }
                else if (u == 5)
                {
                    words.Add(t >= 1 ? "lăm" : "năm");
                }
                else
                {
                    words.Add(DigitWords[u]);
                }
            }

            return string.Join(" ", words);
        }

        private static string GetScaleUnit(int scaleIndex, string thousandWord)
        {
            if (scaleIndex <= 0) return string.Empty;

            int mod = scaleIndex % 3;
            if (mod == 1) return thousandWord;
            if (mod == 2) return "triệu";
            return "tỷ";
        }

        private static string ReadDecimalFraction(string fracDigits, string thousandWord, string zeroTensWord, bool useTu)
        {
            if (string.IsNullOrEmpty(fracDigits))
            {
                return string.Empty;
            }

            List<string> parts = new List<string>();
            int idx = 0;
            while (idx < fracDigits.Length && fracDigits[idx] == '0')
            {
                parts.Add("không");
                idx++;
            }

            if (idx < fracDigits.Length)
            {
                string remaining = fracDigits.Substring(idx);
                if (remaining.Length <= 3)
                {
                    parts.Add(ReadIntegerDigits(remaining, thousandWord, zeroTensWord, useTu));
                }
                else
                {
                    for (int i = 0; i < remaining.Length; i++)
                    {
                        parts.Add(DigitWords[remaining[i] - '0']);
                    }
                }
            }

            return string.Join(" ", parts);
        }

        private static string NormalizeSubUnitDigits(string fracDigits)
        {
            if (string.IsNullOrEmpty(fracDigits)) return "0";
            if (fracDigits.Length == 1)
            {
                // VD: .5 USD -> 50 xu
                return fracDigits + "0";
            }
            if (fracDigits.Length >= 2)
            {
                string twoDigits = fracDigits.Substring(0, 2);
                return twoDigits.StartsWith("0") ? twoDigits.Substring(1) : twoDigits;
            }
            return fracDigits;
        }

        private static string NormalizeSpaces(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;
            string[] tokens = input.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(" ", tokens);
        }

        public static string ApplyCapitalization(string text, TextCapitalization cap)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            CultureInfo viCulture = new CultureInfo("vi-VN");

            switch (cap)
            {
                case TextCapitalization.LowerCase:
                    return text.ToLower(viCulture);

                case TextCapitalization.UpperCase:
                    return text.ToUpper(viCulture);

                case TextCapitalization.TitleCase:
                    string lower = text.ToLower(viCulture);
                    return viCulture.TextInfo.ToTitleCase(lower);

                case TextCapitalization.SentenceCase:
                default:
                    string s = text;
                    if (s.Length == 1) return s.ToUpper(viCulture);
                    return char.ToUpper(s[0], viCulture) + s.Substring(1);
            }
        }
    }
}
