using System;

namespace CoolWord.Modules.NumberToText
{
    public enum NumberReadingMode
    {
        CurrencyVnd = 0,
        CurrencyUsd = 1,
        PureNumber = 2,
        CustomCurrency = 3
    }

    public enum VietnameseDialect
    {
        StandardNorth = 0, // nghìn, linh, tư
        StandardSouth = 1, // ngàn, lẻ, bốn
        MixedCommon = 2    // nghìn, lẻ, tư (phổ biến trong văn bản hành chính)
    }

    public enum TextCapitalization
    {
        SentenceCase = 0,  // Một triệu hai trăm...
        LowerCase = 1,     // một triệu hai trăm...
        TitleCase = 2,     // Một Triệu Hai Trăm...
        UpperCase = 3      // MỘT TRIỆU HAI TRĂM...
    }

    public enum WordInsertMode
    {
        ReplaceSelection = 0,              // 1.250.000 -> Một triệu hai trăm năm mươi nghìn đồng chẵn
        AppendInParentheses = 1,           // 1.250.000 -> 1.250.000 (Một triệu hai trăm năm mươi nghìn đồng chẵn)
        AppendWithLabelInParentheses = 2,  // 1.250.000 -> 1.250.000 (Bằng chữ: Một triệu hai trăm năm mươi nghìn đồng chẵn)
        AppendOnNewLine = 3                // Chèn dòng mới phía dưới: Bằng chữ: Một triệu...
    }

    public class NumberConvertOptions
    {
        public NumberReadingMode ReadingMode { get; set; }
        public VietnameseDialect Dialect { get; set; }
        public TextCapitalization Capitalization { get; set; }
        public WordInsertMode InsertMode { get; set; }

        public bool AppendChan { get; set; }
        public bool AppendDotSlash { get; set; }
        public bool ItalicizeConvertedText { get; set; }
        public bool AutoDetectCurrencyFromSelection { get; set; }

        public string CustomMajorUnit { get; set; }
        public string CustomMinorUnit { get; set; }

        public NumberConvertOptions()
        {
            ReadingMode = NumberReadingMode.CurrencyVnd;
            Dialect = VietnameseDialect.MixedCommon;
            Capitalization = TextCapitalization.SentenceCase;
            InsertMode = WordInsertMode.AppendInParentheses;
            AppendChan = true;
            AppendDotSlash = false;
            ItalicizeConvertedText = true;
            AutoDetectCurrencyFromSelection = true;
            CustomMajorUnit = "đồng";
            CustomMinorUnit = "xu";
        }

        public string GetThousandWord()
        {
            return Dialect == VietnameseDialect.StandardSouth ? "ngàn" : "nghìn";
        }

        public string GetZeroTensWord()
        {
            return Dialect == VietnameseDialect.StandardNorth ? "linh" : "lẻ";
        }

        public bool UseTuForFour()
        {
            return Dialect != VietnameseDialect.StandardSouth;
        }

        public NumberConvertOptions Clone()
        {
            return (NumberConvertOptions)MemberwiseClone();
        }
    }
}
