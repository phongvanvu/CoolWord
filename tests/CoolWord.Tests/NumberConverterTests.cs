using System;
using CoolWord.Core;
using CoolWord.Modules.AdminFormat;
using CoolWord.Modules.NumberToText;
using CoolWord.Modules.Printing;
using CoolWord.Modules.TextUtilities;

namespace CoolWord.Tests
{
    public static class NumberConverterTests
    {
        private static int _passed;
        private static int _failed;

        public static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("=== CHẠY KIỂM THỬ TỰ ĐỘNG COOLWORD (VIETNAMESE NUMBER CONVERTER) ===");

            NumberConvertOptions commonVnd = new NumberConvertOptions
            {
                ReadingMode = NumberReadingMode.CurrencyVnd,
                Dialect = VietnameseDialect.MixedCommon,
                Capitalization = TextCapitalization.SentenceCase,
                AppendChan = true
            };

            // 1. Các trường hợp cơ bản & quy tắc âm vị tiếng Việt (mốt, tư, lăm, lẻ, không trăm)
            AssertConvert("0", commonVnd, "Không đồng chẵn");
            AssertConvert("5", commonVnd, "Năm đồng chẵn");
            AssertConvert("10", commonVnd, "Mười đồng chẵn");
            AssertConvert("11", commonVnd, "Mười một đồng chẵn");
            AssertConvert("14", commonVnd, "Mười bốn đồng chẵn");
            AssertConvert("15", commonVnd, "Mười lăm đồng chẵn");
            AssertConvert("21", commonVnd, "Hai mươi mốt đồng chẵn");
            AssertConvert("24", commonVnd, "Hai mươi tư đồng chẵn");
            AssertConvert("25", commonVnd, "Hai mươi lăm đồng chẵn");
            AssertConvert("101", commonVnd, "Một trăm lẻ một đồng chẵn");
            AssertConvert("105", commonVnd, "Một trăm lẻ năm đồng chẵn");
            AssertConvert("115", commonVnd, "Một trăm mười lăm đồng chẵn");
            AssertConvert("125", commonVnd, "Một trăm hai mươi lăm đồng chẵn");

            // 2. Hàng nghìn, triệu, tỷ, nghìn tỷ & nhóm 000 ở giữa
            AssertConvert("1.000", commonVnd, "Một nghìn đồng chẵn");
            AssertConvert("1.005", commonVnd, "Một nghìn không trăm lẻ năm đồng chẵn");
            AssertConvert("1.050", commonVnd, "Một nghìn không trăm năm mươi đồng chẵn");
            AssertConvert("1.000.005", commonVnd, "Một triệu không trăm lẻ năm đồng chẵn");
            AssertConvert("1.250.000", commonVnd, "Một triệu hai trăm năm mươi nghìn đồng chẵn");
            AssertConvert("15.680.500", commonVnd, "Mười lăm triệu sáu trăm tám mươi nghìn năm trăm đồng chẵn");
            AssertConvert("1.000.000.000", commonVnd, "Một tỷ đồng chẵn");
            AssertConvert("1.000.000.000.000", commonVnd, "Một nghìn tỷ đồng chẵn");
            AssertConvert("1.002.000.000.000", commonVnd, "Một nghìn không trăm lẻ hai tỷ đồng chẵn");

            // 3. Kiểm thử Phương ngữ Bắc Bộ (nghìn, linh, tư) vs Nam Bộ (ngàn, lẻ, bốn)
            NumberConvertOptions northOpts = commonVnd.Clone();
            northOpts.Dialect = VietnameseDialect.StandardNorth;
            AssertConvert("1.005.024", northOpts, "Một triệu không trăm linh năm nghìn không trăm hai mươi tư đồng chẵn");

            NumberConvertOptions southOpts = commonVnd.Clone();
            southOpts.Dialect = VietnameseDialect.StandardSouth;
            AssertConvert("1.005.024", southOpts, "Một triệu không trăm lẻ năm ngàn không trăm hai mươi bốn đồng chẵn");

            // 4. Kiểm thử Ngoại tệ USD & Tự động nhận diện ký hiệu tiền tệ
            AssertConvert("$1,250.50", commonVnd, "Một nghìn hai trăm năm mươi đô la Mỹ và năm mươi xu");
            AssertConvert("2.450,05 USD", commonVnd, "Hai nghìn bốn trăm năm mươi đô la Mỹ và năm xu");

            // 5. Kiểm thử Số thuần túy & Số thập phân, Số âm
            NumberConvertOptions pureOpts = commonVnd.Clone();
            pureOpts.ReadingMode = NumberReadingMode.PureNumber;
            AssertConvert("3,14", pureOpts, "Ba phẩy mười bốn");
            AssertConvert("0,05", pureOpts, "Không phẩy không năm");
            AssertConvert("-1.250", pureOpts, "Âm một nghìn hai trăm năm mươi");

            // 6. Kiểm thử Định dạng chèn vào Word
            NumberConvertOptions insertOpts = commonVnd.Clone();
            insertOpts.InsertMode = WordInsertMode.AppendInParentheses;
            string words = VietnameseNumberConverter.Convert("1.250.000", insertOpts);
            string formatted = VietnameseNumberConverter.FormatForWordInsertion("1.250.000", words, insertOpts);
            AssertEqual("Format AppendInParentheses", "1.250.000 (Một triệu hai trăm năm mươi nghìn đồng chẵn)", formatted);

            // 7. Kiểm thử Module Chuẩn hóa văn bản Tiếng Việt
            string dirtyText = "Cộng   hòa xã hội   chủ nghĩa Việt Nam ,Độc lập - Tự do - Hạnh phúc ! (  Báo cáo tài chính ) .";
            string cleanedText = TextCleanModule.NormalizeVietnameseText(dirtyText);
            AssertEqual("NormalizeVietnameseText", "Cộng hòa xã hội chủ nghĩa Việt Nam, Độc lập - Tự do - Hạnh phúc! (Báo cáo tài chính).", cleanedText);

            // 8. Kiểm thử Quy chuẩn Nghị định 30/2020/NĐ-CP (Định dạng ngày tháng & Đọc ngày bằng chữ)
            // Quy tắc NĐ 30: Ngày < 10 và Tháng 1, 2 phải có số 0 phía trước; các tháng 3..12 không có số 0
            string adminDateJan = Decree30Formatter.FormatAdminDateLine("Hà Nội", new DateTime(2026, 2, 5));
            AssertEqual("Decree30 Date (05/02/2026)", "Hà Nội, ngày 05 tháng 02 năm 2026", adminDateJan);

            string adminDateMay = Decree30Formatter.FormatAdminDateLine("TP. Hồ Chí Minh", new DateTime(2026, 5, 15));
            AssertEqual("Decree30 Date (15/5/2026)", "TP. Hồ Chí Minh, ngày 15 tháng 5 năm 2026", adminDateMay);

            string wordsDate = Decree30Formatter.FormatDateInVietnameseWords(new DateTime(2026, 11, 13), "Hà Nội");
            AssertEqual("Date In Vietnamese Words", "Hà Nội, ngày mười ba tháng mười một năm hai nghìn không trăm hai mươi sáu", wordsDate);

            // 9. Kiểm thử Module In Ấn & Xuất Bản (PrintModule & PrintService)
            ModuleRegistry.EnsureInitialized();
            IToolModule printMod = ModuleRegistry.GetById("print-utilities");
            AssertEqual("PrintModule Registration", "true", (printMod != null).ToString().ToLowerInvariant());
            if (printMod != null)
            {
                AssertEqual("PrintModule Title", "In Ấn & Xuất Bản", printMod.Title);
                AssertEqual("PrintModule Order", "4", printMod.Order.ToString());
            }

            // Kiểm thử các hàm xử lý PrintService với null app (xử lý ngoại lệ an toàn)
            string printStatus;
            bool printRes = PrintService.PrintCurrentPage(null, out printStatus);
            AssertEqual("PrintCurrentPage Null Check", "false", printRes.ToString().ToLowerInvariant());

            string voucherStatus;
            bool voucherRes = PrintService.PrintMultiCopyAccountingVoucher(null, 2, "Phiếu thu", false, out voucherStatus);
            AssertEqual("PrintMultiCopyAccountingVoucher Null Check", "false", voucherRes.ToString().ToLowerInvariant());

            Console.WriteLine("--------------------------------------------------------------------");
            Console.WriteLine(string.Format("KẾT QUẢ: Đạt {0} / {1} bài kiểm thử (Thất bại: {2})", _passed, _passed + _failed, _failed));
            return _failed == 0 ? 0 : 1;
        }

        private static void AssertConvert(string input, NumberConvertOptions options, string expected)
        {
            string actual = VietnameseNumberConverter.Convert(input, options);
            AssertEqual("Convert [" + input + "]", expected, actual);
        }

        private static void AssertEqual(string testName, string expected, string actual)
        {
            if (string.Equals(expected, actual, StringComparison.Ordinal))
            {
                _passed++;
                Console.WriteLine("[PASS] " + testName + " -> " + actual);
            }
            else
            {
                _failed++;
                Console.WriteLine("[FAIL] " + testName);
                Console.WriteLine("   Mong đợi : " + expected);
                Console.WriteLine("   Thực tế  : " + actual);
            }
        }
    }
}
