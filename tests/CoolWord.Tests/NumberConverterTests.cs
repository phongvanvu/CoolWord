using System;
using CoolWord.Core;
using CoolWord.Modules.AdminFormat;
using CoolWord.Modules.NameSplit;
using CoolWord.Modules.NumberToText;
using CoolWord.Modules.Pdf;
using CoolWord.Modules.Printing;
using CoolWord.Modules.QuickInput;
using CoolWord.Modules.TextUtilities;
using CoolWord.Modules.WordToWord;
using CoolWord.Modules.WordToExcel;
using CoolWord.Modules.WordToPowerPoint;

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

            // 10. Kiểm thử Module Tiện Ích PDF (PdfModule, PdfOptions & PdfService)
            IToolModule pdfMod = ModuleRegistry.GetById("pdf-utilities");
            AssertEqual("PdfModule Registration", "true", (pdfMod != null).ToString().ToLowerInvariant());
            if (pdfMod != null)
            {
                AssertEqual("PdfModule Title", "Tiện Ích PDF", pdfMod.Title);
                AssertEqual("PdfModule Order", "5", pdfMod.Order.ToString());
            }

            PdfOptions accPdfOpts = PdfOptions.CreateAccountingPdfAStandard();
            AssertEqual("Accounting PDF/A Option UsePdfA", "true", accPdfOpts.UsePdfA.ToString().ToLowerInvariant());
            AssertEqual("Accounting PDF/A Option OpenAfterExport", "true", accPdfOpts.OpenAfterExport.ToString().ToLowerInvariant());

            // Kiểm thử các hàm xử lý PdfService với null app (xử lý ngoại lệ an toàn)
            string pdfStatus;
            bool pdfRes = PdfService.ExportAccountingPdfA(null, null, out pdfStatus);
            AssertEqual("ExportAccountingPdfA Null Check", "false", pdfRes.ToString().ToLowerInvariant());

            string wmStatus;
            bool wmRes = PdfService.AddAccountingWatermark(null, "ĐÃ THANH TOÁN", out wmStatus);
            AssertEqual("AddAccountingWatermark Null Check", "false", wmRes.ToString().ToLowerInvariant());

            // 11. Kiểm thử Module Nhập Liệu Nhanh (QuickInputModule & QuickInputService)
            IToolModule qiMod = ModuleRegistry.GetById("quick-input");
            AssertEqual("QuickInputModule Registration", "true", (qiMod != null).ToString().ToLowerInvariant());
            if (qiMod != null)
            {
                AssertEqual("QuickInputModule Title", "Nhập Liệu Nhanh", qiMod.Title);
                AssertEqual("QuickInputModule Order", "6", qiMod.Order.ToString());
            }

            var builtIns = QuickInputService.GetBuiltInSnippets();
            AssertEqual("BuiltIn Snippets Count > 0", "true", (builtIns.Count >= 10).ToString().ToLowerInvariant());

            string dateTemplate = "Hôm nay {DATE}";
            string resolvedDate = QuickInputService.ResolvePlaceholders(dateTemplate);
            AssertEqual("ResolvePlaceholders DATE", "true", resolvedDate.Contains(DateTime.Now.ToString("dd/MM/yyyy")).ToString().ToLowerInvariant());

            string docDateTemplate = "Hà Nội, {DOC_DATE}";
            string resolvedDocDate = QuickInputService.ResolvePlaceholders(docDateTemplate);
            AssertEqual("ResolvePlaceholders DOC_DATE", "true", resolvedDocDate.Contains("tháng").ToString().ToLowerInvariant());

            // Kiểm thử Thêm & Xóa mẫu cá nhân
            string addStatus;
            bool addOk = QuickInputService.AddCustomSnippet("Mẫu thử nghiệm", "Cá nhân", "Nội dung mẫu {DATE}", "mtn", out addStatus);
            AssertEqual("AddCustomSnippet", "true", addOk.ToString().ToLowerInvariant());

            var allSnippets = QuickInputService.GetAllSnippets();
            var added = allSnippets.Find(s => s.Title == "Mẫu thử nghiệm");
            AssertEqual("Find Added Snippet", "true", (added != null).ToString().ToLowerInvariant());

            if (added != null)
            {
                string delStatus;
                bool delOk = QuickInputService.DeleteCustomSnippet(added.Id, out delStatus);
                AssertEqual("DeleteCustomSnippet", "true", delOk.ToString().ToLowerInvariant());
            }

            // Kiểm thử hàm Insert với null app (xử lý ngoại lệ an toàn)
            string insertStatus;
            bool insertRes = QuickInputService.InsertSnippetIntoWord(null, "Test", out insertStatus);
            AssertEqual("InsertSnippetIntoWord Null Check", "false", insertRes.ToString().ToLowerInvariant());

            // 12. Kiểm thử Module Tách Họ - Tên (VietnameseNameSplitter & NameSplitModule)
            IToolModule nameMod = ModuleRegistry.GetById("name-split");
            AssertEqual("NameSplitModule Registration", "true", (nameMod != null).ToString().ToLowerInvariant());
            if (nameMod != null)
            {
                AssertEqual("NameSplitModule Title", "Tách Họ - Tên", nameMod.Title);
                AssertEqual("NameSplitModule Order", "7", nameMod.Order.ToString());
            }

            // Trường hợp 3 từ: "Nguyễn Văn An"
            ParsedVietnameseName p1 = VietnameseNameSplitter.Split("Nguyễn Văn An");
            AssertEqual("Split 3-words HoDem", "Nguyễn Văn", p1.HoDem);
            AssertEqual("Split 3-words Ten", "An", p1.Ten);
            AssertEqual("Split 3-words Ho", "Nguyễn", p1.Ho);
            AssertEqual("Split 3-words Dem", "Văn", p1.Dem);

            // Trường hợp 4 từ: "Trần Thị Bích Ngọc"
            ParsedVietnameseName p2 = VietnameseNameSplitter.Split("Trần Thị Bích Ngọc");
            AssertEqual("Split 4-words HoDem", "Trần Thị Bích", p2.HoDem);
            AssertEqual("Split 4-words Ten", "Ngọc", p2.Ten);
            AssertEqual("Split 4-words Ho", "Trần", p2.Ho);
            AssertEqual("Split 4-words Dem", "Thị Bích", p2.Dem);

            // Trường hợp 2 từ: "Lê Nam"
            ParsedVietnameseName p3 = VietnameseNameSplitter.Split("Lê Nam");
            AssertEqual("Split 2-words HoDem", "Lê", p3.HoDem);
            AssertEqual("Split 2-words Ten", "Nam", p3.Ten);
            AssertEqual("Split 2-words Ho", "Lê", p3.Ho);
            AssertEqual("Split 2-words Dem", "", p3.Dem);

            // Chuẩn hóa họ tên viết hoa
            string normName = VietnameseNameSplitter.NormalizeName("  nguyễn   văn   an  ");
            AssertEqual("Normalize Name TitleCase", "Nguyễn Văn An", normName);

            // Chuẩn hóa họ tên dân tộc thiểu số
            string normEthnic = VietnameseNameSplitter.NormalizeName("h'hen niê");
            AssertEqual("Normalize Ethnic Name", "H'Hen Niê", normEthnic);

            // Gộp Họ đệm và Tên
            string merged = VietnameseNameSplitter.Merge("Nguyễn Văn", "An");
            AssertEqual("Merge Name", "Nguyễn Văn An", merged);

            // Kiểm thử các hàm Word table với null app (xử lý ngoại lệ an toàn)
            string splitTableStatus;
            bool splitTableRes = NameSplitService.SplitTableColumn2(null, out splitTableStatus);
            AssertEqual("SplitTableColumn2 Null Check", "false", splitTableRes.ToString().ToLowerInvariant());

            // 7. Kiểm thử Mô-đun Word ➔ Word (TCVN3, VNI, Registry, An toàn null)
            string tcvnText = "\u00A7\u00B6ng c\u00EAng s\u00B6n"; // "Đảng cộng sản"
            string uFromTcvn = VietnameseEncodingConverter.ConvertTcvn3ToUnicode(tcvnText);
            AssertEqual("TCVN3 to Unicode", "Đảng cộng sản", uFromTcvn);

            string vniSample = "Vieät Nam ñeán"; // "Việt Nam đến"
            string uFromVni = VietnameseEncodingConverter.ConvertVniToUnicode(vniSample);
            AssertEqual("VNI to Unicode", "Việt Nam đến", uFromVni);

            AssertEqual("IsTcvn3Font .VnTime", "true", VietnameseEncodingConverter.IsTcvn3Font(".VnTime").ToString().ToLowerInvariant());
            AssertEqual("IsTcvn3Font Times", "false", VietnameseEncodingConverter.IsTcvn3Font("Times New Roman").ToString().ToLowerInvariant());
            AssertEqual("IsVniFont VNI-Times", "true", VietnameseEncodingConverter.IsVniFont("VNI-Times").ToString().ToLowerInvariant());
            AssertEqual("IsVniFont Arial", "false", VietnameseEncodingConverter.IsVniFont("Arial").ToString().ToLowerInvariant());

            IToolModule w2wMod = ModuleRegistry.GetById("word-to-word");
            AssertEqual("WordToWord Module Registered", "true", (w2wMod != null).ToString().ToLowerInvariant());
            if (w2wMod != null)
            {
                AssertEqual("WordToWord Module Title", "Word ➔ Word", w2wMod.Title);
            }

            string w2wStatus;
            bool w2wNullRes = WordToWordService.ConvertEncodingToUnicode(null, false, "TCVN3", out w2wStatus);
            AssertEqual("WordToWord Convert Null Safe", "false", w2wNullRes.ToString().ToLowerInvariant());

            bool w2wDocxRes = WordToWordService.SaveAsDocx(null, out w2wStatus);
            AssertEqual("WordToWord Docx Null Safe", "false", w2wDocxRes.ToString().ToLowerInvariant());

            bool w2wDoc97Res = WordToWordService.SaveAsDoc97(null, out w2wStatus);
            AssertEqual("WordToWord Doc97 Null Safe", "false", w2wDoc97Res.ToString().ToLowerInvariant());

            bool w2wMergeRes = WordToWordService.MergeWordDocuments(null, null, out w2wStatus);
            AssertEqual("WordToWord Merge Null Safe", "false", w2wMergeRes.ToString().ToLowerInvariant());

            bool w2wSplitSecRes = WordToWordService.SplitDocumentBySections(null, null, out w2wStatus);
            AssertEqual("WordToWord SplitSec Null Safe", "false", w2wSplitSecRes.ToString().ToLowerInvariant());

            bool w2wSplitPageRes = WordToWordService.SplitDocumentByPages(null, null, out w2wStatus);
            AssertEqual("WordToWord SplitPage Null Safe", "false", w2wSplitPageRes.ToString().ToLowerInvariant());

            // 8. Kiểm thử Mô-đun Word ➔ Excel (Dữ liệu, số kế toán, Registry, An toàn null)
            string cleanCell = WordToExcelService.CleanCellText("Doanh thu năm 2026\r\a");
            AssertEqual("CleanCellText", "Doanh thu năm 2026", cleanCell);

            double accVal1;
            bool accParse1 = WordToExcelService.TryParseAccountingNumber("1.250.000", out accVal1);
            AssertEqual("TryParseAccountingNumber 1.250.000", "true", accParse1.ToString().ToLowerInvariant());
            AssertEqual("Accounting Number Value", "1250000", ((long)accVal1).ToString());

            double accVal2;
            bool accParse2 = WordToExcelService.TryParseAccountingNumber("(500.000)", out accVal2);
            AssertEqual("TryParseAccountingNumber Negative (500.000)", "true", accParse2.ToString().ToLowerInvariant());
            AssertEqual("Negative Number Value", "-500000", ((long)accVal2).ToString());

            double accVal3;
            bool accParse3 = WordToExcelService.TryParseAccountingNumber("15%", out accVal3);
            AssertEqual("TryParseAccountingNumber 15%", "true", accParse3.ToString().ToLowerInvariant());
            AssertEqual("Percent Value", "15", ((long)accVal3).ToString());

            IToolModule w2eMod = ModuleRegistry.GetById("word-to-excel");
            AssertEqual("WordToExcel Module Registered", "true", (w2eMod != null).ToString().ToLowerInvariant());
            if (w2eMod != null)
            {
                AssertEqual("WordToExcel Module Title", "Word ➔ Excel", w2eMod.Title);
            }

            string w2eStatus;
            bool w2eAllRes = WordToExcelService.ExportAllTablesToExcel(null, null, false, false, out w2eStatus);
            AssertEqual("WordToExcel AllTables Null Safe", "false", w2eAllRes.ToString().ToLowerInvariant());

            bool w2eSelRes = WordToExcelService.ExportSelectedTableToExcel(null, null, false, out w2eStatus);
            AssertEqual("WordToExcel SelTable Null Safe", "false", w2eSelRes.ToString().ToLowerInvariant());

            bool w2eAccRes = WordToExcelService.ExportAccountingTableToExcel(null, null, false, out w2eStatus);
            AssertEqual("WordToExcel AccTable Null Safe", "false", w2eAccRes.ToString().ToLowerInvariant());

            bool w2eParaRes = WordToExcelService.ExportParagraphsToExcel(null, null, false, out w2eStatus);
            AssertEqual("WordToExcel Paras Null Safe", "false", w2eParaRes.ToString().ToLowerInvariant());

            // 9. Kiểm thử Mô-đun Word ➔ PowerPoint (Registry, An toàn null)
            IToolModule w2pMod = ModuleRegistry.GetById("word-to-ppt");
            AssertEqual("WordToPpt Module Registered", "true", (w2pMod != null).ToString().ToLowerInvariant());
            if (w2pMod != null)
            {
                AssertEqual("WordToPpt Module Title", "Word ➔ PowerPoint", w2pMod.Title);
            }

            string w2pStatus;
            bool w2pHeadRes = WordToPowerPointService.CreateSlidesFromHeadings(null, null, false, out w2pStatus);
            AssertEqual("WordToPpt Headings Null Safe", "false", w2pHeadRes.ToString().ToLowerInvariant());

            bool w2pPageRes = WordToPowerPointService.CreateSlidesFromPages(null, null, false, out w2pStatus);
            AssertEqual("WordToPpt Pages Null Safe", "false", w2pPageRes.ToString().ToLowerInvariant());

            bool w2pSecRes = WordToPowerPointService.CreateSlidesFromSections(null, null, false, out w2pStatus);
            AssertEqual("WordToPpt Sections Null Safe", "false", w2pSecRes.ToString().ToLowerInvariant());

            bool w2pTableRes = WordToPowerPointService.ExportTablesToSlides(null, null, false, out w2pStatus);
            AssertEqual("WordToPpt Tables Null Safe", "false", w2pTableRes.ToString().ToLowerInvariant());

            bool w2pFinRes = WordToPowerPointService.CreateFinancialPresentation(null, null, false, out w2pStatus);
            AssertEqual("WordToPpt FinancialPresentation Null Safe", "false", w2pFinRes.ToString().ToLowerInvariant());

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
