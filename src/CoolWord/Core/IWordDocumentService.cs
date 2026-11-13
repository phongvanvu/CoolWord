using System;
using CoolWord.Modules.AdminFormat;
using CoolWord.Modules.NumberToText;
using CoolWord.Modules.Pdf;

namespace CoolWord.Core
{
    public class WordSelectionEventArgs : EventArgs
    {
        public string SelectedText { get; set; }
    }

    public interface IWordDocumentService
    {
        bool IsConnected { get; }
        event EventHandler<WordSelectionEventArgs> SelectionChanged;

        string GetSelectedText();
        bool ConvertSelectedNumber(NumberConvertOptions options, out string convertedWords, out string statusMessage);
        int ConvertAllNumbersInSelection(NumberConvertOptions options, out string statusMessage);
        bool InsertConvertedText(string originalInput, string convertedWords, NumberConvertOptions options, out string statusMessage);
        bool TransformSelectedText(Func<string, string> transformer, string actionName, out string statusMessage);

        // Các tác vụ Chuẩn hóa Văn bản Hành chính Nghị định 30/2020/NĐ-CP
        bool ApplyDecree30FullStandard(AdminDocOptions options, out string statusMessage);
        bool ApplyDecree30MarginsOnly(AdminDocOptions options, out string statusMessage);
        bool InsertDecree30Header(AdminDocOptions options, out string statusMessage);
        bool InsertDecree30FooterSignature(AdminDocOptions options, out string statusMessage);
        bool CleanFullDocumentGarbage(out string statusMessage);

        // Các tác vụ In ấn & Xuất bản thông minh
        bool PrintCurrentPage(out string statusMessage);
        bool PrintSelection(out string statusMessage);
        bool PrintManualDuplex(out string statusMessage);
        bool PrintOddPagesOnly(out string statusMessage);
        bool PrintEvenPagesOnly(out string statusMessage);
        bool PrintTwoPagesPerSheet(out string statusMessage);
        bool ExportQuickPdf(string targetFilePath, out string statusMessage);

        // Nghiệp vụ kế toán: In chứng từ kế toán nhiều liên (Không hiển thị trên UI)
        bool PrintMultiCopyAccountingVoucher(int copies, string voucherType, bool fitTwoPerSheet, out string statusMessage);

        // Các tác vụ Tiện ích PDF & Chuẩn lưu trữ kế toán PDF/A
        bool ExportCustomPdf(PdfOptions options, string targetFilePath, out string statusMessage);
        bool ExportAccountingPdfA(string targetFilePath, out string statusMessage);
        bool ExportSelectionToPdf(string targetFilePath, out string statusMessage);
        bool ExportCurrentPageToPdf(string targetFilePath, out string statusMessage);
        bool AddAccountingWatermark(string watermarkText, out string statusMessage);
        bool RemoveAccountingWatermark(out string statusMessage);

        // Các tác vụ Nhập Liệu Nhanh (Quick Input & Snippets)
        bool InsertTextSnippet(string text, out string statusMessage);
        bool SaveCurrentSelectionAsSnippet(string title, string category, string shortcut, out string statusMessage);

        // Các tác vụ Tách Họ - Tên (Vietnamese Name Splitting)
        bool SplitTableColumn2(out string statusMessage);
        bool SplitTableColumn3(out string statusMessage);
        bool MergeTableColumns(out string statusMessage);
        bool SplitSelectedTextLines(string delimiter, out string statusMessage);
        bool NormalizeSelectedNames(out string statusMessage);
    }
}
