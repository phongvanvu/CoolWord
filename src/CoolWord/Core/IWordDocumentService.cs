using System;
using CoolWord.Modules.AdminFormat;
using CoolWord.Modules.NumberToText;

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
    }
}
