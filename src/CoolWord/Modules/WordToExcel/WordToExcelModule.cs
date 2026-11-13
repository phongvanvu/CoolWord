using System;
using System.Windows;
using CoolWord.Core;

namespace CoolWord.Modules.WordToExcel
{
    public class WordToExcelModule : IToolModule
    {
        public string Id { get { return "word-to-excel"; } }
        public string Title { get { return "Word ➔ Excel"; } }
        public string Subtitle { get { return "Xuất bảng biểu, chứng từ kế toán và danh sách từ Word sang Excel"; } }
        public string IconText { get { return "W➔X"; } }
        public string Category { get { return "Chuyển đổi & Tệp tin"; } }
        public int Order { get { return 9; } }

        public bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)
        {
            if (wordService == null)
            {
                statusMessage = "Dịch vụ Word chưa sẵn sàng.";
                return false;
            }
            return wordService.ExportAllTablesToExcel(null, true, false, out statusMessage);
        }

        public UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            return new WordToExcelModuleView(wordService, notifyStatus);
        }
    }
}
