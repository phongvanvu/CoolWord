using System;
using System.Windows;
using CoolWord.Core;

namespace CoolWord.Modules.Printing
{
    public class PrintModule : IToolModule
    {
        public string Id { get { return "print-utilities"; } }
        public string Title { get { return "In Ấn & Xuất Bản"; } }
        public string Subtitle { get { return "In 2 mặt thủ công, in trang chẵn/lẻ, in nhanh trang & xuất PDF"; } }
        public string IconText { get { return "In"; } }
        public string Category { get { return "In ấn & Xuất bản"; } }
        public int Order { get { return 4; } }

        public bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)
        {
            if (wordService == null)
            {
                statusMessage = "Dịch vụ Word chưa sẵn sàng.";
                return false;
            }
            return wordService.PrintCurrentPage(out statusMessage);
        }

        public UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            return new PrintModuleView(wordService, notifyStatus);
        }
    }
}
