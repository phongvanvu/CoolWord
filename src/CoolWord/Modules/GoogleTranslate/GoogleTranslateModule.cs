using System;
using System.Windows;
using CoolWord.Core;

namespace CoolWord.Modules.GoogleTranslate
{
    public class GoogleTranslateModule : IToolModule
    {
        public string Id { get { return "google-translate"; } }
        public string Title { get { return "Google Dịch"; } }
        public string Subtitle { get { return "Dịch đoạn văn, toàn bộ tài liệu và từ điển song ngữ Kế toán - Hợp đồng"; } }
        public string IconText { get { return "G文"; } }
        public string Category { get { return "Tiện ích & Ngôn ngữ"; } }
        public int Order { get { return 13; } }

        public bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)
        {
            if (wordService == null)
            {
                statusMessage = "Dịch vụ Word chưa sẵn sàng.";
                return false;
            }
            return wordService.TranslateSelectedText("auto", "vi", false, out statusMessage);
        }

        public UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            return new GoogleTranslateModuleView(wordService, notifyStatus);
        }
    }
}
