using System;
using System.Windows;
using CoolWord.Core;

namespace CoolWord.Modules.QuickInput
{
    public class QuickInputModule : IToolModule
    {
        public string Id { get { return "quick-input"; } }
        public string Title { get { return "Nhập Liệu Nhanh"; } }
        public string Subtitle { get { return "Chèn nhanh mẫu văn bản, biểu mẫu kế toán, căn cứ pháp lý & cụm từ viết tắt"; } }
        public string IconText { get { return "Mẫu"; } }
        public string Category { get { return "Soạn thảo & Nhập liệu"; } }
        public int Order { get { return 6; } }

        public bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)
        {
            if (wordService == null)
            {
                statusMessage = "Dịch vụ Word chưa sẵn sàng.";
                return false;
            }
            return wordService.SaveCurrentSelectionAsSnippet(null, "Cá nhân", null, out statusMessage);
        }

        public UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            return new QuickInputModuleView(wordService, notifyStatus);
        }
    }
}
