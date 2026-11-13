using System;
using System.Windows;
using CoolWord.Core;

namespace CoolWord.Modules.BatchEmail
{
    public class BatchEmailModule : IToolModule
    {
        public string Id { get { return "batch-email"; } }
        public string Title { get { return "Gửi Mail Hàng Loạt"; } }
        public string Subtitle { get { return "Gửi mail trộn, gửi phiếu lương/chứng từ kế toán và gửi tài liệu Word qua Outlook/Gmail"; } }
        public string IconText { get { return "✉"; } }
        public string Category { get { return "Tiện ích & Kết nối"; } }
        public int Order { get { return 11; } }

        public bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)
        {
            if (wordService == null)
            {
                statusMessage = "Dịch vụ Word chưa sẵn sàng.";
                return false;
            }
            return wordService.SendCurrentDocumentViaEmail(string.Empty, string.Empty, string.Empty, true, out statusMessage);
        }

        public UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            return new BatchEmailModuleView(wordService, notifyStatus);
        }
    }
}
