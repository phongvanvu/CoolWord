using System;
using System.Windows;
using CoolWord.Core;

namespace CoolWord.Modules.TableOfContents
{
    public class TableOfContentsModule : IToolModule
    {
        public string Id { get { return "table-of-contents"; } }
        public string Title { get { return "Tạo Mục Lục"; } }
        public string Subtitle { get { return "Tạo mục lục tự động Heading 1-3, Báo cáo tài chính, gán Heading thông minh"; } }
        public string IconText { get { return "📋"; } }
        public string Category { get { return "Soạn thảo & Bố cục"; } }
        public int Order { get { return 14; } }

        public bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)
        {
            if (wordService == null)
            {
                statusMessage = "Dịch vụ Word chưa sẵn sàng.";
                return false;
            }
            return wordService.InsertTableOfContents(3, false, out statusMessage);
        }

        public UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            return new TableOfContentsModuleView(wordService, notifyStatus);
        }
    }
}
