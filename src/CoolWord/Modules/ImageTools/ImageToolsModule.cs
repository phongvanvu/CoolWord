using System;
using System.Windows;
using CoolWord.Core;

namespace CoolWord.Modules.ImageTools
{
    public class ImageToolsModule : IToolModule
    {
        public string Id { get { return "image-tools"; } }
        public string Title { get { return "Xử Lý Hình Ảnh"; } }
        public string Subtitle { get { return "Trích xuất ảnh chất lượng gốc, chèn con dấu/chữ ký số, chèn ảnh hàng loạt"; } }
        public string IconText { get { return "🖼"; } }
        public string Category { get { return "Tiện ích & Đồ họa"; } }
        public int Order { get { return 12; } }

        public bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)
        {
            if (wordService == null)
            {
                statusMessage = "Dịch vụ Word chưa sẵn sàng.";
                return false;
            }
            return wordService.ExportAllImages(null, out statusMessage);
        }

        public UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            return new ImageToolsModuleView(wordService, notifyStatus);
        }
    }
}
