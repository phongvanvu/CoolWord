using System;
using System.Windows;
using CoolWord.Core;

namespace CoolWord.Modules.Pdf
{
    public class PdfModule : IToolModule
    {
        public string Id { get { return "pdf-utilities"; } }
        public string Title { get { return "Tiện Ích PDF"; } }
        public string Subtitle { get { return "Xuất PDF chuẩn lưu trữ PDF/A kế toán & thuế, xuất vùng chọn & đóng dấu chìm"; } }
        public string IconText { get { return "PDF"; } }
        public string Category { get { return "Xuất bản & PDF"; } }
        public int Order { get { return 5; } }

        public bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)
        {
            if (wordService == null)
            {
                statusMessage = "Dịch vụ Word chưa sẵn sàng.";
                return false;
            }
            return wordService.ExportAccountingPdfA(null, out statusMessage);
        }

        public UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            return new PdfModuleView(wordService, notifyStatus);
        }
    }
}
