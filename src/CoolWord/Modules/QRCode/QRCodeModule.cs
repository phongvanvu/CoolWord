using System;
using System.Windows;
using CoolWord.Core;

namespace CoolWord.Modules.QRCode
{
    public class QRCodeModule : IToolModule
    {
        public string Id { get { return "qr-code"; } }
        public string Title { get { return "Mã QR Code"; } }
        public string Subtitle { get { return "Tạo & chèn mã QR Code (Liên kết, Văn bản, VietQR Ngân hàng, Danh thiếp vCard)"; } }
        public string IconText { get { return "QR"; } }
        public string Category { get { return "Tiện ích & Xuất bản"; } }
        public int Order { get { return 15; } }

        public bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)
        {
            if (wordService == null)
            {
                statusMessage = "Dịch vụ Word chưa sẵn sàng.";
                return false;
            }

            return wordService.InsertQrCodeFromSelection(out statusMessage);
        }

        public UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            return new QRCodeModuleView(wordService, notifyStatus);
        }
    }
}
