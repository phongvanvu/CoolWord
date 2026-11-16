using System;
using System.Windows;
using CoolWord.Core;

namespace CoolWord.Modules.WordToWord
{
    public class WordToWordModule : IToolModule
    {
        public string Id { get { return "word-to-word"; } }
        public string Title { get { return "Word ➔ Word"; } }
        public string Subtitle { get { return "Chuyển mã TCVN3/VNI, đổi .doc/.docm ➔ .docx, gộp/tách file Word"; } }
        public string IconText { get { return "W➔W"; } }
        public string Category { get { return "Chuyển đổi & Tệp tin"; } }
        public int Order { get { return 8; } }

        public bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)
        {
            if (wordService == null)
            {
                statusMessage = "Dịch vụ Word chưa sẵn sàng.";
                return false;
            }
            return wordService.ConvertEncodingToUnicode(false, "AUTO", out statusMessage);
        }

        public UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            return new WordToWordModuleView(wordService, notifyStatus);
        }
    }
}
