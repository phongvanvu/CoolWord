using System;
using System.Windows;
using CoolWord.Core;

namespace CoolWord.Modules.WordToPowerPoint
{
    public class WordToPowerPointModule : IToolModule
    {
        public string Id { get { return "word-to-ppt"; } }
        public string Title { get { return "Word ➔ PowerPoint"; } }
        public string Subtitle { get { return "Tự động chuyển dàn ý, đề mục và nội dung Word thành bài thuyết trình PowerPoint"; } }
        public string IconText { get { return "W➔P"; } }
        public string Category { get { return "Chuyển đổi & Tệp tin"; } }
        public int Order { get { return 10; } }

        public bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)
        {
            if (wordService == null)
            {
                statusMessage = "Dịch vụ Word chưa sẵn sàng.";
                return false;
            }
            return wordService.CreateSlidesFromHeadings(null, true, out statusMessage);
        }

        public UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            return new WordToPowerPointModuleView(wordService, notifyStatus);
        }
    }
}
