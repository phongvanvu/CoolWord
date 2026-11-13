using System;
using System.Windows;
using CoolWord.Core;

namespace CoolWord.Modules.NameSplit
{
    public class NameSplitModule : IToolModule
    {
        public string Id { get { return "name-split"; } }
        public string Title { get { return "Tách Họ - Tên"; } }
        public string Subtitle { get { return "Tách cột họ đệm & tên trong bảng biểu, chuẩn hóa danh sách nhân sự & học sinh"; } }
        public string IconText { get { return "HọTên"; } }
        public string Category { get { return "Soạn thảo & Bảng biểu"; } }
        public int Order { get { return 7; } }

        public bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)
        {
            if (wordService == null)
            {
                statusMessage = "Dịch vụ Word chưa sẵn sàng.";
                return false;
            }
            return wordService.SplitTableColumn2(out statusMessage);
        }

        public UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            return new NameSplitModuleView(wordService, notifyStatus);
        }
    }
}
