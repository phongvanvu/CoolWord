using System;
using System.Windows;

namespace CoolWord.Core
{
    /// <summary>
    /// Giao diện chuẩn cho mọi Module tính năng trong CoolWord.
    /// Khi muốn mở rộng thêm tính năng mới sau này, chỉ cần cài đặt IToolModule và đăng ký vào ModuleRegistry.
    /// </summary>
    public interface IToolModule
    {
        string Id { get; }
        string Title { get; }
        string Subtitle { get; }
        string IconText { get; }
        string Category { get; }
        int Order { get; }

        /// <summary>
        /// Thực thi nhanh hành động mặc định của module trên vùng văn bản đang chọn trong Word.
        /// </summary>
        bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage);

        /// <summary>
        /// Tạo giao diện WPF chi tiết của module để hiển thị trong Task Pane hoặc CoolWord Studio.
        /// </summary>
        UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus);
    }
}
