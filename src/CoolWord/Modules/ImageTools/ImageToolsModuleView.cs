using System;
using System.Windows;
using System.Windows.Controls;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Modules.ImageTools
{
    public class ImageToolsModuleView : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Action<string> _notifyStatus;

        private CheckBox _chkStampInFront;

        public ImageToolsModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            _wordService = wordService;
            _notifyStatus = notifyStatus ?? (_ => { });

            BuildUi();
        }

        private void BuildUi()
        {
            StackPanel root = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };

            // 1. Card: Trích xuất hình ảnh
            Border exportCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel expStack = new StackPanel();

            expStack.Children.Add(ThemeColors.CreateSectionHeader("Trích Xuất & Căn Chỉnh Hình Ảnh"));
            expStack.Children.Add(new TextBlock
            {
                Text = "Trích xuất toàn bộ hình ảnh nhúng trong tài liệu Word ra một thư mục riêng biệt với độ phân giải và chất lượng gốc:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Button btnExport = ThemeColors.CreatePrimaryButton(
                "📥 Xuất Tất Cả Ảnh Gốc Ra Thư Mục...",
                "Lưu toàn bộ ảnh trong file Word thành các tệp PNG riêng");
            btnExport.Margin = new Thickness(0, 0, 0, 6);
            btnExport.Click += (s, e) =>
            {
                string msg;
                _wordService.ExportAllImages(null, out msg);
                _notifyStatus(msg);
            };
            expStack.Children.Add(btnExport);

            Button btnCenter = ThemeColors.CreateSecondaryButton(
                "📐 Căn Giữa Tất Cả Hình Ảnh Trong Tài Liệu",
                "Định dạng căn giữa toàn bộ ảnh inline để trình bày đẹp mắt");
            btnCenter.Margin = new Thickness(0, 0, 0, 4);
            btnCenter.Click += (s, e) =>
            {
                string msg;
                _wordService.CenterAllImages(out msg);
                _notifyStatus(msg);
            };
            expStack.Children.Add(btnCenter);

            exportCard.Child = expStack;
            root.Children.Add(exportCard);

            // 2. Card: Nghiệp vụ Kế toán & Hành chính: Con Dấu Đỏ & Chữ Ký Số
            Border stampCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel stpStack = new StackPanel();

            stpStack.Children.Add(ThemeColors.CreateSectionHeader("Con Dấu Đỏ & Chữ Ký Số"));
            stpStack.Children.Add(new TextBlock
            {
                Text = "Chèn con dấu tròn công ty / chữ ký số, tự động tách nền trắng trong suốt để con dấu trùm lên chữ ký đúng thể thức hành chính & kế toán:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            _chkStampInFront = new CheckBox
            {
                Content = "Hiển thị con dấu nổi đè lên chữ ký (In front of text)",
                IsChecked = true,
                Margin = new Thickness(0, 2, 0, 6),
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5
            };
            stpStack.Children.Add(_chkStampInFront);

            Button btnStamp = ThemeColors.CreatePrimaryButton(
                "🔴 Chèn Con Dấu Đỏ / Chữ Ký Vào Vị Trí Con Trỏ...",
                "Chọn file ảnh con dấu để chèn chuẩn thể thức");
            btnStamp.Margin = new Thickness(0, 0, 0, 4);
            btnStamp.Click += (s, e) =>
            {
                string msg;
                bool inFront = _chkStampInFront.IsChecked == true;
                _wordService.InsertOfficialStampOrSignature(null, inFront, 130f, out msg);
                _notifyStatus(msg);
            };
            stpStack.Children.Add(btnStamp);

            stampCard.Child = stpStack;
            root.Children.Add(stampCard);

            // 3. Card: Chèn ảnh hàng loạt
            Border batchCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 4));
            StackPanel batStack = new StackPanel();

            batStack.Children.Add(ThemeColors.CreateSectionHeader("Chèn Ảnh Hàng Loạt"));
            batStack.Children.Add(new TextBlock
            {
                Text = "Chọn nhiều ảnh cùng lúc và tự động chèn vào tài liệu Word theo kích thước chuẩn thống nhất, căn giữa đẹp mắt:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Button btnBatchInsert = ThemeColors.CreateSecondaryButton(
                "🖼 Chèn Nhiều Ảnh Hàng Loạt Vào Tài Liệu...",
                "Chọn nhiều file ảnh và chèn tự động");
            btnBatchInsert.Margin = new Thickness(0, 0, 0, 4);
            btnBatchInsert.Click += (s, e) =>
            {
                string msg;
                _wordService.BatchInsertImages(null, 360f, out msg);
                _notifyStatus(msg);
            };
            batStack.Children.Add(btnBatchInsert);

            batchCard.Child = batStack;
            root.Children.Add(batchCard);

            Content = root;
        }
    }
}
