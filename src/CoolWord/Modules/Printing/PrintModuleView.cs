using System;
using System.Windows;
using System.Windows.Controls;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Modules.Printing
{
    public class PrintModuleView : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Action<string> _notifyStatus;

        public PrintModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            _wordService = wordService;
            _notifyStatus = notifyStatus ?? (_ => { });

            BuildUi();
        }

        private void BuildUi()
        {
            StackPanel root = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };

            // 1. Card: In Nhanh Trang Hiện Tại & Vùng Chọn
            Border quickCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel quickStack = new StackPanel();

            quickStack.Children.Add(ThemeColors.CreateSectionHeader("In Nhanh Trang & Vùng Chọn"));
            quickStack.Children.Add(new TextBlock
            {
                Text = "Gửi lệnh in tức thì cho trang tài liệu đang thao tác hoặc đoạn văn bản được chọn mà không cần qua hộp thoại phức tạp:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Button btnPrintCurrent = ThemeColors.CreatePrimaryButton(
                "🖨️ In Nhanh Trang Hiện Tại",
                "Gửi lệnh in trang đang đặt con trỏ chuột");
            btnPrintCurrent.Margin = new Thickness(0, 0, 0, 6);
            btnPrintCurrent.Click += (s, e) =>
            {
                string msg;
                _wordService.PrintCurrentPage(out msg);
                _notifyStatus(msg);
            };
            quickStack.Children.Add(btnPrintCurrent);

            Button btnPrintSelection = ThemeColors.CreateSecondaryButton(
                "📄 In Vùng Văn Bản Đang Bôi Đen",
                "Chỉ in đoạn văn bản hoặc bảng biểu đang được bôi đen");
            btnPrintSelection.Click += (s, e) =>
            {
                string msg;
                _wordService.PrintSelection(out msg);
                _notifyStatus(msg);
            };
            quickStack.Children.Add(btnPrintSelection);

            quickCard.Child = quickStack;
            root.Children.Add(quickCard);

            // 2. Card: In 2 Mặt Thủ Công Thông Minh (Manual Duplex)
            Border duplexCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel duplexStack = new StackPanel();

            duplexStack.Children.Add(ThemeColors.CreateSectionHeader("In 2 Mặt Thủ Công (Khay Lật Giấy)"));
            duplexStack.Children.Add(new TextBlock
            {
                Text = "Giải pháp in 2 mặt cho máy in văn phòng 1 mặt: Tự động in toàn bộ trang lẻ trước, sau đó hướng dẫn lật mặt giấy để in tiếp trang chẵn.",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Button btnDuplex = ThemeColors.CreatePrimaryButton(
                "🔄 Bắt Đầu In 2 Mặt Thủ Công",
                "Quy trình in 2 mặt tự động có hướng dẫn đảo giấy");
            btnDuplex.Margin = new Thickness(0, 0, 0, 6);
            btnDuplex.Click += (s, e) =>
            {
                string msg;
                _wordService.PrintManualDuplex(out msg);
                _notifyStatus(msg);
            };
            duplexStack.Children.Add(btnDuplex);

            Grid oddEvenGrid = new Grid();
            oddEvenGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            oddEvenGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            oddEvenGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Button btnOdd = ThemeColors.CreateSecondaryButton("In Trang Lẻ (1, 3, 5...)", "Chỉ in các trang lẻ");
            btnOdd.Click += (s, e) =>
            {
                string msg;
                _wordService.PrintOddPagesOnly(out msg);
                _notifyStatus(msg);
            };
            Grid.SetColumn(btnOdd, 0);
            oddEvenGrid.Children.Add(btnOdd);

            Button btnEven = ThemeColors.CreateSecondaryButton("In Trang Chẵn (2, 4, 6...)", "Chỉ in các trang chẵn");
            btnEven.Click += (s, e) =>
            {
                string msg;
                _wordService.PrintEvenPagesOnly(out msg);
                _notifyStatus(msg);
            };
            Grid.SetColumn(btnEven, 2);
            oddEvenGrid.Children.Add(btnEven);

            duplexStack.Children.Add(oddEvenGrid);
            duplexCard.Child = duplexStack;
            root.Children.Add(duplexCard);

            // 3. Card: Tiết Kiệm Giấy & Xuất Bản PDF
            Border layoutCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel layoutStack = new StackPanel();

            layoutStack.Children.Add(ThemeColors.CreateSectionHeader("Tiết Kiệm Giấy & Xuất Bản PDF"));

            Button btnTwoPerSheet = ThemeColors.CreateSecondaryButton(
                "📰 In 2 Trang Trên 1 Mặt Giấy A4 (N-up)",
                "Co tỷ lệ in 2 trang tài liệu trên một mặt giấy A4 để tiết kiệm giấy và đọc duyệt nhanh");
            btnTwoPerSheet.Margin = new Thickness(0, 0, 0, 6);
            btnTwoPerSheet.Click += (s, e) =>
            {
                string msg;
                _wordService.PrintTwoPagesPerSheet(out msg);
                _notifyStatus(msg);
            };
            layoutStack.Children.Add(btnTwoPerSheet);

            Button btnExportPdf = ThemeColors.CreateAccentOutlineButton(
                "📑 Xuất Nhanh File PDF Chuẩn In Ấn",
                "Chuyển đổi văn bản sang định dạng PDF giữ nguyên phông và bố cục");
            btnExportPdf.Click += (s, e) =>
            {
                string msg;
                _wordService.ExportQuickPdf(null, out msg);
                _notifyStatus(msg);
            };
            layoutStack.Children.Add(btnExportPdf);

            layoutCard.Child = layoutStack;
            root.Children.Add(layoutCard);

            Content = root;
        }
    }
}
