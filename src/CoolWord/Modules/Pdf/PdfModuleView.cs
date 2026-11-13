using System;
using System.Windows;
using System.Windows.Controls;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Modules.Pdf
{
    public class PdfModuleView : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Action<string> _notifyStatus;

        private CheckBox _chkOpenAfter;
        private CheckBox _chkBookmarks;
        private TextBox _customWatermarkBox;

        public PdfModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            _wordService = wordService;
            _notifyStatus = notifyStatus ?? (_ => { });

            BuildUi();
        }

        private void BuildUi()
        {
            StackPanel root = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };

            // 1. Card: Chuẩn Lưu Trữ Kế Toán & Pháp Lý PDF/A (ISO 19005-1)
            Border accountingCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel accStack = new StackPanel();

            accStack.Children.Add(ThemeColors.CreateSectionHeader("Chuẩn Lưu Trữ Kế Toán PDF/A (ISO 19005-1)"));
            accStack.Children.Add(new TextBlock
            {
                Text = "Xuất Báo cáo tài chính, Sổ kế toán & Chứng từ điện tử chuẩn lưu trữ quốc tế ISO 19005-1. Đảm bảo tính toàn vẹn, nhúng 100% phông chữ, phục vụ quyết toán thuế và thanh tra lưu trữ dài hạn (10-20 năm):",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Button btnAccountingPdfA = ThemeColors.CreatePrimaryButton(
                "🏛️ Xuất PDF/A Chuẩn Lưu Trữ Kế Toán & Thuế (1-Click)",
                "Áp dụng chuẩn ISO 19005-1 nhúng đầy đủ phông chữ và dữ liệu toàn vẹn cho hồ sơ kế toán");
            btnAccountingPdfA.Margin = new Thickness(0, 0, 0, 8);
            btnAccountingPdfA.Click += (s, e) =>
            {
                string msg;
                _wordService.ExportAccountingPdfA(null, out msg);
                _notifyStatus(msg);
            };
            accStack.Children.Add(btnAccountingPdfA);

            _chkOpenAfter = new CheckBox
            {
                IsChecked = true,
                Margin = new Thickness(0, 0, 0, 4),
                Content = new TextBlock
                {
                    Text = "Tự động mở file PDF sau khi xuất thành công",
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 11.5,
                    Foreground = ThemeColors.TextPrimary
                }
            };
            accStack.Children.Add(_chkOpenAfter);

            _chkBookmarks = new CheckBox
            {
                IsChecked = true,
                Margin = new Thickness(0, 0, 0, 4),
                Content = new TextBlock
                {
                    Text = "Tạo mục lục Bookmarks từ các đề mục Heading",
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 11.5,
                    Foreground = ThemeColors.TextPrimary
                }
            };
            accStack.Children.Add(_chkBookmarks);

            accountingCard.Child = accStack;
            root.Children.Add(accountingCard);

            // 2. Card: Xuất Nhanh Theo Phạm Vi & Tối Ưu Dung Lượng
            Border scopeCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel scopeStack = new StackPanel();

            scopeStack.Children.Add(ThemeColors.CreateSectionHeader("Xuất PDF Theo Vùng & Tối Ưu"));

            Grid scopeGrid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            scopeGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            scopeGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            scopeGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Button btnExportSelection = ThemeColors.CreateSecondaryButton(
                "📄 Xuất Vùng Đang Chọn",
                "Chỉ xuất đoạn văn bản hoặc bảng biểu đang được bôi đen");
            btnExportSelection.Click += (s, e) =>
            {
                string msg;
                _wordService.ExportSelectionToPdf(null, out msg);
                _notifyStatus(msg);
            };
            Grid.SetColumn(btnExportSelection, 0);
            scopeGrid.Children.Add(btnExportSelection);

            Button btnExportCurrentPage = ThemeColors.CreateSecondaryButton(
                "📑 Xuất Trang Này",
                "Chỉ xuất trang tài liệu hiện tại đang đặt con trỏ");
            btnExportCurrentPage.Click += (s, e) =>
            {
                string msg;
                _wordService.ExportCurrentPageToPdf(null, out msg);
                _notifyStatus(msg);
            };
            Grid.SetColumn(btnExportCurrentPage, 2);
            scopeGrid.Children.Add(btnExportCurrentPage);

            scopeStack.Children.Add(scopeGrid);

            Button btnExportOptimized = ThemeColors.CreateSecondaryButton(
                "🗜️ Xuất PDF Nén Nhỏ Gọn (Gửi Email / Zalo)",
                "Tối ưu hóa hình ảnh và cấu trúc tài liệu để có dung lượng nhẹ nhất gửi qua mạng");
            btnExportOptimized.Click += (s, e) =>
            {
                PdfOptions opts = new PdfOptions
                {
                    OptimizeForMinimumSize = true,
                    OpenAfterExport = _chkOpenAfter.IsChecked == true,
                    CreateBookmarksFromHeadings = _chkBookmarks.IsChecked == true
                };
                string msg;
                _wordService.ExportCustomPdf(opts, null, out msg);
                _notifyStatus(msg);
            };
            scopeStack.Children.Add(btnExportOptimized);

            scopeCard.Child = scopeStack;
            root.Children.Add(scopeCard);

            // 3. Card: Đóng Dấu Chìm Kế Toán & Bản Quyền (Watermark)
            Border wmCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 4));
            StackPanel wmStack = new StackPanel();

            wmStack.Children.Add(ThemeColors.CreateSectionHeader("Đóng Dấu Chìm Kế Toán (Watermark)"));
            wmStack.Children.Add(new TextBlock
            {
                Text = "Chèn dấu chìm mờ nghiêng 45° lên toàn bộ trang để phân biệt trạng thái chứng từ kế toán và bảo mật tài liệu:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            // Các mẫu dấu chìm kế toán phổ biến
            WrapPanel quickPills = new WrapPanel { Margin = new Thickness(0, 0, 0, 8) };
            string[] standardWm = new[] { "ĐÃ THANH TOÁN", "ĐÃ KIỂM TOÁN", "BẢN SAO KẾ TOÁN", "LƯU HÀNH NỘI BỘ", "DỰ THẢO" };
            foreach (string wm in standardWm)
            {
                string captured = wm;
                Button pillBtn = new Button
                {
                    Content = captured,
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = ThemeColors.WordBlue,
                    Background = ThemeColors.WordBlueLight,
                    BorderBrush = ThemeColors.WordBlueBorder,
                    BorderThickness = new Thickness(1),
                    Padding = new Thickness(8, 4, 8, 4),
                    Margin = new Thickness(0, 0, 6, 6),
                    Cursor = System.Windows.Input.Cursors.Hand
                };
                pillBtn.Click += (s, e) =>
                {
                    string msg;
                    _wordService.AddAccountingWatermark(captured, out msg);
                    _notifyStatus(msg);
                };
                quickPills.Children.Add(pillBtn);
            }
            wmStack.Children.Add(quickPills);

            // Tùy chỉnh dấu chìm
            wmStack.Children.Add(ThemeColors.CreateFieldLabel("Hoặc nhập nội dung dấu chìm tùy biến:"));
            Grid customGrid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            customGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            customGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            customGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            _customWatermarkBox = new TextBox
            {
                Text = "BẢN SAO LƯU TRỮ",
                FontFamily = ThemeColors.MainFont,
                FontSize = 12,
                Padding = new Thickness(8, 5, 8, 5),
                BorderBrush = ThemeColors.InputBorderBrush
            };
            Grid.SetColumn(_customWatermarkBox, 0);
            customGrid.Children.Add(_customWatermarkBox);

            Button btnApplyCustomWm = ThemeColors.CreateSecondaryButton("Áp Dụng", "Đóng dấu chìm tùy chỉnh");
            btnApplyCustomWm.Click += (s, e) =>
            {
                string msg;
                _wordService.AddAccountingWatermark(_customWatermarkBox.Text, out msg);
                _notifyStatus(msg);
            };
            Grid.SetColumn(btnApplyCustomWm, 2);
            customGrid.Children.Add(btnApplyCustomWm);
            wmStack.Children.Add(customGrid);

            Button btnRemoveWm = ThemeColors.CreateAccentOutlineButton("🗑️ Gỡ Bỏ Dấu Chìm Khỏi Tài Liệu", "Xóa toàn bộ dấu chìm đang có");
            btnRemoveWm.Click += (s, e) =>
            {
                string msg;
                _wordService.RemoveAccountingWatermark(out msg);
                _notifyStatus(msg);
            };
            wmStack.Children.Add(btnRemoveWm);

            wmCard.Child = wmStack;
            root.Children.Add(wmCard);

            Content = root;
        }
    }
}
