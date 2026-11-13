using System;
using System.Windows;
using System.Windows.Controls;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Modules.WordToExcel
{
    public class WordToExcelModuleView : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Action<string> _notifyStatus;

        private CheckBox _chkOpenAfterExport;
        private CheckBox _chkSeparateSheets;

        public WordToExcelModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            _wordService = wordService;
            _notifyStatus = notifyStatus ?? (_ => { });

            BuildUi();
        }

        private void BuildUi()
        {
            StackPanel root = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };

            // 1. Card: Xuất Bảng Biểu
            Border tableCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel tableStack = new StackPanel();

            tableStack.Children.Add(ThemeColors.CreateSectionHeader("Xuất Bảng Biểu Word ➔ Excel"));
            tableStack.Children.Add(new TextBlock
            {
                Text = "Trích xuất toàn bộ bảng biểu từ Word sang bảng tính Microsoft Excel (.xlsx), giữ nguyên cấu trúc dòng cột và tự động căn chỉnh ô:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Button btnExportAll = ThemeColors.CreatePrimaryButton(
                "📊 Xuất Tất Cả Bảng Biểu Sang Excel",
                "Xuất toàn bộ bảng trong tài liệu sang file Excel mới");
            btnExportAll.Margin = new Thickness(0, 0, 0, 6);
            btnExportAll.Click += (s, e) =>
            {
                string msg;
                bool open = _chkOpenAfterExport.IsChecked == true;
                bool separate = _chkSeparateSheets.IsChecked == true;
                _wordService.ExportAllTablesToExcel(null, open, separate, out msg);
                _notifyStatus(msg);
            };
            tableStack.Children.Add(btnExportAll);

            Button btnExportSel = ThemeColors.CreateSecondaryButton(
                "🔍 Xuất Bảng Đang Chọn Sang Excel",
                "Chỉ xuất bảng biểu nơi con trỏ đang đứng");
            btnExportSel.Margin = new Thickness(0, 0, 0, 6);
            btnExportSel.Click += (s, e) =>
            {
                string msg;
                bool open = _chkOpenAfterExport.IsChecked == true;
                _wordService.ExportSelectedTableToExcel(null, open, out msg);
                _notifyStatus(msg);
            };
            tableStack.Children.Add(btnExportSel);

            tableCard.Child = tableStack;
            root.Children.Add(tableCard);

            // 2. Card: Nghiệp Vụ Kế Toán (Bảng Kê & Báo Cáo Tài Chính)
            Border accCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel accStack = new StackPanel();

            accStack.Children.Add(ThemeColors.CreateSectionHeader("Nghiệp Vụ Kế Toán: Bảng Kê & Báo Cáo"));
            accStack.Children.Add(new TextBlock
            {
                Text = "Tự động nhận diện các cột tiền tệ (Số tiền, Thành tiền, Đơn giá...), chuyển sang định dạng số kế toán (#,##0) và chèn hàng TỔNG CỘNG với công thức =SUM(...) chuẩn Excel:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Button btnExportAcc = ThemeColors.CreatePrimaryButton(
                "💰 Xuất Bảng Kế Toán / Bảng Kê (Kèm Công Thức SUM)",
                "Tự động định dạng số tiền và chèn hàng Tổng cộng có công thức SUM");
            btnExportAcc.Margin = new Thickness(0, 0, 0, 6);
            btnExportAcc.Click += (s, e) =>
            {
                string msg;
                bool open = _chkOpenAfterExport.IsChecked == true;
                _wordService.ExportAccountingTableToExcel(null, open, out msg);
                _notifyStatus(msg);
            };
            accStack.Children.Add(btnExportAcc);

            accCard.Child = accStack;
            root.Children.Add(accCard);

            // 3. Card: Xuất Dòng Văn Bản
            Border listCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel listStack = new StackPanel();

            listStack.Children.Add(ThemeColors.CreateSectionHeader("Xuất Danh Sách Đoạn Văn"));
            listStack.Children.Add(new TextBlock
            {
                Text = "Chuyển từng đoạn văn bản hoặc các mục gạch đầu dòng từ Word thành các dòng trong cột Excel (kèm cột STT tự động):",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Button btnExportLines = ThemeColors.CreateSecondaryButton(
                "📝 Xuất Đoạn Văn / Danh Mục Sang Dòng Excel",
                "Chuyển các đoạn văn hoặc mục bôi đen thành các dòng trong Excel");
            btnExportLines.Margin = new Thickness(0, 0, 0, 6);
            btnExportLines.Click += (s, e) =>
            {
                string msg;
                bool open = _chkOpenAfterExport.IsChecked == true;
                _wordService.ExportParagraphsToExcel(null, open, out msg);
                _notifyStatus(msg);
            };
            listStack.Children.Add(btnExportLines);

            listCard.Child = listStack;
            root.Children.Add(listCard);

            // 4. Card: Cấu Hình Xuất Excel
            Border optCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 4));
            StackPanel optStack = new StackPanel();

            optStack.Children.Add(ThemeColors.CreateSectionHeader("Tùy Chọn Bảng Tính"));

            _chkOpenAfterExport = new CheckBox
            {
                Content = "Mở file Excel ngay sau khi xuất xong",
                IsChecked = true,
                Margin = new Thickness(0, 4, 0, 4),
                FontFamily = ThemeColors.MainFont,
                FontSize = 12
            };
            optStack.Children.Add(_chkOpenAfterExport);

            _chkSeparateSheets = new CheckBox
            {
                Content = "Nếu có nhiều bảng: Mỗi bảng đặt trên một Sheet riêng biệt",
                IsChecked = false,
                Margin = new Thickness(0, 4, 0, 4),
                FontFamily = ThemeColors.MainFont,
                FontSize = 12
            };
            optStack.Children.Add(_chkSeparateSheets);

            optCard.Child = optStack;
            root.Children.Add(optCard);

            Content = root;
        }
    }
}
