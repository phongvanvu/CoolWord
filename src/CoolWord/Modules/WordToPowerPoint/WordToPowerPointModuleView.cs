using System;
using System.Windows;
using System.Windows.Controls;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Modules.WordToPowerPoint
{
    public class WordToPowerPointModuleView : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Action<string> _notifyStatus;

        private CheckBox _chkOpenAfterExport;

        public WordToPowerPointModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            _wordService = wordService;
            _notifyStatus = notifyStatus ?? (_ => { });

            BuildUi();
        }

        private void BuildUi()
        {
            StackPanel root = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };

            // 1. Card: Tạo Slide từ Dàn Ý / Đề Mục Heading
            Border headingCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel headingStack = new StackPanel();

            headingStack.Children.Add(ThemeColors.CreateSectionHeader("Tạo Slide Từ Dàn Ý & Đề Mục"));
            headingStack.Children.Add(new TextBlock
            {
                Text = "Tự động phân tích tài liệu: Heading 1 làm Tiêu đề slide, Heading 2-3 và các ý chính làm gạch đầu dòng (Bullets) chuyên nghiệp:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Button btnHeadings = ThemeColors.CreatePrimaryButton(
                "🎯 Tạo Slide Từ Đề Mục & Dàn Ý (Khuyên Dùng)",
                "Chuyển đổi các cấp đề mục trong văn bản thành các trang trình chiếu hoàn chỉnh");
            btnHeadings.Margin = new Thickness(0, 0, 0, 6);
            btnHeadings.Click += (s, e) =>
            {
                string msg;
                bool open = _chkOpenAfterExport.IsChecked == true;
                _wordService.CreateSlidesFromHeadings(null, open, out msg);
                _notifyStatus(msg);
            };
            headingStack.Children.Add(btnHeadings);

            headingCard.Child = headingStack;
            root.Children.Add(headingCard);

            // 2. Card: Tạo Slide theo Trang & Section
            Border pageCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel pageStack = new StackPanel();

            pageStack.Children.Add(ThemeColors.CreateSectionHeader("Tạo Slide Theo Trang & Section"));
            pageStack.Children.Add(new TextBlock
            {
                Text = "Chuyển đổi linh hoạt theo cấu trúc phân trang hoặc phân đoạn của tài liệu Word:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Grid splitGrid = new Grid { Margin = new Thickness(0, 0, 0, 6) };
            splitGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            splitGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            splitGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Button btnPages = ThemeColors.CreateSecondaryButton("📄 Theo Từng Trang", "Mỗi trang Word thành 1 slide");
            btnPages.Click += (s, e) =>
            {
                string msg;
                bool open = _chkOpenAfterExport.IsChecked == true;
                _wordService.CreateSlidesFromPages(null, open, out msg);
                _notifyStatus(msg);
            };
            Grid.SetColumn(btnPages, 0);
            splitGrid.Children.Add(btnPages);

            Button btnSections = ThemeColors.CreateSecondaryButton("📑 Theo Từng Section", "Mỗi Section thành nhóm slide");
            btnSections.Click += (s, e) =>
            {
                string msg;
                bool open = _chkOpenAfterExport.IsChecked == true;
                _wordService.CreateSlidesFromSections(null, open, out msg);
                _notifyStatus(msg);
            };
            Grid.SetColumn(btnSections, 2);
            splitGrid.Children.Add(btnSections);

            pageStack.Children.Add(splitGrid);
            pageCard.Child = pageStack;
            root.Children.Add(pageCard);

            // 3. Card: Bảng Biểu & Báo Cáo Tài Chính
            Border tableCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel tableStack = new StackPanel();

            tableStack.Children.Add(ThemeColors.CreateSectionHeader("Bảng Biểu & Báo Cáo Tài Chính"));
            tableStack.Children.Add(new TextBlock
            {
                Text = "Xuất toàn bộ bảng dữ liệu sang Slide hoặc tự động tạo bộ Slide thuyết trình Báo Cáo Tài Chính hoàn chỉnh:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Button btnTables = ThemeColors.CreateSecondaryButton(
                "📊 Xuất Toàn Bộ Bảng Biểu Sang Slide",
                "Mỗi bảng dữ liệu được đặt căn giữa trên một slide riêng");
            btnTables.Margin = new Thickness(0, 0, 0, 6);
            btnTables.Click += (s, e) =>
            {
                string msg;
                bool open = _chkOpenAfterExport.IsChecked == true;
                _wordService.ExportTablesToSlides(null, open, out msg);
                _notifyStatus(msg);
            };
            tableStack.Children.Add(btnTables);

            Button btnFinPpt = ThemeColors.CreatePrimaryButton(
                "💼 Tạo Bộ Slide Báo Cáo Tài Chính & Tổng Kết",
                "Tạo bài thuyết trình tổng quan gồm bìa, căn cứ, số liệu và kết luận");
            btnFinPpt.Margin = new Thickness(0, 0, 0, 6);
            btnFinPpt.Click += (s, e) =>
            {
                string msg;
                bool open = _chkOpenAfterExport.IsChecked == true;
                _wordService.CreateFinancialPresentation(null, open, out msg);
                _notifyStatus(msg);
            };
            tableStack.Children.Add(btnFinPpt);

            tableCard.Child = tableStack;
            root.Children.Add(tableCard);

            // 4. Card: Cấu hình trình chiếu
            Border optCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 4));
            StackPanel optStack = new StackPanel();

            optStack.Children.Add(ThemeColors.CreateSectionHeader("Tùy Chọn Trình Chiếu"));

            _chkOpenAfterExport = new CheckBox
            {
                Content = "Mở Microsoft PowerPoint ngay sau khi tạo xong",
                IsChecked = true,
                Margin = new Thickness(0, 4, 0, 4),
                FontFamily = ThemeColors.MainFont,
                FontSize = 12
            };
            optStack.Children.Add(_chkOpenAfterExport);

            optCard.Child = optStack;
            root.Children.Add(optCard);

            Content = root;
        }
    }
}
