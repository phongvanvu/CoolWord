using System;
using System.Windows;
using System.Windows.Controls;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Modules.NameSplit
{
    public class NameSplitModuleView : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Action<string> _notifyStatus;

        private TextBox _testInputBox;
        private TextBlock _hoDemPreview;
        private TextBlock _tenPreview;
        private TextBlock _detailPreview;

        public NameSplitModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            _wordService = wordService;
            _notifyStatus = notifyStatus ?? (_ => { });

            BuildUi();
            UpdateLivePreview();
        }

        private void BuildUi()
        {
            StackPanel root = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };

            // 1. Card: Thao Tác Trong Bảng Biểu (Word Table)
            Border tableCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel tableStack = new StackPanel();

            tableStack.Children.Add(ThemeColors.CreateSectionHeader("Tách Họ - Tên Trong Bảng Biểu Word"));
            tableStack.Children.Add(new TextBlock
            {
                Text = "Đặt con trỏ chuột vào cột chứa \"Họ và tên\" trong bảng dữ liệu cần tách và chọn chế độ xử lý bên dưới:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Button btnSplit2 = ThemeColors.CreatePrimaryButton(
                "👥 Tách Thành 2 Cột: [Họ đệm] & [Tên]",
                "Thêm 1 cột mới bên phải, tách từ cuối cùng làm Tên để sắp xếp A-Z");
            btnSplit2.Margin = new Thickness(0, 0, 0, 6);
            btnSplit2.Click += (s, e) =>
            {
                string msg;
                _wordService.SplitTableColumn2(out msg);
                _notifyStatus(msg);
            };
            tableStack.Children.Add(btnSplit2);

            Grid actionGrid = new Grid { Margin = new Thickness(0, 0, 0, 6) };
            actionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            actionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            actionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Button btnSplit3 = ThemeColors.CreateSecondaryButton(
                "Tách 3 Cột: [Họ] [Đệm] [Tên]",
                "Tách chi tiết 3 cột riêng biệt");
            btnSplit3.Click += (s, e) =>
            {
                string msg;
                _wordService.SplitTableColumn3(out msg);
                _notifyStatus(msg);
            };
            Grid.SetColumn(btnSplit3, 0);
            actionGrid.Children.Add(btnSplit3);

            Button btnMerge = ThemeColors.CreateSecondaryButton(
                "Gộp 2 Cột ➔ [Họ và tên]",
                "Gộp cột Họ đệm và cột Tên liền kề thành 1 cột duy nhất");
            btnMerge.Click += (s, e) =>
            {
                string msg;
                _wordService.MergeTableColumns(out msg);
                _notifyStatus(msg);
            };
            Grid.SetColumn(btnMerge, 2);
            actionGrid.Children.Add(btnMerge);

            tableStack.Children.Add(actionGrid);

            tableCard.Child = tableStack;
            root.Children.Add(tableCard);

            // 2. Card: Thao Tác Trên Vùng Văn Bản Đang Bôi Đen (Ngoài Bảng)
            Border textCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel textStack = new StackPanel();

            textStack.Children.Add(ThemeColors.CreateSectionHeader("Xử Lý Danh Sách Văn Bản Bôi Đen"));

            Button btnSplitTextTab = ThemeColors.CreateSecondaryButton(
                "📑 Tách Họ - Tên Bằng Dấu Tab (Dán sang Excel)",
                "Tách từng dòng văn bản thành dạng: Họ đệm [Tab] Tên");
            btnSplitTextTab.Margin = new Thickness(0, 0, 0, 6);
            btnSplitTextTab.Click += (s, e) =>
            {
                string msg;
                _wordService.SplitSelectedTextLines("\t", out msg);
                _notifyStatus(msg);
            };
            textStack.Children.Add(btnSplitTextTab);

            Button btnNormalizeNames = ThemeColors.CreateAccentOutlineButton(
                "🔤 Chuẩn Hóa Viết Hoa Danh Từ Riêng (Họ Tên)",
                "Tự động viết hoa chuẩn Họ Tên (ví dụ: nguyễn văn an -> Nguyễn Văn An)");
            btnNormalizeNames.Click += (s, e) =>
            {
                string msg;
                _wordService.NormalizeSelectedNames(out msg);
                _notifyStatus(msg);
            };
            textStack.Children.Add(btnNormalizeNames);

            textCard.Child = textStack;
            root.Children.Add(textCard);

            // 3. Card: Thử Nghiệm Phân Tích Tên Trực Quan
            Border previewCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 4));
            StackPanel prevStack = new StackPanel();

            prevStack.Children.Add(ThemeColors.CreateSectionHeader("Kiểm Tra Nhanh Phân Tích Họ - Tên"));

            prevStack.Children.Add(ThemeColors.CreateFieldLabel("Nhập họ và tên bất kỳ để xem quy tắc tách:"));
            _testInputBox = new TextBox
            {
                Text = "Trần Thị Bích Ngọc",
                FontFamily = ThemeColors.MainFont,
                FontSize = 12.5,
                Padding = new Thickness(8, 5, 8, 5),
                BorderBrush = ThemeColors.InputBorderBrush,
                Margin = new Thickness(0, 0, 0, 8)
            };
            _testInputBox.TextChanged += (s, e) => UpdateLivePreview();
            prevStack.Children.Add(_testInputBox);

            Border resultBox = new Border
            {
                Background = ThemeColors.WordBlueLight,
                BorderBrush = ThemeColors.WordBlueBorder,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(10)
            };

            StackPanel resStack = new StackPanel();

            _hoDemPreview = new TextBlock
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 12,
                Foreground = ThemeColors.TextPrimary,
                Margin = new Thickness(0, 0, 0, 3)
            };
            resStack.Children.Add(_hoDemPreview);

            _tenPreview = new TextBlock
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 12,
                FontWeight = FontWeights.Bold,
                Foreground = ThemeColors.WordBlue,
                Margin = new Thickness(0, 0, 0, 3)
            };
            resStack.Children.Add(_tenPreview);

            _detailPreview = new TextBlock
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary
            };
            resStack.Children.Add(_detailPreview);

            resultBox.Child = resStack;
            prevStack.Children.Add(resultBox);

            previewCard.Child = prevStack;
            root.Children.Add(previewCard);

            Content = root;
        }

        private void UpdateLivePreview()
        {
            if (_testInputBox == null || _hoDemPreview == null) return;

            string raw = _testInputBox.Text;
            ParsedVietnameseName p = VietnameseNameSplitter.Split(raw);

            _hoDemPreview.Text = "• Họ và tên đệm: " + (string.IsNullOrEmpty(p.HoDem) ? "(Trống)" : p.HoDem);
            _tenPreview.Text = "• Tên chính (xếp A-Z): " + (string.IsNullOrEmpty(p.Ten) ? "(Trống)" : p.Ten);
            _detailPreview.Text = string.Format("• Tách 3 cột: Họ: [{0}] | Đệm: [{1}] | Tên: [{2}]",
                p.Ho, string.IsNullOrEmpty(p.Dem) ? "Không có" : p.Dem, p.Ten);
        }
    }
}
