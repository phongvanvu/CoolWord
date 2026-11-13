using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Modules.TextUtilities
{
    public class TextCleanModule : IToolModule
    {
        public string Id { get { return "text-utilities"; } }
        public string Title { get { return "Chuẩn Hóa Văn Bản"; } }
        public string Subtitle { get { return "Xóa khoảng trắng thừa, chuẩn hóa dấu câu & đổi kiểu chữ Hoa/Thường"; } }
        public string IconText { get { return "Aa"; } }
        public string Category { get { return "Soạn thảo & Định dạng"; } }
        public int Order { get { return 3; } }

        public bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)
        {
            return wordService.TransformSelectedText(NormalizeVietnameseText, "Chuẩn hóa văn bản Tiếng Việt", out statusMessage);
        }

        public UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            return new TextCleanModuleView(wordService, notifyStatus);
        }

        public static string NormalizeVietnameseText(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return input;

            string s = input.Trim();
            // Gộp nhiều khoảng trắng liên tiếp thành 1 khoảng trắng
            s = Regex.Replace(s, @"[ \t]+", " ");
            // Xóa khoảng trắng thừa trước dấu câu (, . ; : ! ? ) ])
            s = Regex.Replace(s, @"\s+([,.;:!?\)\]])", "$1");
            // Xóa khoảng trắng thừa ngay sau dấu mở ngoặc ( [
            s = Regex.Replace(s, @"([\(\[])\s+", "$1");
            // Đảm bảo có khoảng trắng sau dấu phẩy, chấm, chấm phẩy nếu theo sau là chữ cái
            s = Regex.Replace(s, @"([,.;:!?])([^\s\d,.;:!?\)\]])", "$1 $2");
            return s;
        }

        public static string ToTitleCaseVi(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return input;
            CultureInfo vi = new CultureInfo("vi-VN");
            return vi.TextInfo.ToTitleCase(input.ToLower(vi));
        }

        public static string ToSentenceCaseVi(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return input;
            CultureInfo vi = new CultureInfo("vi-VN");
            string lower = input.ToLower(vi);
            return Regex.Replace(lower, @"(^\s*\w|[.!?]\s+\w)", m => m.Value.ToUpper(vi));
        }
    }

    public class TextCleanModuleView : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Action<string> _notifyStatus;

        public TextCleanModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            _wordService = wordService;
            _notifyStatus = notifyStatus ?? (_ => { });

            BuildUi();
        }

        private void BuildUi()
        {
            StackPanel root = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };

            // Card 1: Chuẩn hóa văn bản
            Border cleanCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel cleanStack = new StackPanel();
            cleanStack.Children.Add(ThemeColors.CreateSectionHeader("Làm sạch & Chuẩn hóa dấu câu Tiếng Việt"));

            cleanStack.Children.Add(new TextBlock
            {
                Text = "Tự động loại bỏ khoảng trắng kép, xóa dấu cách thừa trước dấu chấm/phẩy và chuẩn hóa khoảng cách sau dấu câu trên vùng văn bản đang chọn.",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 10)
            });

            Button btnNormalize = ThemeColors.CreatePrimaryButton(
                "✨ Chuẩn Hóa Đoạn Văn Đang Chọn",
                "Xử lý khoảng trắng và dấu câu chuẩn thể thức văn bản hành chính");
            btnNormalize.Click += (s, e) =>
            {
                string msg;
                _wordService.TransformSelectedText(TextCleanModule.NormalizeVietnameseText, "Chuẩn hóa dấu câu & khoảng trắng", out msg);
                _notifyStatus(msg);
            };
            cleanStack.Children.Add(btnNormalize);
            cleanCard.Child = cleanStack;
            root.Children.Add(cleanCard);

            // Card 2: Chuyển đổi Chữ Hoa / Chữ Thường Tiếng Việt
            Border caseCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel caseStack = new StackPanel();
            caseStack.Children.Add(ThemeColors.CreateSectionHeader("Đổi kiểu chữ Hoa / Thường (Hỗ trợ chuẩn Tiếng Việt)"));

            CultureInfo vi = new CultureInfo("vi-VN");

            Button btnTitleCase = ThemeColors.CreateSecondaryButton("Viết Hoa Đầu Mỗi Từ (Chuẩn Họ Tên / Địa Danh)");
            btnTitleCase.Margin = new Thickness(0, 0, 0, 6);
            btnTitleCase.Click += (s, e) => RunTransform(TextCleanModule.ToTitleCaseVi, "Viết Hoa Đầu Từ");
            caseStack.Children.Add(btnTitleCase);

            Button btnSentenceCase = ThemeColors.CreateSecondaryButton("Viết hoa chữ cái đầu mỗi câu");
            btnSentenceCase.Margin = new Thickness(0, 0, 0, 6);
            btnSentenceCase.Click += (s, e) => RunTransform(TextCleanModule.ToSentenceCaseVi, "Viết hoa đầu câu");
            caseStack.Children.Add(btnSentenceCase);

            Grid row = new Grid();
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Button btnUpper = ThemeColors.CreateSecondaryButton("CHỮ IN HOA TOÀN BỘ");
            btnUpper.Click += (s, e) => RunTransform(t => t.ToUpper(vi), "Đổi sang CHỮ IN HOA");
            Grid.SetColumn(btnUpper, 0);
            row.Children.Add(btnUpper);

            Button btnLower = ThemeColors.CreateSecondaryButton("chữ thường toàn bộ");
            btnLower.Click += (s, e) => RunTransform(t => t.ToLower(vi), "Đổi sang chữ thường");
            Grid.SetColumn(btnLower, 2);
            row.Children.Add(btnLower);

            caseStack.Children.Add(row);
            caseCard.Child = caseStack;
            root.Children.Add(caseCard);

            // Card 3: Thông tin mở rộng Module mới
            Border extCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0));
            StackPanel extStack = new StackPanel();
            extStack.Children.Add(ThemeColors.CreateSectionHeader("Kiến trúc Module Mở Rộng (Plugin Architecture)"));
            extStack.Children.Add(new TextBlock
            {
                Text = "CoolWord được thiết kế theo mô hình IToolModule & ModuleRegistry. Để thêm tính năng mới (ví dụ: Chuyển mã TCVN3/VNI sang Unicode, Tạo báo giá tự động, Kiểm tra chính tả...), bạn chỉ cần tạo một lớp thực thi IToolModule và đăng ký tại ModuleRegistry.cs.",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap
            });
            extCard.Child = extStack;
            root.Children.Add(extCard);

            Content = root;
        }

        private void RunTransform(Func<string, string> fn, string name)
        {
            string msg;
            _wordService.TransformSelectedText(fn, name, out msg);
            _notifyStatus(msg);
        }
    }
}
