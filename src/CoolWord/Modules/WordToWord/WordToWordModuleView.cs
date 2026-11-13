using System;
using System.Windows;
using System.Windows.Controls;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Modules.WordToWord
{
    public class WordToWordModuleView : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Action<string> _notifyStatus;

        private CheckBox _chkSelectionOnly;

        public WordToWordModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            _wordService = wordService;
            _notifyStatus = notifyStatus ?? (_ => { });

            BuildUi();
        }

        private void BuildUi()
        {
            StackPanel root = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };

            // 1. Card: Chuyển Mã TCVN3 / VNI sang Unicode
            Border encCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel encStack = new StackPanel();

            encStack.Children.Add(ThemeColors.CreateSectionHeader("Chuyển Bảng Mã Cũ Sang Unicode"));
            encStack.Children.Add(new TextBlock
            {
                Text = "Khắc phục triệt để lỗi mất dấu phông chữ cũ (.VnTime, .VnArial, VNI-Times), tự động đổi bảng mã sang Unicode dựng sẵn và đổi phông sang Times New Roman:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Button btnAutoUnicode = ThemeColors.CreatePrimaryButton(
                "🔤 Chuyển Sang Unicode Chuẩn (Tự Động)",
                "Nhận diện và chuyển đổi bảng mã sang Unicode Times New Roman");
            btnAutoUnicode.Margin = new Thickness(0, 0, 0, 6);
            btnAutoUnicode.Click += (s, e) =>
            {
                string msg;
                bool selOnly = _chkSelectionOnly.IsChecked == true;
                _wordService.ConvertEncodingToUnicode(selOnly, "AUTO", out msg);
                _notifyStatus(msg);
            };
            encStack.Children.Add(btnAutoUnicode);

            Grid encSubGrid = new Grid { Margin = new Thickness(0, 0, 0, 6) };
            encSubGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            encSubGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            encSubGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Button btnTcvn = ThemeColors.CreateSecondaryButton("TCVN3 (.VnTime)", "Chuyển mã TCVN3 sang Unicode");
            btnTcvn.Click += (s, e) =>
            {
                string msg;
                bool selOnly = _chkSelectionOnly.IsChecked == true;
                _wordService.ConvertEncodingToUnicode(selOnly, "TCVN3", out msg);
                _notifyStatus(msg);
            };
            Grid.SetColumn(btnTcvn, 0);
            encSubGrid.Children.Add(btnTcvn);

            Button btnVni = ThemeColors.CreateSecondaryButton("VNI (VNI-Times)", "Chuyển mã VNI sang Unicode");
            btnVni.Click += (s, e) =>
            {
                string msg;
                bool selOnly = _chkSelectionOnly.IsChecked == true;
                _wordService.ConvertEncodingToUnicode(selOnly, "VNI", out msg);
                _notifyStatus(msg);
            };
            Grid.SetColumn(btnVni, 2);
            encSubGrid.Children.Add(btnVni);

            encStack.Children.Add(encSubGrid);

            _chkSelectionOnly = new CheckBox
            {
                IsChecked = false,
                Content = new TextBlock
                {
                    Text = "Chỉ áp dụng cho vùng văn bản đang bôi đen (Selection)",
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 11.5,
                    Foreground = ThemeColors.TextPrimary
                }
            };
            encStack.Children.Add(_chkSelectionOnly);

            encCard.Child = encStack;
            root.Children.Add(encCard);

            // 2. Card: Chuyển Đổi Định Dạng Tệp .doc <-> .docx
            Border fmtCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel fmtStack = new StackPanel();

            fmtStack.Children.Add(ThemeColors.CreateSectionHeader("Chuyển Đổi Định Dạng Tệp (.doc ↔ .docx)"));

            Button btnDocx = ThemeColors.CreatePrimaryButton(
                "💾 Nâng Cấp Sang Chuẩn .docx (Word Hiện Đại)",
                "Chuyển đổi file .doc cũ sang định dạng OpenXML (.docx) dung lượng nhẹ hơn");
            btnDocx.Margin = new Thickness(0, 0, 0, 6);
            btnDocx.Click += (s, e) =>
            {
                string msg;
                _wordService.SaveAsDocx(out msg);
                _notifyStatus(msg);
            };
            fmtStack.Children.Add(btnDocx);

            Button btnDoc97 = ThemeColors.CreateSecondaryButton(
                "💾 Lưu Bản Sao .doc (Tương Thích Word 97-2003)",
                "Lưu một bản sao định dạng .doc để gửi cho hệ thống phần mềm quản lý cũ");
            btnDoc97.Click += (s, e) =>
            {
                string msg;
                _wordService.SaveAsDoc97(out msg);
                _notifyStatus(msg);
            };
            fmtStack.Children.Add(btnDoc97);

            fmtCard.Child = fmtStack;
            root.Children.Add(fmtCard);

            // 3. Card: Gộp & Tách File Word
            Border mergeCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 4));
            StackPanel mergeStack = new StackPanel();

            mergeStack.Children.Add(ThemeColors.CreateSectionHeader("Gộp & Tách Tài Liệu Word"));

            Button btnMerge = ThemeColors.CreateSecondaryButton(
                "📑 Gộp Nhiều File Word Thành 1 File Duy Nhất...",
                "Chọn danh sách các file Word và nối liên tiếp vào tài liệu (có ngắt trang)");
            btnMerge.Margin = new Thickness(0, 0, 0, 6);
            btnMerge.Click += (s, e) =>
            {
                string msg;
                _wordService.MergeWordDocuments(null, out msg);
                _notifyStatus(msg);
            };
            mergeStack.Children.Add(btnMerge);

            Grid splitGrid = new Grid();
            splitGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            splitGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            splitGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Button btnSplitSec = ThemeColors.CreateAccentOutlineButton(
                "Tách Theo Section",
                "Tách mỗi Section thành 1 file Word riêng");
            btnSplitSec.Click += (s, e) =>
            {
                string msg;
                _wordService.SplitDocumentBySections(null, out msg);
                _notifyStatus(msg);
            };
            Grid.SetColumn(btnSplitSec, 0);
            splitGrid.Children.Add(btnSplitSec);

            Button btnSplitPage = ThemeColors.CreateAccentOutlineButton(
                "Tách Theo Trang",
                "Tách mỗi trang thành 1 file Word riêng");
            btnSplitPage.Click += (s, e) =>
            {
                string msg;
                _wordService.SplitDocumentByPages(null, out msg);
                _notifyStatus(msg);
            };
            Grid.SetColumn(btnSplitPage, 2);
            splitGrid.Children.Add(btnSplitPage);

            mergeStack.Children.Add(splitGrid);

            mergeCard.Child = mergeStack;
            root.Children.Add(mergeCard);

            Content = root;
        }
    }
}
