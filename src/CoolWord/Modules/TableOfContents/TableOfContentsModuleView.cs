using System;
using System.Windows;
using System.Windows.Controls;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Modules.TableOfContents
{
    public class TableOfContentsModuleView : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Action<string> _notifyStatus;

        private ComboBox _cboTocType;
        private ComboBox _cboLeaderChar;
        private CheckBox _chkAtStartOfDoc;

        public TableOfContentsModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            _wordService = wordService;
            _notifyStatus = notifyStatus;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            ScrollViewer scroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
            };

            StackPanel panel = new StackPanel
            {
                Margin = new Thickness(12, 12, 12, 20)
            };

            // 1. Thẻ Chèn Mục Lục
            panel.Children.Add(BuildInsertTocCard());

            // 2. Thẻ Quét & Gán Heading Tự Động
            panel.Children.Add(BuildSmartHeadingCard());

            // 3. Thẻ Quản Lý & Cập Nhật
            panel.Children.Add(BuildManageTocCard());

            // 4. Thẻ Mẫu Tiêu Đề BCTC
            panel.Children.Add(BuildAccountingTemplateCard());

            scroll.Content = panel;
            Content = scroll;
        }

        private Border BuildInsertTocCard()
        {
            Border card = ThemeColors.CreateCard(new Thickness(14), new Thickness(0, 0, 0, 12));
            StackPanel sp = new StackPanel();

            sp.Children.Add(ThemeColors.CreateSectionHeader("Tạo Mục Lục Tự Động", "Chèn bảng mục lục liên kết theo cấp độ Heading hoặc danh mục"));

            sp.Children.Add(ThemeColors.CreateFieldLabel("Kiểu mục lục:"));
            _cboTocType = new ComboBox
            {
                Margin = new Thickness(0, 0, 0, 10),
                Height = 28,
                FontFamily = ThemeColors.MainFont,
                FontSize = 12
            };
            _cboTocType.Items.Add(new ComboBoxItem { Content = "Mục lục tiêu chuẩn (Heading 1 - 3 cấp)", IsSelected = true, Tag = "std3" });
            _cboTocType.Items.Add(new ComboBoxItem { Content = "Mục lục chi tiết mở rộng (Heading 1 - 5 cấp)", Tag = "std5" });
            _cboTocType.Items.Add(new ComboBoxItem { Content = "Mục lục Báo cáo tài chính & Kế toán chuẩn", Tag = "acc" });
            _cboTocType.Items.Add(new ComboBoxItem { Content = "Danh mục Bảng Biểu (List of Tables)", Tag = "tables" });
            _cboTocType.Items.Add(new ComboBoxItem { Content = "Danh mục Hình Ảnh & Sơ đồ (List of Figures)", Tag = "figures" });
            sp.Children.Add(_cboTocType);

            sp.Children.Add(ThemeColors.CreateFieldLabel("Đường dẫn số trang (Dấu chấm dẫn dòng):"));
            _cboLeaderChar = new ComboBox
            {
                Margin = new Thickness(0, 0, 0, 10),
                Height = 28,
                FontFamily = ThemeColors.MainFont,
                FontSize = 12
            };
            _cboLeaderChar.Items.Add(new ComboBoxItem { Content = "Dấu chấm . . . (Chuẩn phổ biến)", IsSelected = true, Tag = "dots" });
            _cboLeaderChar.Items.Add(new ComboBoxItem { Content = "Dấu gạch ngang - - -", Tag = "hyphens" });
            _cboLeaderChar.Items.Add(new ComboBoxItem { Content = "Dấu gạch dưới _ _ _", Tag = "lines" });
            _cboLeaderChar.Items.Add(new ComboBoxItem { Content = "Khoảng trắng (Không có ký tự dẫn)", Tag = "none" });
            sp.Children.Add(_cboLeaderChar);

            _chkAtStartOfDoc = new CheckBox
            {
                Content = "Chèn ở đầu tài liệu (Tạo trang bìa / trang đầu)",
                IsChecked = true,
                FontFamily = ThemeColors.MainFont,
                FontSize = 12,
                Margin = new Thickness(0, 0, 0, 12)
            };
            sp.Children.Add(_chkAtStartOfDoc);

            Button btnCreate = ThemeColors.CreatePrimaryButton("Thực Hiện Chèn Mục Lục", "Tạo và chèn bảng mục lục vào tài liệu Word");
            btnCreate.Height = 32;
            btnCreate.Click += OnCreateTocClick;
            sp.Children.Add(btnCreate);

            card.Child = sp;
            return card;
        }

        private Border BuildSmartHeadingCard()
        {
            Border card = ThemeColors.CreateCard(new Thickness(14), new Thickness(0, 0, 0, 12));
            StackPanel sp = new StackPanel();

            sp.Children.Add(ThemeColors.CreateSectionHeader("Quét & Gán Heading Thông Minh", "Tự động nhận diện tiêu đề văn bản"));

            TextBlock tbDesc = new TextBlock
            {
                Text = "Tự động quét toàn bộ văn bản để tìm các dòng tiêu đề như: Chương I, Phần 1, Điều 12, Khoản 3, Mục 1, 1.1, 1.2, 1.1.1, a), b) hoặc tiêu đề BCTC và gán Style Heading 1-2-3 tương ứng để tạo mục lục tự động.",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 10)
            };
            sp.Children.Add(tbDesc);

            Button btnAssign = ThemeColors.CreateAccentOutlineButton("Quét & Tự Động Gán Heading", "Phát hiện và áp dụng Style Heading 1, 2, 3 cho văn bản");
            btnAssign.Height = 32;
            btnAssign.Click += OnAutoAssignHeadingsClick;
            sp.Children.Add(btnAssign);

            card.Child = sp;
            return card;
        }

        private Border BuildManageTocCard()
        {
            Border card = ThemeColors.CreateCard(new Thickness(14), new Thickness(0, 0, 0, 12));
            StackPanel sp = new StackPanel();

            sp.Children.Add(ThemeColors.CreateSectionHeader("Quản Lý & Cập Nhật Mục Lục", "Làm mới hoặc định dạng lại bảng mục lục"));

            Grid gridButtons = new Grid { Margin = new Thickness(0, 0, 0, 6) };
            gridButtons.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            gridButtons.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(8, GridUnitType.Pixel) });
            gridButtons.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            gridButtons.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            gridButtons.RowDefinitions.Add(new RowDefinition { Height = new GridLength(8, GridUnitType.Pixel) });
            gridButtons.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            Button btnUpdateAll = ThemeColors.CreateSecondaryButton("Cập Nhật Toàn Bộ", "Làm mới nội dung và số trang");
            btnUpdateAll.Height = 30;
            btnUpdateAll.Click += (s, e) =>
            {
                if (_wordService == null) return;
                string status;
                _wordService.UpdateAllTablesOfContents(true, out status);
                Notify(status);
            };
            Grid.SetRow(btnUpdateAll, 0);
            Grid.SetColumn(btnUpdateAll, 0);
            gridButtons.Children.Add(btnUpdateAll);

            Button btnUpdatePageNum = ThemeColors.CreateSecondaryButton("Chỉ Cập Nhật Trang", "Làm mới lại số trang");
            btnUpdatePageNum.Height = 30;
            btnUpdatePageNum.Click += (s, e) =>
            {
                if (_wordService == null) return;
                string status;
                _wordService.UpdateAllTablesOfContents(false, out status);
                Notify(status);
            };
            Grid.SetRow(btnUpdatePageNum, 0);
            Grid.SetColumn(btnUpdatePageNum, 2);
            gridButtons.Children.Add(btnUpdatePageNum);

            Button btnFormatNĐ30 = ThemeColors.CreateSecondaryButton("Chuẩn Hóa NĐ 30", "Times New Roman 13pt, lề chuẩn");
            btnFormatNĐ30.Height = 30;
            btnFormatNĐ30.Click += (s, e) =>
            {
                if (_wordService == null) return;
                string status;
                _wordService.FormatTableOfContentsDecree30(out status);
                Notify(status);
            };
            Grid.SetRow(btnFormatNĐ30, 2);
            Grid.SetColumn(btnFormatNĐ30, 0);
            gridButtons.Children.Add(btnFormatNĐ30);

            Button btnDelete = ThemeColors.CreateSecondaryButton("Xóa Bỏ Mục Lục", "Xóa các bảng mục lục hiện có");
            btnDelete.Height = 30;
            btnDelete.Click += (s, e) =>
            {
                if (_wordService == null) return;
                string status;
                _wordService.DeleteAllTablesOfContents(out status);
                Notify(status);
            };
            Grid.SetRow(btnDelete, 2);
            Grid.SetColumn(btnDelete, 2);
            gridButtons.Children.Add(btnDelete);

            sp.Children.Add(gridButtons);
            card.Child = sp;
            return card;
        }

        private Border BuildAccountingTemplateCard()
        {
            Border card = ThemeColors.CreateCard(new Thickness(14), new Thickness(0, 0, 0, 0));
            StackPanel sp = new StackPanel();

            sp.Children.Add(ThemeColors.CreateSectionHeader("Hồ Sơ & Thuyết Minh BCTC", "Mẫu cấu trúc chuẩn báo cáo kiểm toán & kế toán"));

            TextBlock tbAcc = new TextBlock
            {
                Text = "Bao gồm đầy đủ 4 báo cáo tài chính bắt buộc (Cân đối kế toán, Kết quả kinh doanh, Lưu chuyển tiền tệ, Thuyết minh BCTC) cùng danh mục chứng từ sổ sách kế toán.",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 10)
            };
            sp.Children.Add(tbAcc);

            Button btnAccToc = ThemeColors.CreateSecondaryButton("Tạo Mục Lục Kế Toán Chuẩn Ngay", "Chèn mục lục chuẩn Báo cáo tài chính doanh nghiệp");
            btnAccToc.Height = 30;
            btnAccToc.Click += (s, e) =>
            {
                if (_wordService == null) return;
                string status;
                _wordService.InsertAccountingFinancialToc(_chkAtStartOfDoc != null && _chkAtStartOfDoc.IsChecked == true, out status);
                Notify(status);
            };
            sp.Children.Add(btnAccToc);

            card.Child = sp;
            return card;
        }

        private void OnCreateTocClick(object sender, RoutedEventArgs e)
        {
            if (_wordService == null) return;

            string typeTag = "std3";
            ComboBoxItem item = _cboTocType != null ? _cboTocType.SelectedItem as ComboBoxItem : null;
            if (item != null && item.Tag != null) typeTag = item.Tag.ToString();

            bool atStart = _chkAtStartOfDoc != null && _chkAtStartOfDoc.IsChecked == true;
            string status;

            if (typeTag == "std3")
            {
                _wordService.InsertTableOfContents(3, atStart, out status);
            }
            else if (typeTag == "std5")
            {
                _wordService.InsertTableOfContents(5, atStart, out status);
            }
            else if (typeTag == "acc")
            {
                _wordService.InsertAccountingFinancialToc(atStart, out status);
            }
            else if (typeTag == "tables")
            {
                _wordService.InsertTableOfFigures(true, atStart, out status);
            }
            else if (typeTag == "figures")
            {
                _wordService.InsertTableOfFigures(false, atStart, out status);
            }
            else
            {
                _wordService.InsertTableOfContents(3, atStart, out status);
            }

            Notify(status);
        }

        private void OnAutoAssignHeadingsClick(object sender, RoutedEventArgs e)
        {
            if (_wordService == null) return;
            int count;
            string status;
            _wordService.AutoAssignHeadings(out count, out status);
            Notify(status);
        }

        private void Notify(string message)
        {
            if (_notifyStatus != null)
            {
                _notifyStatus(message);
            }
        }
    }
}
