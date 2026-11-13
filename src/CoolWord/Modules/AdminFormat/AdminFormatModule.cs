using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Modules.AdminFormat
{
    public class AdminFormatModule : IToolModule
    {
        public string Id { get { return "admin-decree30"; } }
        public string Title { get { return "Thể Thức NĐ 30"; } }
        public string Subtitle { get { return "Chuẩn hóa văn bản hành chính Nghị định 30/2020/NĐ-CP 1-Click"; } }
        public string IconText { get { return "NĐ30"; } }
        public string Category { get { return "Văn thư & Hành chính"; } }
        public int Order { get { return 2; } }

        public bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)
        {
            return wordService.ApplyDecree30FullStandard(new AdminDocOptions(), out statusMessage);
        }

        public UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            return new AdminFormatModuleView(wordService, notifyStatus);
        }
    }

    public class AdminFormatModuleView : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Action<string> _notifyStatus;
        private readonly AdminDocOptions _options = new AdminDocOptions();

        private ComboBox _marginPresetCombo;
        private ComboBox _fontSizeCombo;
        private CheckBox _chkPageNumber;

        private TextBox _parentOrgBox;
        private TextBox _issuingOrgBox;
        private TextBox _docNumberBox;
        private TextBox _locationBox;
        private ComboBox _docTypeCombo;
        private TextBox _subjectBox;

        private TextBox _authorityBox;
        private TextBox _positionBox;
        private TextBox _signerNameBox;
        private TextBox _recipientsBox;

        private TextBlock _dateAdminPreview;
        private TextBlock _dateWordsPreview;

        public AdminFormatModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            _wordService = wordService;
            _notifyStatus = notifyStatus ?? (_ => { });

            BuildUi();
            RefreshDatePreview();
        }

        private void BuildUi()
        {
            StackPanel root = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };

            // 1. Card: 1-Click Chuẩn hóa Thể thức Nghị định 30/2020/NĐ-CP
            Border oneClickCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel oneClickStack = new StackPanel();

            oneClickStack.Children.Add(ThemeColors.CreateSectionHeader("Chuẩn hóa văn bản hành chính 1-Click (NĐ 30/2020/NĐ-CP)"));
            oneClickStack.Children.Add(new TextBlock
            {
                Text = "Tự động thiết lập khổ giấy A4, căn lề chuẩn văn thư, chuẩn hóa phông Times New Roman, giãn dòng, thụt đầu dòng 1,27cm, dọn rác định dạng và đánh số trang chuẩn Nghị định 30/2020/NĐ-CP:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Button btnFullStandard = ThemeColors.CreatePrimaryButton(
                "🏛️ Chuẩn Hóa Toàn Diện NĐ 30/2020 (1-Click)",
                "Áp dụng toàn bộ quy chuẩn Phụ lục I - Nghị định 30/2020/NĐ-CP cho văn bản hiện tại");
            btnFullStandard.Margin = new Thickness(0, 0, 0, 6);
            btnFullStandard.Click += OnApplyFullStandardClick;
            oneClickStack.Children.Add(btnFullStandard);

            Grid quickRow = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            quickRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            quickRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            quickRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Button btnMarginsOnly = ThemeColors.CreateSecondaryButton("📐 Chỉ Căn Lề A4 & Trang", "Thiết lập khổ A4, căn lề chuẩn NĐ 30 và đánh số trang");
            btnMarginsOnly.Click += OnApplyMarginsOnlyClick;
            Grid.SetColumn(btnMarginsOnly, 0);
            quickRow.Children.Add(btnMarginsOnly);

            Button btnCleanGarbage = ThemeColors.CreateSecondaryButton("🧹 Dọn Rác Định Dạng", "Xóa khoảng trắng kép, tab đầu dòng, dấu cách thừa trước dấu câu & dòng trống thừa");
            btnCleanGarbage.Click += OnCleanGarbageClick;
            Grid.SetColumn(btnCleanGarbage, 2);
            quickRow.Children.Add(btnCleanGarbage);

            oneClickStack.Children.Add(quickRow);

            // Tùy chọn thông số lề & cỡ chữ
            oneClickStack.Children.Add(ThemeColors.CreateFieldLabel("Bộ thông số căn lề trang A4 (Phụ lục I - NĐ 30/2020):"));
            _marginPresetCombo = CreateComboBox(new[]
            {
                "Chuẩn tiêu chuẩn: Trên 20mm - Dưới 20mm - Trái 30mm - Phải 15mm",
                "Chuẩn lề trung bình: Trên 20mm - Dưới 20mm - Trái 30mm - Phải 20mm",
                "Chuẩn lề rộng: Trên 25mm - Dưới 25mm - Trái 35mm - Phải 20mm"
            });
            _marginPresetCombo.SelectedIndex = 0;
            oneClickStack.Children.Add(_marginPresetCombo);

            oneClickStack.Children.Add(ThemeColors.CreateFieldLabel("Phông chữ & Cỡ chữ nội dung văn bản:"));
            _fontSizeCombo = CreateComboBox(new[]
            {
                "Times New Roman - Cỡ 14 pt (Phổ biến nhất)",
                "Times New Roman - Cỡ 13 pt (Văn bản dài)"
            });
            _fontSizeCombo.SelectedIndex = 0;
            oneClickStack.Children.Add(_fontSizeCombo);

            _chkPageNumber = new CheckBox
            {
                IsChecked = true,
                Margin = new Thickness(0, 2, 0, 0),
                Content = new TextBlock
                {
                    Text = "Đánh số trang chuẩn NĐ 30 (Canh giữa lề trên, cỡ 13pt, ẩn ở trang 1)",
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 11.5,
                    Foreground = ThemeColors.TextPrimary,
                    TextWrapping = TextWrapping.Wrap
                }
            };
            oneClickStack.Children.Add(_chkPageNumber);

            oneClickCard.Child = oneClickStack;
            root.Children.Add(oneClickCard);

            // 2. Card: Tạo nhanh Khung Quốc hiệu - Tiêu ngữ & Tiêu đề văn bản
            Border headerCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel headerStack = new StackPanel();
            headerStack.Children.Add(ThemeColors.CreateSectionHeader("Tạo khung Quốc hiệu - Tiêu ngữ & Tiêu đề văn bản"));

            headerStack.Children.Add(ThemeColors.CreateFieldLabel("Cơ quan chủ quản cấp trên (để trống nếu không có):"));
            _parentOrgBox = CreateTextBox(_options.ParentOrgName);
            headerStack.Children.Add(_parentOrgBox);

            headerStack.Children.Add(ThemeColors.CreateFieldLabel("Tên cơ quan, tổ chức ban hành văn bản:"));
            _issuingOrgBox = CreateTextBox(_options.IssuingOrgName);
            headerStack.Children.Add(_issuingOrgBox);

            Grid numLocGrid = new Grid();
            numLocGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.1, GridUnitType.Star) });
            numLocGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(8) });
            numLocGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0.9, GridUnitType.Star) });

            StackPanel numPanel = new StackPanel();
            numPanel.Children.Add(ThemeColors.CreateFieldLabel("Số, ký hiệu:"));
            _docNumberBox = CreateTextBox(_options.DocumentNumber);
            numPanel.Children.Add(_docNumberBox);
            Grid.SetColumn(numPanel, 0);
            numLocGrid.Children.Add(numPanel);

            StackPanel locPanel = new StackPanel();
            locPanel.Children.Add(ThemeColors.CreateFieldLabel("Địa danh ban hành:"));
            _locationBox = CreateTextBox(_options.LocationName);
            _locationBox.TextChanged += (s, e) => RefreshDatePreview();
            locPanel.Children.Add(_locationBox);
            Grid.SetColumn(locPanel, 2);
            numLocGrid.Children.Add(locPanel);

            headerStack.Children.Add(numLocGrid);

            headerStack.Children.Add(ThemeColors.CreateFieldLabel("Tên loại văn bản:"));
            _docTypeCombo = CreateComboBox(new[]
            {
                "TỜ TRÌNH",
                "QUYẾT ĐỊNH",
                "BÁO CÁO",
                "THÔNG BÁO",
                "KẾ HOẠCH",
                "CHƯƠNG TRÌNH",
                "BIÊN BẢN",
                "(Công văn - Không có dòng tên loại văn bản)"
            });
            _docTypeCombo.SelectedIndex = 0;
            headerStack.Children.Add(_docTypeCombo);

            headerStack.Children.Add(ThemeColors.CreateFieldLabel("Trích yếu nội dung văn bản:"));
            _subjectBox = CreateTextBox(_options.DocumentSubject);
            headerStack.Children.Add(_subjectBox);

            Button btnInsertHeader = ThemeColors.CreateAccentOutlineButton(
                "➕ Chèn Khung Quốc Hiệu & Tiêu Đề Vào Đầu Trang",
                "Tạo bảng ẩn viền 2x2 chuẩn Nghị định 30/2020/NĐ-CP ở đầu tài liệu Word");
            btnInsertHeader.Margin = new Thickness(0, 4, 0, 0);
            btnInsertHeader.Click += OnInsertHeaderClick;
            headerStack.Children.Add(btnInsertHeader);

            headerCard.Child = headerStack;
            root.Children.Add(headerCard);

            // 3. Card: Tạo nhanh Khung Nơi nhận & Chữ ký cuối văn bản
            Border footerCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel footerStack = new StackPanel();
            footerStack.Children.Add(ThemeColors.CreateSectionHeader("Tạo khung Nơi nhận & Chữ ký cuối văn bản"));

            Grid signGrid = new Grid();
            signGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            signGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(8) });
            signGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            StackPanel authPanel = new StackPanel();
            authPanel.Children.Add(ThemeColors.CreateFieldLabel("Quyền hạn (TM./KT./TL.):"));
            _authorityBox = CreateTextBox(_options.AuthorityPrefix);
            _authorityBox.ToolTip = "Ví dụ: TM. ỦY BAN NHÂN DÂN hoặc KT. GIÁM ĐỐC (để trống nếu ký trực tiếp)";
            authPanel.Children.Add(_authorityBox);
            Grid.SetColumn(authPanel, 0);
            signGrid.Children.Add(authPanel);

            StackPanel posPanel = new StackPanel();
            posPanel.Children.Add(ThemeColors.CreateFieldLabel("Chức vụ người ký:"));
            _positionBox = CreateTextBox(_options.SignerPosition);
            posPanel.Children.Add(_positionBox);
            Grid.SetColumn(posPanel, 2);
            signGrid.Children.Add(posPanel);

            footerStack.Children.Add(signGrid);

            footerStack.Children.Add(ThemeColors.CreateFieldLabel("Họ và tên người ký:"));
            _signerNameBox = CreateTextBox(_options.SignerFullName);
            footerStack.Children.Add(_signerNameBox);

            footerStack.Children.Add(ThemeColors.CreateFieldLabel("Danh sách Nơi nhận (mỗi dòng 1 nơi nhận):"));
            _recipientsBox = new TextBox
            {
                Text = "- Như trên;\r\n- Ban Giám đốc;\r\n- Lưu: VT, KH.",
                FontFamily = ThemeColors.MainFont,
                FontSize = 12,
                AcceptsReturn = true,
                Height = 62,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                BorderBrush = ThemeColors.InputBorderBrush,
                Padding = new Thickness(8, 5, 8, 5),
                Margin = new Thickness(0, 0, 0, 8)
            };
            footerStack.Children.Add(_recipientsBox);

            Button btnInsertFooter = ThemeColors.CreateAccentOutlineButton(
                "✍️ Chèn Khung Nơi Nhận & Chữ Ký Cuối Văn Bản",
                "Tạo bảng ẩn viền Nơi nhận (12pt đậm nghiêng / 11pt) và Chức vụ - Chữ ký chuẩn NĐ 30/2020");
            btnInsertFooter.Click += OnInsertFooterClick;
            footerStack.Children.Add(btnInsertFooter);

            footerCard.Child = footerStack;
            root.Children.Add(footerCard);

            // 4. Card: Tiện ích Ngày tháng hành chính & Đọc ngày tháng bằng chữ
            Border dateCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 4));
            StackPanel dateStack = new StackPanel();
            dateStack.Children.Add(ThemeColors.CreateSectionHeader("Tiện ích Ngày tháng hành chính & Đọc ngày ra chữ"));

            Border datePreviewBox = new Border
            {
                Background = ThemeColors.WordBlueLight,
                BorderBrush = ThemeColors.WordBlueBorder,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(10),
                Margin = new Thickness(0, 0, 0, 8)
            };
            StackPanel dpStack = new StackPanel();
            _dateAdminPreview = new TextBlock
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 12.5,
                FontStyle = FontStyles.Italic,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeColors.WordBlue,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 4)
            };
            dpStack.Children.Add(_dateAdminPreview);

            _dateWordsPreview = new TextBlock
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 12,
                Foreground = ThemeColors.TextPrimary,
                TextWrapping = TextWrapping.Wrap
            };
            dpStack.Children.Add(_dateWordsPreview);
            datePreviewBox.Child = dpStack;
            dateStack.Children.Add(datePreviewBox);

            Button btnCopyDateWords = ThemeColors.CreateSecondaryButton("📋 Sao Chép Ngày Tháng Bằng Chữ", "Sao chép dòng ngày tháng bằng chữ vào Clipboard");
            btnCopyDateWords.Click += (s, e) =>
            {
                try
                {
                    string text = Decree30Formatter.FormatDateInVietnameseWords(DateTime.Now, _locationBox.Text);
                    Clipboard.SetText(text);
                    _notifyStatus("Đã sao chép: " + text);
                }
                catch { }
            };
            dateStack.Children.Add(btnCopyDateWords);

            dateCard.Child = dateStack;
            root.Children.Add(dateCard);

            Content = root;
        }

        private TextBox CreateTextBox(string initialText)
        {
            return new TextBox
            {
                Text = initialText ?? string.Empty,
                FontFamily = ThemeColors.MainFont,
                FontSize = 12,
                BorderBrush = ThemeColors.InputBorderBrush,
                Padding = new Thickness(8, 5, 8, 5),
                Margin = new Thickness(0, 0, 0, 8)
            };
        }

        private ComboBox CreateComboBox(string[] items)
        {
            ComboBox combo = new ComboBox
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 12,
                Padding = new Thickness(8, 5, 8, 5),
                Margin = new Thickness(0, 0, 0, 8),
                Background = Brushes.White,
                BorderBrush = ThemeColors.InputBorderBrush
            };
            foreach (string item in items)
            {
                combo.Items.Add(item);
            }
            return combo;
        }

        private AdminDocOptions ReadOptionsFromUi()
        {
            if (_marginPresetCombo.SelectedIndex == 1)
            {
                _options.TopMarginMm = 20;
                _options.BottomMarginMm = 20;
                _options.LeftMarginMm = 30;
                _options.RightMarginMm = 20;
            }
            else if (_marginPresetCombo.SelectedIndex == 2)
            {
                _options.TopMarginMm = 25;
                _options.BottomMarginMm = 25;
                _options.LeftMarginMm = 35;
                _options.RightMarginMm = 20;
            }
            else
            {
                _options.TopMarginMm = 20;
                _options.BottomMarginMm = 20;
                _options.LeftMarginMm = 30;
                _options.RightMarginMm = 15;
            }

            _options.BodyFontSize = _fontSizeCombo.SelectedIndex == 1 ? 13.0f : 14.0f;
            _options.EnableDecree30PageNumbering = _chkPageNumber.IsChecked == true;

            _options.ParentOrgName = _parentOrgBox.Text;
            _options.IssuingOrgName = _issuingOrgBox.Text;
            _options.DocumentNumber = _docNumberBox.Text;
            _options.LocationName = _locationBox.Text;
            _options.IssueDate = DateTime.Now;

            string selectedDocType = _docTypeCombo.SelectedItem as string ?? string.Empty;
            _options.DocumentTypeTitle = selectedDocType.StartsWith("(") ? string.Empty : selectedDocType;
            _options.DocumentSubject = _subjectBox.Text;

            _options.AuthorityPrefix = _authorityBox.Text;
            _options.SignerPosition = _positionBox.Text;
            _options.SignerFullName = _signerNameBox.Text;
            _options.RecipientsList = _recipientsBox.Text;

            return _options;
        }

        private void RefreshDatePreview()
        {
            string loc = _locationBox != null ? _locationBox.Text : "Hà Nội";
            _dateAdminPreview.Text = "Chuẩn NĐ 30/2020: " + Decree30Formatter.FormatAdminDateLine(loc, DateTime.Now);
            _dateWordsPreview.Text = "Bằng chữ: " + Decree30Formatter.FormatDateInVietnameseWords(DateTime.Now, loc);
        }

        private void OnApplyFullStandardClick(object sender, RoutedEventArgs e)
        {
            string msg;
            _wordService.ApplyDecree30FullStandard(ReadOptionsFromUi(), out msg);
            _notifyStatus(msg);
        }

        private void OnApplyMarginsOnlyClick(object sender, RoutedEventArgs e)
        {
            string msg;
            _wordService.ApplyDecree30MarginsOnly(ReadOptionsFromUi(), out msg);
            _notifyStatus(msg);
        }

        private void OnCleanGarbageClick(object sender, RoutedEventArgs e)
        {
            string msg;
            _wordService.CleanFullDocumentGarbage(out msg);
            _notifyStatus(msg);
        }

        private void OnInsertHeaderClick(object sender, RoutedEventArgs e)
        {
            string msg;
            _wordService.InsertDecree30Header(ReadOptionsFromUi(), out msg);
            _notifyStatus(msg);
        }

        private void OnInsertFooterClick(object sender, RoutedEventArgs e)
        {
            string msg;
            _wordService.InsertDecree30FooterSignature(ReadOptionsFromUi(), out msg);
            _notifyStatus(msg);
        }
    }
}
