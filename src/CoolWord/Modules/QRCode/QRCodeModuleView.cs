using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Modules.QRCode
{
    public class QRCodeModuleView : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Action<string> _notifyStatus;

        // Mode: 0 = URL/Text, 1 = VietQR, 2 = vCard, 3 = AdminDoc
        private int _currentMode = 0;

        // Preview image
        private System.Windows.Controls.Image _imgPreview;
        private Bitmap _currentBitmap;

        // Controls Mode 0: URL / Text
        private TextBox _txtContent;

        // Controls Mode 1: VietQR
        private ComboBox _cboBank;
        private TextBox _txtAccountNo;
        private TextBox _txtAccountName;
        private TextBox _txtAmount;
        private TextBox _txtDescription;

        // Controls Mode 2: vCard
        private TextBox _txtVCardName;
        private TextBox _txtVCardPhone;
        private TextBox _txtVCardEmail;
        private TextBox _txtVCardTitle;
        private TextBox _txtVCardCompany;
        private TextBox _txtVCardWebsite;

        // Controls Mode 3: Admin Doc
        private TextBox _txtDocNumber;
        private TextBox _txtDocDate;
        private TextBox _txtDocAuthority;
        private TextBox _txtDocTitle;
        private TextBox _txtDocUrl;

        // Customization
        private ComboBox _cboSize;
        private CheckBox _chkCenter;
        private CheckBox _chkCaption;
        private TextBox _txtCaption;

        // Panels
        private StackPanel _panelUrlText;
        private StackPanel _panelVietQr;
        private StackPanel _panelVCard;
        private StackPanel _panelAdminDoc;

        public QRCodeModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            _wordService = wordService;
            _notifyStatus = notifyStatus ?? (_ => { });

            BuildUi();
            UpdatePreview();
        }

        private void BuildUi()
        {
            StackPanel root = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };

            // 1. Card: Loại Mã QR Code
            Border typeCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel typeStack = new StackPanel();
            typeStack.Children.Add(ThemeColors.CreateSectionHeader("Tạo & Chèn Mã QR Code"));
            typeStack.Children.Add(new TextBlock
            {
                Text = "Tạo mã QR Code chất lượng cao chuẩn VietQR ngân hàng, URL, liên hệ vCard và chèn trực tiếp vào tài liệu Word:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            // Grid 2x2 cho 4 loại QR Code
            Grid typeGrid = new Grid { Margin = new Thickness(0, 0, 0, 4) };
            typeGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            typeGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            typeGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            typeGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            typeGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(6) });
            typeGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            Button btnTypeUrl = ThemeColors.CreatePrimaryButton("🔗 Liên kết / Chữ", "URL website hoặc văn bản bất kỳ");
            btnTypeUrl.Height = 28;
            btnTypeUrl.Click += (s, e) => SwitchMode(0);
            Grid.SetRow(btnTypeUrl, 0);
            Grid.SetColumn(btnTypeUrl, 0);
            typeGrid.Children.Add(btnTypeUrl);

            Button btnTypeVietQr = ThemeColors.CreateSecondaryButton("🏦 VietQR Ngân hàng", "Chuẩn thanh toán NAPAS 247");
            btnTypeVietQr.Height = 28;
            btnTypeVietQr.Click += (s, e) => SwitchMode(1);
            Grid.SetRow(btnTypeVietQr, 0);
            Grid.SetColumn(btnTypeVietQr, 2);
            typeGrid.Children.Add(btnTypeVietQr);

            Button btnTypeVCard = ThemeColors.CreateSecondaryButton("📇 Danh thiếp vCard", "Thông tin liên hệ, SĐT, Email");
            btnTypeVCard.Height = 28;
            btnTypeVCard.Click += (s, e) => SwitchMode(2);
            Grid.SetRow(btnTypeVCard, 2);
            Grid.SetColumn(btnTypeVCard, 0);
            typeGrid.Children.Add(btnTypeVCard);

            Button btnTypeAdmin = ThemeColors.CreateSecondaryButton("🏛️ Tra cứu NĐ 30", "Số hiệu văn bản, cơ quan, link");
            btnTypeAdmin.Height = 28;
            btnTypeAdmin.Click += (s, e) => SwitchMode(3);
            Grid.SetRow(btnTypeAdmin, 2);
            Grid.SetColumn(btnTypeAdmin, 2);
            typeGrid.Children.Add(btnTypeAdmin);

            typeStack.Children.Add(typeGrid);
            typeCard.Child = typeStack;
            root.Children.Add(typeCard);

            // 2. Card: Nhập Liệu Thông Tin
            Border inputCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel inputStack = new StackPanel();

            // Panel 0: URL / Text
            _panelUrlText = new StackPanel();
            _panelUrlText.Children.Add(ThemeColors.CreateFieldLabel("Đường dẫn (URL) hoặc Đoạn văn bản:"));
            _txtContent = CreateInputTextBox("https://dantri.com.vn", 60, true);
            _txtContent.TextChanged += (s, e) => UpdatePreview();
            _panelUrlText.Children.Add(_txtContent);

            Button btnGetSelection = ThemeColors.CreateSecondaryButton("📋 Lấy từ đoạn bôi đen trong Word", "Đọc nội dung vùng đang chọn");
            btnGetSelection.Margin = new Thickness(0, 4, 0, 0);
            btnGetSelection.Click += (s, e) =>
            {
                if (_wordService != null)
                {
                    string msg;
                    bool ok = _wordService.InsertQrCodeFromSelection(out msg);
                    _notifyStatus(msg);
                }
            };
            _panelUrlText.Children.Add(btnGetSelection);
            inputStack.Children.Add(_panelUrlText);

            // Panel 1: VietQR
            _panelVietQr = new StackPanel { Visibility = Visibility.Collapsed };
            _panelVietQr.Children.Add(ThemeColors.CreateFieldLabel("Ngân hàng nhận:"));
            _cboBank = new ComboBox
            {
                Height = 28,
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Margin = new Thickness(0, 0, 0, 6)
            };
            foreach (var b in QRCodeService.PopularBanks)
            {
                _cboBank.Items.Add(b);
            }
            _cboBank.SelectedIndex = 0;
            _cboBank.SelectionChanged += (s, e) => UpdatePreview();
            _panelVietQr.Children.Add(_cboBank);

            _panelVietQr.Children.Add(ThemeColors.CreateFieldLabel("Số tài khoản:"));
            _txtAccountNo = CreateInputTextBox("0123456789", 28, false);
            _txtAccountNo.TextChanged += (s, e) => UpdatePreview();
            _panelVietQr.Children.Add(_txtAccountNo);

            _panelVietQr.Children.Add(ThemeColors.CreateFieldLabel("Tên chủ tài khoản:"));
            _txtAccountName = CreateInputTextBox("CONG TY TNHH COOLWORD", 28, false);
            _txtAccountName.TextChanged += (s, e) => UpdatePreview();
            _panelVietQr.Children.Add(_txtAccountName);

            _panelVietQr.Children.Add(ThemeColors.CreateFieldLabel("Số tiền chuyển (VNĐ - tùy chọn):"));
            _txtAmount = CreateInputTextBox("500000", 28, false);
            _txtAmount.TextChanged += (s, e) => UpdatePreview();
            _panelVietQr.Children.Add(_txtAmount);

            _panelVietQr.Children.Add(ThemeColors.CreateFieldLabel("Nội dung chuyển tiền (tùy chọn):"));
            _txtDescription = CreateInputTextBox("Thanh toan hop dong", 28, false);
            _txtDescription.TextChanged += (s, e) => UpdatePreview();
            _panelVietQr.Children.Add(_txtDescription);
            inputStack.Children.Add(_panelVietQr);

            // Panel 2: vCard
            _panelVCard = new StackPanel { Visibility = Visibility.Collapsed };
            _panelVCard.Children.Add(ThemeColors.CreateFieldLabel("Họ và tên:"));
            _txtVCardName = CreateInputTextBox("Nguyễn Văn An", 28, false);
            _txtVCardName.TextChanged += (s, e) => UpdatePreview();
            _panelVCard.Children.Add(_txtVCardName);

            _panelVCard.Children.Add(ThemeColors.CreateFieldLabel("Số điện thoại:"));
            _txtVCardPhone = CreateInputTextBox("0987654321", 28, false);
            _txtVCardPhone.TextChanged += (s, e) => UpdatePreview();
            _panelVCard.Children.Add(_txtVCardPhone);

            _panelVCard.Children.Add(ThemeColors.CreateFieldLabel("Email:"));
            _txtVCardEmail = CreateInputTextBox("an.nguyen@coolword.vn", 28, false);
            _txtVCardEmail.TextChanged += (s, e) => UpdatePreview();
            _panelVCard.Children.Add(_txtVCardEmail);

            _panelVCard.Children.Add(ThemeColors.CreateFieldLabel("Chức vụ & Đơn vị:"));
            _txtVCardTitle = CreateInputTextBox("Giám đốc kinh doanh", 28, false);
            _txtVCardTitle.TextChanged += (s, e) => UpdatePreview();
            _panelVCard.Children.Add(_txtVCardTitle);

            _panelVCard.Children.Add(ThemeColors.CreateFieldLabel("Công ty:"));
            _txtVCardCompany = CreateInputTextBox("CoolWord Technology", 28, false);
            _txtVCardCompany.TextChanged += (s, e) => UpdatePreview();
            _panelVCard.Children.Add(_txtVCardCompany);

            _panelVCard.Children.Add(ThemeColors.CreateFieldLabel("Website (tùy chọn):"));
            _txtVCardWebsite = CreateInputTextBox("https://coolword.vn", 28, false);
            _txtVCardWebsite.TextChanged += (s, e) => UpdatePreview();
            _panelVCard.Children.Add(_txtVCardWebsite);
            inputStack.Children.Add(_panelVCard);

            // Panel 3: Admin Doc
            _panelAdminDoc = new StackPanel { Visibility = Visibility.Collapsed };
            _panelAdminDoc.Children.Add(ThemeColors.CreateFieldLabel("Số & Ký hiệu văn bản:"));
            _txtDocNumber = CreateInputTextBox("12/QĐ-UBND", 28, false);
            _txtDocNumber.TextChanged += (s, e) => UpdatePreview();
            _panelAdminDoc.Children.Add(_txtDocNumber);

            _panelAdminDoc.Children.Add(ThemeColors.CreateFieldLabel("Ngày ban hành:"));
            _txtDocDate = CreateInputTextBox(DateTime.Now.ToString("dd/MM/yyyy"), 28, false);
            _txtDocDate.TextChanged += (s, e) => UpdatePreview();
            _panelAdminDoc.Children.Add(_txtDocDate);

            _panelAdminDoc.Children.Add(ThemeColors.CreateFieldLabel("Cơ quan ban hành:"));
            _txtDocAuthority = CreateInputTextBox("Ủy ban nhân dân tỉnh", 28, false);
            _txtDocAuthority.TextChanged += (s, e) => UpdatePreview();
            _panelAdminDoc.Children.Add(_txtDocAuthority);

            _panelAdminDoc.Children.Add(ThemeColors.CreateFieldLabel("Trích yếu nội dung:"));
            _txtDocTitle = CreateInputTextBox("Về việc phê duyệt kế hoạch chuyển đổi số", 28, false);
            _txtDocTitle.TextChanged += (s, e) => UpdatePreview();
            _panelAdminDoc.Children.Add(_txtDocTitle);

            _panelAdminDoc.Children.Add(ThemeColors.CreateFieldLabel("Đường dẫn tra cứu trực tuyến (URL):"));
            _txtDocUrl = CreateInputTextBox("https://vanban.chinhphu.vn", 28, false);
            _txtDocUrl.TextChanged += (s, e) => UpdatePreview();
            _panelAdminDoc.Children.Add(_txtDocUrl);
            inputStack.Children.Add(_panelAdminDoc);

            inputCard.Child = inputStack;
            root.Children.Add(inputCard);

            // 3. Card: Xem Trước & Tùy Chọn Trình Bày
            Border previewCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel prevStack = new StackPanel();
            prevStack.Children.Add(ThemeColors.CreateSectionHeader("Xem Trước Mã QR Code"));

            // Khung ảnh QR Code
            Border imgBorder = new Border
            {
                Background = ThemeColors.SubtleBackground,
                BorderBrush = ThemeColors.BorderBrush,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(10),
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 4, 0, 10)
            };
            _imgPreview = new System.Windows.Controls.Image
            {
                Width = 140,
                Height = 140,
                Stretch = System.Windows.Media.Stretch.Uniform
            };
            imgBorder.Child = _imgPreview;
            prevStack.Children.Add(imgBorder);

            // Tùy chọn Kích thước
            prevStack.Children.Add(ThemeColors.CreateFieldLabel("Kích thước chèn vào Word:"));
            _cboSize = new ComboBox
            {
                Height = 28,
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Margin = new Thickness(0, 0, 0, 6)
            };
            _cboSize.Items.Add("Nhỏ gọn (100 pt ~ 3.5 cm)");
            _cboSize.Items.Add("Vừa phải (130 pt ~ 4.5 cm)");
            _cboSize.Items.Add("Lớn rõ nét (170 pt ~ 6 cm)");
            _cboSize.Items.Add("Rất lớn (220 pt ~ 7.8 cm)");
            _cboSize.SelectedIndex = 1;
            prevStack.Children.Add(_cboSize);

            _chkCenter = new CheckBox
            {
                IsChecked = true,
                Margin = new Thickness(0, 0, 0, 4),
                Content = new TextBlock
                {
                    Text = "Căn giữa trang (Center align)",
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 11.5,
                    Foreground = ThemeColors.TextPrimary
                }
            };
            prevStack.Children.Add(_chkCenter);

            _chkCaption = new CheckBox
            {
                IsChecked = true,
                Margin = new Thickness(0, 0, 0, 4),
                Content = new TextBlock
                {
                    Text = "Kèm dòng chú thích bên dưới ảnh:",
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 11.5,
                    Foreground = ThemeColors.TextPrimary
                }
            };
            prevStack.Children.Add(_chkCaption);

            _txtCaption = CreateInputTextBox("Quét mã để truy cập liên kết", 26, false);
            prevStack.Children.Add(_txtCaption);

            previewCard.Child = prevStack;
            root.Children.Add(previewCard);

            // 4. Các nút hành động
            Button btnInsert = ThemeColors.CreatePrimaryButton("📥 Chèn QR Code Vào Word", "Chèn ảnh mã QR vào vị trí con trỏ trong tài liệu Word");
            btnInsert.Margin = new Thickness(0, 0, 0, 6);
            btnInsert.Click += (s, e) => ExecuteInsert();
            root.Children.Add(btnInsert);

            Grid actGrid = new Grid();
            actGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            actGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            actGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Button btnSave = ThemeColors.CreateSecondaryButton("💾 Lưu Ảnh (.png)...", "Lưu mã QR Code thành file ảnh PNG");
            btnSave.Click += (s, e) => ExecuteSave();
            Grid.SetColumn(btnSave, 0);
            actGrid.Children.Add(btnSave);

            Button btnCopy = ThemeColors.CreateSecondaryButton("📋 Sao Chép Ảnh", "Sao chép ảnh QR Code vào Clipboard");
            btnCopy.Click += (s, e) => ExecuteCopy();
            Grid.SetColumn(btnCopy, 2);
            actGrid.Children.Add(btnCopy);

            root.Children.Add(actGrid);

            Content = root;
        }

        private void SwitchMode(int mode)
        {
            _currentMode = mode;
            _panelUrlText.Visibility = mode == 0 ? Visibility.Visible : Visibility.Collapsed;
            _panelVietQr.Visibility = mode == 1 ? Visibility.Visible : Visibility.Collapsed;
            _panelVCard.Visibility = mode == 2 ? Visibility.Visible : Visibility.Collapsed;
            _panelAdminDoc.Visibility = mode == 3 ? Visibility.Visible : Visibility.Collapsed;

            // Cập nhật gợi ý chú thích phù hợp
            switch (mode)
            {
                case 0:
                    _txtCaption.Text = "Quét mã để truy cập liên kết";
                    break;
                case 1:
                    _txtCaption.Text = "Quét mã VietQR để thanh toán chuyển khoản";
                    break;
                case 2:
                    _txtCaption.Text = "Quét mã để lưu danh bạ liên hệ";
                    break;
                case 3:
                    _txtCaption.Text = "Quét mã để tra cứu văn bản chính thức";
                    break;
            }

            UpdatePreview();
        }

        private string GetCurrentPayload()
        {
            switch (_currentMode)
            {
                case 0: // URL / Text
                    return string.IsNullOrWhiteSpace(_txtContent.Text) ? "https://coolword.vn" : _txtContent.Text.Trim();

                case 1: // VietQR
                    VietQrBankInfo bank = _cboBank.SelectedItem as VietQrBankInfo;
                    string bin = bank != null ? bank.Bin : "970436";
                    string acc = string.IsNullOrWhiteSpace(_txtAccountNo.Text) ? "0123456789" : _txtAccountNo.Text.Trim();
                    string name = _txtAccountName.Text;
                    decimal amt = 0;
                    decimal.TryParse(_txtAmount.Text.Replace(".", "").Replace(",", ""), out amt);
                    string desc = _txtDescription.Text;
                    return QRCodeService.GenerateVietQrPayload(bin, acc, name, amt, desc);

                case 2: // vCard
                    return QRCodeService.GenerateVCardPayload(
                        _txtVCardName.Text,
                        _txtVCardPhone.Text,
                        _txtVCardEmail.Text,
                        _txtVCardTitle.Text,
                        _txtVCardCompany.Text,
                        _txtVCardWebsite != null ? _txtVCardWebsite.Text : "");

                case 3: // Admin Doc
                    return QRCodeService.GenerateAdminDocPayload(
                        _txtDocNumber.Text,
                        _txtDocDate.Text,
                        _txtDocAuthority.Text,
                        _txtDocTitle.Text,
                        _txtDocUrl.Text);

                default:
                    return "CoolWord";
            }
        }

        private void UpdatePreview()
        {
            try
            {
                string payload = GetCurrentPayload();
                if (string.IsNullOrEmpty(payload)) return;

                if (_currentBitmap != null)
                {
                    _currentBitmap.Dispose();
                    _currentBitmap = null;
                }

                _currentBitmap = QRCodeService.GenerateQrBitmap(payload, 8, Color.FromArgb(15, 23, 42), Color.White);
                if (_currentBitmap != null)
                {
                    _imgPreview.Source = BitmapToBitmapImage(_currentBitmap);
                }
            }
            catch { }
        }

        private void ExecuteInsert()
        {
            if (_wordService == null)
            {
                _notifyStatus("Dịch vụ Word chưa sẵn sàng.");
                return;
            }

            string payload = GetCurrentPayload();
            if (string.IsNullOrEmpty(payload))
            {
                _notifyStatus("Vui lòng điền thông tin để tạo mã QR Code.");
                return;
            }

            float widthPt = 130f;
            switch (_cboSize.SelectedIndex)
            {
                case 0: widthPt = 100f; break;
                case 1: widthPt = 130f; break;
                case 2: widthPt = 170f; break;
                case 3: widthPt = 220f; break;
            }

            bool center = _chkCenter.IsChecked == true;
            string cap = (_chkCaption.IsChecked == true) ? _txtCaption.Text : null;

            string msg;
            bool ok = _wordService.InsertQrCode(payload, widthPt, center, cap, out msg);
            _notifyStatus(msg);
        }

        private void ExecuteSave()
        {
            string payload = GetCurrentPayload();
            if (string.IsNullOrEmpty(payload))
            {
                _notifyStatus("Vui lòng điền thông tin để tạo mã QR Code.");
                return;
            }

            using (var sfd = new System.Windows.Forms.SaveFileDialog())
            {
                sfd.Title = "Lưu mã QR Code thành tệp hình ảnh";
                sfd.Filter = "Hình ảnh PNG (*.png)|*.png|Tất cả tệp (*.*)|*.*";
                sfd.FileName = "QRCode_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".png";
                if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string msg;
                    bool ok = QRCodeService.SaveQrCodeAsImage(payload, sfd.FileName, 12, null, out msg);
                    _notifyStatus(msg);
                }
            }
        }

        private void ExecuteCopy()
        {
            try
            {
                if (_currentBitmap != null)
                {
                    System.Windows.Forms.Clipboard.SetImage(_currentBitmap);
                    _notifyStatus("Đã sao chép hình ảnh mã QR Code vào Clipboard.");
                }
            }
            catch (Exception ex)
            {
                _notifyStatus("Lỗi khi sao chép ảnh: " + ex.Message);
            }
        }

        private static BitmapImage BitmapToBitmapImage(Bitmap bitmap)
        {
            if (bitmap == null) return null;
            using (MemoryStream ms = new MemoryStream())
            {
                bitmap.Save(ms, ImageFormat.Png);
                ms.Position = 0;
                BitmapImage bi = new BitmapImage();
                bi.BeginInit();
                bi.CacheOption = BitmapCacheOption.OnLoad;
                bi.StreamSource = ms;
                bi.EndInit();
                bi.Freeze();
                return bi;
            }
        }

        private TextBox CreateInputTextBox(string defaultText, double height, bool multiline)
        {
            return new TextBox
            {
                Text = defaultText,
                Height = height,
                AcceptsReturn = multiline,
                TextWrapping = multiline ? TextWrapping.Wrap : TextWrapping.NoWrap,
                VerticalScrollBarVisibility = multiline ? ScrollBarVisibility.Auto : ScrollBarVisibility.Disabled,
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextPrimary,
                Background = ThemeColors.SubtleBackground,
                BorderBrush = ThemeColors.InputBorderBrush,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(6, 4, 6, 4),
                Margin = new Thickness(0, 0, 0, 6)
            };
        }
    }
}
