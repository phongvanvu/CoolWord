using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Modules.NumberToText
{
    public class NumberToTextModule : IToolModule
    {
        public string Id { get { return "number-to-text"; } }
        public string Title { get { return "Đổi Số Ra Chữ"; } }
        public string Subtitle { get { return "Đọc số tiền VNĐ, ngoại tệ & số thập phân chuẩn tiếng Việt"; } }
        public string IconText { get { return "123"; } }
        public string Category { get { return "Tài chính & Hành chính"; } }
        public int Order { get { return 1; } }

        public bool ExecuteQuickAction(IWordDocumentService wordService, out string statusMessage)
        {
            string converted;
            NumberConvertOptions options = SettingsService.Current.ToNumberConvertOptions();
            return wordService.ConvertSelectedNumber(options, out converted, out statusMessage);
        }

        public UIElement CreateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            return new NumberToTextModuleView(wordService, notifyStatus);
        }
    }

    public class NumberToTextModuleView : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Action<string> _notifyStatus;
        private bool _isUpdatingUi;

        private TextBox _inputBox;
        private TextBlock _resultTextBlock;
        private TextBlock _insertionPreviewTextBlock;

        private ComboBox _insertModeCombo;
        private ComboBox _readingModeCombo;
        private ComboBox _dialectCombo;
        private ComboBox _capitalizationCombo;

        private Grid _customUnitPanel;
        private TextBox _customMajorBox;
        private TextBox _customMinorBox;

        private CheckBox _chkAppendChan;
        private CheckBox _chkAppendDotSlash;
        private CheckBox _chkItalic;
        private CheckBox _chkLiveSync;

        private StackPanel _historyPanel;

        public NumberToTextModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            _wordService = wordService;
            _notifyStatus = notifyStatus ?? (_ => { });

            BuildUi();
            LoadSettingsIntoUi();
            RefreshPreview();
            RefreshHistoryUi();

            if (_wordService != null)
            {
                _wordService.SelectionChanged += OnWordSelectionChanged;
            }
            SettingsService.SettingsChanged += OnExternalSettingsChanged;
        }

        private void OnExternalSettingsChanged(object sender, EventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                LoadSettingsIntoUi();
                RefreshPreview();
                RefreshHistoryUi();
            }));
        }

        private void OnWordSelectionChanged(object sender, WordSelectionEventArgs e)
        {
            if (!SettingsService.Current.LiveSyncSelectionToTaskPane)
                return;

            if (string.IsNullOrWhiteSpace(e.SelectedText))
                return;

            ParsedNumberInfo parsed = NumberExtractor.Extract(e.SelectedText, SettingsService.Current.ToNumberConvertOptions().ReadingMode);
            if (parsed.IsValid)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    _inputBox.Text = parsed.MatchedText;
                    _notifyStatus("Đã nhận diện số từ vùng chọn Word: " + parsed.MatchedText);
                }));
            }
        }

        private void BuildUi()
        {
            StackPanel root = new StackPanel
            {
                Margin = new Thickness(0, 0, 0, 12)
            };

            // 1. Thẻ thao tác nhanh với văn bản Word
            Border quickActionCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel quickActionStack = new StackPanel();

            quickActionStack.Children.Add(ThemeColors.CreateSectionHeader("Thao tác trực tiếp trên văn bản Word"));

            TextBlock quickHint = new TextBlock
            {
                Text = "Bôi đen con số trong tài liệu Word (hoặc đặt con trỏ vào giữa số) rồi nhấn nút bên dưới:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            };
            quickActionStack.Children.Add(quickHint);

            Button btnConvertSelection = ThemeColors.CreatePrimaryButton(
                "⚡ Đổi Số Đang Chọn Trong Word",
                "Đổi ngay con số đang bôi đen trong Word theo chế độ bên dưới");
            btnConvertSelection.Margin = new Thickness(0, 0, 0, 6);
            btnConvertSelection.Click += OnConvertSelectionClick;
            quickActionStack.Children.Add(btnConvertSelection);

            Grid subActionGrid = new Grid();
            subActionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            subActionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            subActionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Button btnPullFromWord = ThemeColors.CreateSecondaryButton("📥 Lấy số đang chọn", "Đưa số đang bôi đen trong Word vào ô nhập bên dưới");
            btnPullFromWord.Click += OnPullFromWordClick;
            Grid.SetColumn(btnPullFromWord, 0);
            subActionGrid.Children.Add(btnPullFromWord);

            Button btnBatchConvert = ThemeColors.CreateSecondaryButton("📑 Đổi mọi số vùng chọn", "Tìm và đổi toàn bộ các con số nằm trong đoạn văn/bảng đang bôi đen");
            btnBatchConvert.Click += OnBatchConvertClick;
            Grid.SetColumn(btnBatchConvert, 2);
            subActionGrid.Children.Add(btnBatchConvert);

            quickActionStack.Children.Add(subActionGrid);
            quickActionCard.Child = quickActionStack;
            root.Children.Add(quickActionCard);

            // 2. Thẻ Nhập số & Xem trước kết quả trực tiếp
            Border converterCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel converterStack = new StackPanel();

            converterStack.Children.Add(ThemeColors.CreateSectionHeader("Nhập số & Xem trước kết quả"));

            _inputBox = new TextBox
            {
                Text = "15.680.000",
                FontFamily = ThemeColors.MainFont,
                FontSize = 15,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeColors.TextPrimary,
                BorderBrush = ThemeColors.InputBorderBrush,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(10, 7, 10, 7)
            };
            _inputBox.TextChanged += (s, e) => RefreshPreview();
            converterStack.Children.Add(_inputBox);

            // Dãy nút mẫu thử nhanh
            WrapPanel sampleRow = new WrapPanel
            {
                Margin = new Thickness(0, 6, 0, 10)
            };
            sampleRow.Children.Add(new TextBlock
            {
                Text = "Mẫu thử: ",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11,
                Foreground = ThemeColors.TextMuted,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 4, 0)
            });

            string[] samples = new[] { "1.250.000", "15.680.500", "1.005.021", "2,450.50 USD", "3,14" };
            foreach (string sample in samples)
            {
                string captured = sample;
                Button chip = new Button
                {
                    Content = captured,
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 11,
                    Foreground = ThemeColors.WordBlue,
                    Background = ThemeColors.WordBlueLight,
                    BorderBrush = ThemeColors.WordBlueBorder,
                    BorderThickness = new Thickness(1),
                    Padding = new Thickness(6, 2, 6, 2),
                    Margin = new Thickness(0, 2, 4, 2),
                    Cursor = System.Windows.Input.Cursors.Hand
                };
                chip.Click += (s, e) => { _inputBox.Text = captured; };
                sampleRow.Children.Add(chip);
            }
            converterStack.Children.Add(sampleRow);

            // Khung hiển thị kết quả chữ
            Border resultBox = new Border
            {
                Background = ThemeColors.WordBlueLight,
                BorderBrush = ThemeColors.WordBlueBorder,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(10),
                Margin = new Thickness(0, 0, 0, 8)
            };
            StackPanel resultStack = new StackPanel();
            resultStack.Children.Add(new TextBlock
            {
                Text = "KẾT QUẢ BẰNG CHỮ:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 10.5,
                FontWeight = FontWeights.Bold,
                Foreground = ThemeColors.WordBlue,
                Margin = new Thickness(0, 0, 0, 4)
            });

            _resultTextBlock = new TextBlock
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 13.5,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeColors.TextPrimary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 6)
            };
            resultStack.Children.Add(_resultTextBlock);

            _insertionPreviewTextBlock = new TextBlock
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                FontStyle = FontStyles.Italic,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap
            };
            resultStack.Children.Add(_insertionPreviewTextBlock);

            resultBox.Child = resultStack;
            converterStack.Children.Add(resultBox);

            // Nút Chèn vào Word & Sao chép
            Grid resultActionGrid = new Grid();
            resultActionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            resultActionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            resultActionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Button btnInsertToWord = ThemeColors.CreateAccentOutlineButton("📝 Chèn Vào Word", "Chèn kết quả bằng chữ vào vị trí con trỏ hoặc vùng chọn hiện tại");
            btnInsertToWord.Click += OnInsertResultToWordClick;
            Grid.SetColumn(btnInsertToWord, 0);
            resultActionGrid.Children.Add(btnInsertToWord);

            Button btnCopyResult = ThemeColors.CreateSecondaryButton("📋 Sao Chép (Copy)", "Sao chép dòng chữ kết quả vào bộ nhớ tạm (Clipboard)");
            btnCopyResult.Click += OnCopyResultClick;
            Grid.SetColumn(btnCopyResult, 2);
            resultActionGrid.Children.Add(btnCopyResult);

            converterStack.Children.Add(resultActionGrid);
            converterCard.Child = converterStack;
            root.Children.Add(converterCard);

            // 3. Thẻ Cấu hình & Tùy chỉnh cách đọc
            Border optionsCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel optionsStack = new StackPanel();

            optionsStack.Children.Add(ThemeColors.CreateSectionHeader("Cấu hình cách đọc & định dạng chèn"));

            // Chế độ chèn trong Word
            optionsStack.Children.Add(ThemeColors.CreateFieldLabel("Chế độ thay thế / chèn trong Word:"));
            _insertModeCombo = CreateComboBox(new[]
            {
                "Thay thế trực tiếp số đang chọn",
                "Chèn trong ngoặc sau số: 1.000 (Một nghìn...)",
                "Chèn kèm nhãn: 1.000 (Bằng chữ: Một nghìn...)",
                "Chèn xuống dòng mới: Bằng chữ: Một nghìn..."
            });
            _insertModeCombo.SelectionChanged += OnOptionSelectionChanged;
            optionsStack.Children.Add(_insertModeCombo);

            // Chế độ đơn vị đọc
            optionsStack.Children.Add(ThemeColors.CreateFieldLabel("Loại số / Đơn vị tiền tệ:"));
            _readingModeCombo = CreateComboBox(new[]
            {
                "Tiền tệ Việt Nam (đồng)",
                "Ngoại tệ USD (đô la Mỹ / xu)",
                "Số thuần túy (không kèm đơn vị)",
                "Đơn vị tiền tệ / đo lường tùy chỉnh..."
            });
            _readingModeCombo.SelectionChanged += OnOptionSelectionChanged;
            optionsStack.Children.Add(_readingModeCombo);

            // Ô nhập đơn vị tùy chỉnh (ẩn/hiện theo loại đơn vị)
            _customUnitPanel = new Grid { Margin = new Thickness(0, 0, 0, 8), Visibility = Visibility.Collapsed };
            _customUnitPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            _customUnitPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(8) });
            _customUnitPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            StackPanel majorPanel = new StackPanel();
            majorPanel.Children.Add(ThemeColors.CreateFieldLabel("Đơn vị chính (VD: EUR, tấn):"));
            _customMajorBox = new TextBox { Text = "đồng", Padding = new Thickness(6, 4, 6, 4), FontFamily = ThemeColors.MainFont };
            _customMajorBox.TextChanged += OnOptionSelectionChanged;
            majorPanel.Children.Add(_customMajorBox);
            Grid.SetColumn(majorPanel, 0);
            _customUnitPanel.Children.Add(majorPanel);

            StackPanel minorPanel = new StackPanel();
            minorPanel.Children.Add(ThemeColors.CreateFieldLabel("Đơn vị lẻ (để trống nếu đọc phẩy):"));
            _customMinorBox = new TextBox { Text = "xu", Padding = new Thickness(6, 4, 6, 4), FontFamily = ThemeColors.MainFont };
            _customMinorBox.TextChanged += OnOptionSelectionChanged;
            minorPanel.Children.Add(_customMinorBox);
            Grid.SetColumn(minorPanel, 2);
            _customUnitPanel.Children.Add(minorPanel);

            optionsStack.Children.Add(_customUnitPanel);

            // Phương ngữ
            optionsStack.Children.Add(ThemeColors.CreateFieldLabel("Phương ngữ đọc số:"));
            _dialectCombo = CreateComboBox(new[]
            {
                "Chuẩn Bắc Bộ (nghìn, linh, tư)",
                "Chuẩn Nam Bộ (ngàn, lẻ, bốn)",
                "Phổ thông hành chính (nghìn, lẻ, tư)"
            });
            _dialectCombo.SelectionChanged += OnOptionSelectionChanged;
            optionsStack.Children.Add(_dialectCombo);

            // Kiểu viết hoa
            optionsStack.Children.Add(ThemeColors.CreateFieldLabel("Kiểu viết hoa:"));
            _capitalizationCombo = CreateComboBox(new[]
            {
                "Viết hoa chữ cái đầu câu (Một triệu...)",
                "viết thường toàn bộ (một triệu...)",
                "Viết Hoa Chữ Đầu Mỗi Từ (Một Triệu...)",
                "IN HOA TOÀN BỘ (MỘT TRIỆU...)"
            });
            _capitalizationCombo.SelectionChanged += OnOptionSelectionChanged;
            optionsStack.Children.Add(_capitalizationCombo);

            // Các checkbox tùy chọn
            _chkAppendChan = CreateCheckBox("Thêm chữ \"chẵn\" khi số tiền nguyên (không có số lẻ)");
            _chkItalic = CreateCheckBox("In nghiêng phần chữ khi chèn vào văn bản Word");
            _chkAppendDotSlash = CreateCheckBox("Thêm ký hiệu \"./.\" ở cuối câu đọc số tiền");
            _chkLiveSync = CreateCheckBox("Tự động đưa số đang bôi đen trong Word vào ô xem trước");

            optionsStack.Children.Add(_chkAppendChan);
            optionsStack.Children.Add(_chkItalic);
            optionsStack.Children.Add(_chkAppendDotSlash);
            optionsStack.Children.Add(_chkLiveSync);

            optionsCard.Child = optionsStack;
            root.Children.Add(optionsCard);

            // 4. Thẻ Lịch sử chuyển đổi gần đây
            Border historyCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 4));
            StackPanel historyWrapper = new StackPanel();
            historyWrapper.Children.Add(ThemeColors.CreateSectionHeader("Lịch sử chuyển đổi gần đây"));

            _historyPanel = new StackPanel();
            historyWrapper.Children.Add(_historyPanel);

            historyCard.Child = historyWrapper;
            root.Children.Add(historyCard);

            Content = root;
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

        private CheckBox CreateCheckBox(string label)
        {
            CheckBox chk = new CheckBox
            {
                Content = new TextBlock
                {
                    Text = label,
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 12,
                    Foreground = ThemeColors.TextPrimary,
                    TextWrapping = TextWrapping.Wrap
                },
                Margin = new Thickness(0, 4, 0, 4),
                VerticalContentAlignment = VerticalAlignment.Center,
                Cursor = System.Windows.Input.Cursors.Hand
            };
            chk.Checked += OnOptionSelectionChanged;
            chk.Unchecked += OnOptionSelectionChanged;
            return chk;
        }

        private void LoadSettingsIntoUi()
        {
            _isUpdatingUi = true;
            try
            {
                AppSettings s = SettingsService.Current;
                _insertModeCombo.SelectedIndex = ClampIndex(s.InsertMode, 0, 3);
                _readingModeCombo.SelectedIndex = ClampIndex(s.ReadingMode, 0, 3);
                _dialectCombo.SelectedIndex = ClampIndex(s.Dialect, 0, 2);
                _capitalizationCombo.SelectedIndex = ClampIndex(s.Capitalization, 0, 3);

                _customMajorBox.Text = s.CustomMajorUnit ?? "đồng";
                _customMinorBox.Text = s.CustomMinorUnit ?? "xu";
                _customUnitPanel.Visibility = s.ReadingMode == (int)NumberReadingMode.CustomCurrency
                    ? Visibility.Visible
                    : Visibility.Collapsed;

                _chkAppendChan.IsChecked = s.AppendChan;
                _chkAppendDotSlash.IsChecked = s.AppendDotSlash;
                _chkItalic.IsChecked = s.ItalicizeConvertedText;
                _chkLiveSync.IsChecked = s.LiveSyncSelectionToTaskPane;
            }
            finally
            {
                _isUpdatingUi = false;
            }
        }

        private void SaveUiToSettings()
        {
            if (_isUpdatingUi) return;

            AppSettings s = SettingsService.Current;
            s.InsertMode = Math.Max(0, _insertModeCombo.SelectedIndex);
            s.ReadingMode = Math.Max(0, _readingModeCombo.SelectedIndex);
            s.Dialect = Math.Max(0, _dialectCombo.SelectedIndex);
            s.Capitalization = Math.Max(0, _capitalizationCombo.SelectedIndex);
            s.CustomMajorUnit = _customMajorBox.Text;
            s.CustomMinorUnit = _customMinorBox.Text;
            s.AppendChan = _chkAppendChan.IsChecked == true;
            s.AppendDotSlash = _chkAppendDotSlash.IsChecked == true;
            s.ItalicizeConvertedText = _chkItalic.IsChecked == true;
            s.LiveSyncSelectionToTaskPane = _chkLiveSync.IsChecked == true;

            _customUnitPanel.Visibility = s.ReadingMode == (int)NumberReadingMode.CustomCurrency
                ? Visibility.Visible
                : Visibility.Collapsed;

            SettingsService.Save(s);
        }

        private void OnOptionSelectionChanged(object sender, EventArgs e)
        {
            if (_isUpdatingUi) return;
            SaveUiToSettings();
            RefreshPreview();
        }

        private void RefreshPreview()
        {
            string raw = _inputBox.Text;
            NumberConvertOptions options = SettingsService.Current.ToNumberConvertOptions();
            ParsedNumberInfo parsed = NumberExtractor.Extract(raw, options.ReadingMode);

            if (!parsed.IsValid)
            {
                _resultTextBlock.Text = "(Vui lòng nhập một số hợp lệ)";
                _resultTextBlock.Foreground = ThemeColors.TextMuted;
                _insertionPreviewTextBlock.Text = string.Empty;
                return;
            }

            string converted = VietnameseNumberConverter.ConvertParsed(parsed, options);
            _resultTextBlock.Text = converted;
            _resultTextBlock.Foreground = ThemeColors.TextPrimary;

            string previewInDoc = VietnameseNumberConverter.FormatForWordInsertion(parsed.MatchedText, converted, options);
            _insertionPreviewTextBlock.Text = "Minh họa trong Word: " + previewInDoc.Replace(Environment.NewLine, " ↵ ");
        }

        private void RefreshHistoryUi()
        {
            _historyPanel.Children.Clear();
            var history = SettingsService.Current.History;
            if (history == null || history.Count == 0)
            {
                _historyPanel.Children.Add(new TextBlock
                {
                    Text = "Chưa có lịch sử chuyển đổi.",
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 11.5,
                    Foreground = ThemeColors.TextMuted
                });
                return;
            }

            int maxShow = Math.Min(5, history.Count);
            for (int i = 0; i < maxShow; i++)
            {
                ConversionHistoryItem item = history[i];
                Border row = new Border
                {
                    Background = ThemeColors.SubtleBackground,
                    BorderBrush = ThemeColors.BorderBrush,
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(5),
                    Padding = new Thickness(8, 5, 8, 5),
                    Margin = new Thickness(0, 0, 0, 5),
                    Cursor = System.Windows.Input.Cursors.Hand,
                    ToolTip = "Nhấn để nạp lại số này và sao chép kết quả"
                };

                StackPanel sp = new StackPanel();
                sp.Children.Add(new TextBlock
                {
                    Text = item.InputNumber,
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 11.5,
                    FontWeight = FontWeights.Bold,
                    Foreground = ThemeColors.WordBlue
                });
                sp.Children.Add(new TextBlock
                {
                    Text = item.OutputWords,
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 11.5,
                    Foreground = ThemeColors.TextSecondary,
                    TextWrapping = TextWrapping.Wrap
                });

                row.Child = sp;
                string capturedInput = item.InputNumber;
                string capturedOutput = item.OutputWords;
                row.MouseLeftButtonUp += (s, e) =>
                {
                    _inputBox.Text = capturedInput;
                    try
                    {
                        Clipboard.SetText(capturedOutput);
                        _notifyStatus("Đã sao chép từ lịch sử: " + capturedOutput);
                    }
                    catch { }
                };

                _historyPanel.Children.Add(row);
            }
        }

        private void OnConvertSelectionClick(object sender, RoutedEventArgs e)
        {
            string converted;
            string status;
            NumberConvertOptions options = SettingsService.Current.ToNumberConvertOptions();
            if (_wordService.ConvertSelectedNumber(options, out converted, out status))
            {
                string sel = _wordService.GetSelectedText();
                if (!string.IsNullOrEmpty(sel))
                {
                    _inputBox.Text = sel;
                }
                RefreshHistoryUi();
            }
            _notifyStatus(status);
        }

        private void OnPullFromWordClick(object sender, RoutedEventArgs e)
        {
            string selected = _wordService.GetSelectedText();
            if (string.IsNullOrWhiteSpace(selected))
            {
                _notifyStatus("Không tìm thấy văn bản đang bôi đen trong Word.");
                return;
            }
            _inputBox.Text = selected;
            _notifyStatus("Đã lấy nội dung từ Word: " + selected);
        }

        private void OnBatchConvertClick(object sender, RoutedEventArgs e)
        {
            string status;
            NumberConvertOptions options = SettingsService.Current.ToNumberConvertOptions();
            _wordService.ConvertAllNumbersInSelection(options, out status);
            RefreshHistoryUi();
            _notifyStatus(status);
        }

        private void OnInsertResultToWordClick(object sender, RoutedEventArgs e)
        {
            NumberConvertOptions options = SettingsService.Current.ToNumberConvertOptions();
            ParsedNumberInfo parsed = NumberExtractor.Extract(_inputBox.Text, options.ReadingMode);
            if (!parsed.IsValid)
            {
                _notifyStatus("Vui lòng nhập số hợp lệ trước khi chèn vào Word.");
                return;
            }

            string converted = VietnameseNumberConverter.ConvertParsed(parsed, options);
            string status;
            if (_wordService.InsertConvertedText(parsed.MatchedText, converted, options, out status))
            {
                SettingsService.Current.AddHistory(parsed.MatchedText, converted);
                SettingsService.Save();
                RefreshHistoryUi();
            }
            _notifyStatus(status);
        }

        private void OnCopyResultClick(object sender, RoutedEventArgs e)
        {
            NumberConvertOptions options = SettingsService.Current.ToNumberConvertOptions();
            ParsedNumberInfo parsed = NumberExtractor.Extract(_inputBox.Text, options.ReadingMode);
            if (!parsed.IsValid)
            {
                _notifyStatus("Chưa có kết quả hợp lệ để sao chép.");
                return;
            }

            string converted = VietnameseNumberConverter.ConvertParsed(parsed, options);
            try
            {
                Clipboard.SetText(converted);
                SettingsService.Current.AddHistory(parsed.MatchedText, converted);
                SettingsService.Save();
                RefreshHistoryUi();
                _notifyStatus("Đã sao chép vào Clipboard: " + converted);
            }
            catch (Exception ex)
            {
                _notifyStatus("Không thể truy cập Clipboard: " + ex.Message);
            }
        }

        private static int ClampIndex(int val, int min, int max)
        {
            if (val < min) return min;
            if (val > max) return max;
            return val;
        }
    }
}
