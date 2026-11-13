using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Modules.QuickInput
{
    public class QuickInputModuleView : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Action<string> _notifyStatus;

        private TextBox _newTitleBox;
        private ComboBox _newCategoryCombo;
        private TextBox _newShortcutBox;

        private TextBox _searchBox;
        private ComboBox _categoryFilterCombo;
        private StackPanel _snippetListPanel;

        public QuickInputModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            _wordService = wordService;
            _notifyStatus = notifyStatus ?? (_ => { });

            BuildUi();
            RefreshSnippetList();
        }

        private void BuildUi()
        {
            StackPanel root = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };

            // 1. Card: Lưu Đoạn Văn Bản Đang Chọn Làm Mẫu Mới
            Border saveCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel saveStack = new StackPanel();

            saveStack.Children.Add(ThemeColors.CreateSectionHeader("⭐ Lưu Đoạn Đang Chọn Thành Mẫu Mới"));
            saveStack.Children.Add(new TextBlock
            {
                Text = "Bôi đen nội dung thường dùng trong tài liệu Word (câu chữ, điều khoản, mẫu biểu) và nhấn lưu để tái sử dụng nhanh chóng:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            saveStack.Children.Add(ThemeColors.CreateFieldLabel("Tên gợi nhớ cho mẫu:"));
            _newTitleBox = new TextBox
            {
                Text = string.Empty,
                FontFamily = ThemeColors.MainFont,
                FontSize = 12,
                Padding = new Thickness(8, 5, 8, 5),
                BorderBrush = ThemeColors.InputBorderBrush,
                Margin = new Thickness(0, 0, 0, 6)
            };
            saveStack.Children.Add(_newTitleBox);

            Grid metaGrid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            metaGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.1, GridUnitType.Star) });
            metaGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            metaGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0.9, GridUnitType.Star) });

            StackPanel catPanel = new StackPanel();
            catPanel.Children.Add(ThemeColors.CreateFieldLabel("Phân nhóm:"));
            _newCategoryCombo = new ComboBox
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 12,
                Padding = new Thickness(6, 4, 6, 4),
                Background = Brushes.White,
                BorderBrush = ThemeColors.InputBorderBrush
            };
            _newCategoryCombo.Items.Add("Cá nhân");
            _newCategoryCombo.Items.Add("Hành chính");
            _newCategoryCombo.Items.Add("Kế toán");
            _newCategoryCombo.Items.Add("Hợp đồng");
            _newCategoryCombo.SelectedIndex = 0;
            catPanel.Children.Add(_newCategoryCombo);
            Grid.SetColumn(catPanel, 0);
            metaGrid.Children.Add(catPanel);

            StackPanel scPanel = new StackPanel();
            scPanel.Children.Add(ThemeColors.CreateFieldLabel("Ký hiệu viết tắt:"));
            _newShortcutBox = new TextBox
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 12,
                Padding = new Thickness(6, 4, 6, 4),
                BorderBrush = ThemeColors.InputBorderBrush
            };
            scPanel.Children.Add(_newShortcutBox);
            Grid.SetColumn(scPanel, 2);
            metaGrid.Children.Add(scPanel);

            saveStack.Children.Add(metaGrid);

            Button btnSaveSelection = ThemeColors.CreatePrimaryButton(
                "⭐ Lưu Đoạn Chọn Vào Thư Viện Mẫu",
                "Lấy toàn bộ nội dung đang bôi đen trong Word lưu thành mẫu");
            btnSaveSelection.Click += (s, e) =>
            {
                string msg;
                string cat = _newCategoryCombo.SelectedItem != null ? _newCategoryCombo.SelectedItem.ToString() : "Cá nhân";
                bool ok = _wordService.SaveCurrentSelectionAsSnippet(_newTitleBox.Text, cat, _newShortcutBox.Text, out msg);
                _notifyStatus(msg);
                if (ok)
                {
                    _newTitleBox.Text = string.Empty;
                    _newShortcutBox.Text = string.Empty;
                    RefreshSnippetList();
                }
            };
            saveStack.Children.Add(btnSaveSelection);

            saveCard.Child = saveStack;
            root.Children.Add(saveCard);

            // 2. Card: Thư Viện Mẫu & Tìm Kiếm Nhập Liệu
            Border listCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel listStack = new StackPanel();

            listStack.Children.Add(ThemeColors.CreateSectionHeader("Thư Viện Mẫu Nhập Liệu Nhanh"));

            // Hàng tìm kiếm & lọc
            Grid filterGrid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            filterGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1.3, GridUnitType.Star) });
            filterGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(6) });
            filterGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0.9, GridUnitType.Star) });

            _searchBox = new TextBox
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Padding = new Thickness(8, 5, 8, 5),
                BorderBrush = ThemeColors.InputBorderBrush
            };
            _searchBox.TextChanged += (s, e) => RefreshSnippetList();
            Grid.SetColumn(_searchBox, 0);
            filterGrid.Children.Add(_searchBox);

            _categoryFilterCombo = new ComboBox
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Padding = new Thickness(6, 4, 6, 4),
                Background = Brushes.White,
                BorderBrush = ThemeColors.InputBorderBrush
            };
            _categoryFilterCombo.Items.Add("Tất cả nhóm");
            _categoryFilterCombo.Items.Add("Hành chính");
            _categoryFilterCombo.Items.Add("Kế toán");
            _categoryFilterCombo.Items.Add("Hợp đồng");
            _categoryFilterCombo.Items.Add("Ký hiệu");
            _categoryFilterCombo.Items.Add("Cá nhân");
            _categoryFilterCombo.SelectedIndex = 0;
            _categoryFilterCombo.SelectionChanged += (s, e) => RefreshSnippetList();
            Grid.SetColumn(_categoryFilterCombo, 2);
            filterGrid.Children.Add(_categoryFilterCombo);

            listStack.Children.Add(filterGrid);

            // Danh sách các mẫu (ScrollViewer)
            ScrollViewer scroll = new ScrollViewer
            {
                MaxHeight = 320,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Margin = new Thickness(0, 0, 0, 4)
            };

            _snippetListPanel = new StackPanel();
            scroll.Content = _snippetListPanel;
            listStack.Children.Add(scroll);

            listCard.Child = listStack;
            root.Children.Add(listCard);

            // 3. Card: Chèn Nhanh Ký Hiệu & Đánh Dấu
            Border symCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 4));
            StackPanel symStack = new StackPanel();

            symStack.Children.Add(ThemeColors.CreateSectionHeader("Ký Hiệu Đánh Dấu & Thời Gian Động"));

            WrapPanel symPills = new WrapPanel { Margin = new Thickness(0, 0, 0, 4) };
            AddQuickPill(symPills, "☐ Ô trống", "☐ ");
            AddQuickPill(symPills, "☑ Ô tích", "☑ ");
            AddQuickPill(symPills, "➔ Mũi tên", "➔ ");
            AddQuickPill(symPills, "⭐ Ghi chú", "⭐ Ghi chú: ");
            AddQuickPill(symPills, "📅 Ngày hôm nay", "{DATE}");
            AddQuickPill(symPills, "🏛️ Ngày hành chính", "{DOC_DATE}");
            symStack.Children.Add(symPills);

            symCard.Child = symStack;
            root.Children.Add(symCard);

            Content = root;
        }

        private void AddQuickPill(Panel container, string label, string contentToInsert)
        {
            Button btn = new Button
            {
                Content = label,
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
            btn.Click += (s, e) =>
            {
                string msg;
                _wordService.InsertTextSnippet(contentToInsert, out msg);
                _notifyStatus(msg);
            };
            container.Children.Add(btn);
        }

        public void RefreshSnippetList()
        {
            if (_snippetListPanel == null) return;
            _snippetListPanel.Children.Clear();

            string keyword = _searchBox != null ? (_searchBox.Text ?? string.Empty).Trim().ToLowerInvariant() : string.Empty;
            string selectedCat = _categoryFilterCombo != null && _categoryFilterCombo.SelectedItem != null
                ? _categoryFilterCombo.SelectedItem.ToString()
                : "Tất cả nhóm";

            List<QuickSnippet> all = QuickInputService.GetAllSnippets();

            var filtered = all.Where(s =>
            {
                if (selectedCat != "Tất cả nhóm" && !string.Equals(s.Category, selectedCat, StringComparison.OrdinalIgnoreCase))
                    return false;

                if (!string.IsNullOrEmpty(keyword))
                {
                    bool matchTitle = (s.Title ?? string.Empty).ToLowerInvariant().Contains(keyword);
                    bool matchContent = (s.Content ?? string.Empty).ToLowerInvariant().Contains(keyword);
                    bool matchShortcut = (s.Shortcut ?? string.Empty).ToLowerInvariant().Contains(keyword);
                    bool matchCat = (s.Category ?? string.Empty).ToLowerInvariant().Contains(keyword);
                    return matchTitle || matchContent || matchShortcut || matchCat;
                }
                return true;
            }).ToList();

            if (filtered.Count == 0)
            {
                _snippetListPanel.Children.Add(new TextBlock
                {
                    Text = "Không tìm thấy mẫu phù hợp.",
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 11.5,
                    FontStyle = FontStyles.Italic,
                    Foreground = ThemeColors.TextSecondary,
                    Margin = new Thickness(0, 8, 0, 8)
                });
                return;
            }

            foreach (QuickSnippet item in filtered)
            {
                QuickSnippet captured = item;

                Border itemCard = new Border
                {
                    Background = Brushes.White,
                    BorderBrush = ThemeColors.BorderBrush,
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(6),
                    Padding = new Thickness(8),
                    Margin = new Thickness(0, 0, 0, 6)
                };

                StackPanel itemStack = new StackPanel();

                // Dòng tiêu đề + badge danh mục
                Grid headerGrid = new Grid();
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                TextBlock titleTb = new TextBlock
                {
                    Text = captured.Title,
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 12,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = ThemeColors.TextPrimary,
                    TextTrimming = TextTrimming.CharacterEllipsis
                };
                Grid.SetColumn(titleTb, 0);
                headerGrid.Children.Add(titleTb);

                Border catBadge = new Border
                {
                    Background = ThemeColors.WordBlueLight,
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(5, 1, 5, 1),
                    Child = new TextBlock
                    {
                        Text = captured.Category,
                        FontFamily = ThemeColors.MainFont,
                        FontSize = 10,
                        FontWeight = FontWeights.SemiBold,
                        Foreground = ThemeColors.WordBlue
                    }
                };
                Grid.SetColumn(catBadge, 1);
                headerGrid.Children.Add(catBadge);
                itemStack.Children.Add(headerGrid);

                // Xem trước nội dung (cắt ngắn)
                string rawPreview = QuickInputService.ResolvePlaceholders(captured.Content)
                    .Replace("\r\n", " ")
                    .Replace("\r", " ")
                    .Replace("\n", " ");
                if (rawPreview.Length > 85) rawPreview = rawPreview.Substring(0, 85) + "...";

                TextBlock contentTb = new TextBlock
                {
                    Text = rawPreview,
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 11,
                    Foreground = ThemeColors.TextSecondary,
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(0, 3, 0, 6)
                };
                itemStack.Children.Add(contentTb);

                // Dòng nút thao tác: Chèn vào Word & Xóa (nếu là mẫu tùy chỉnh)
                Grid actionGrid = new Grid();
                actionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                actionGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                Button btnInsert = ThemeColors.CreatePrimaryButton("➕ Chèn Vào Word", "Chèn nội dung mẫu này vào vị trí con trỏ");
                btnInsert.Padding = new Thickness(8, 3, 8, 3);
                btnInsert.Click += (s, e) =>
                {
                    string msg;
                    _wordService.InsertTextSnippet(captured.Content, out msg);
                    _notifyStatus(msg);
                };
                Grid.SetColumn(btnInsert, 0);
                actionGrid.Children.Add(btnInsert);

                if (!captured.IsBuiltIn)
                {
                    Button btnDel = ThemeColors.CreateAccentOutlineButton("🗑️ Xóa", "Xóa mẫu cá nhân này");
                    btnDel.Padding = new Thickness(8, 3, 8, 3);
                    btnDel.Margin = new Thickness(6, 0, 0, 0);
                    btnDel.Click += (s, e) =>
                    {
                        string msg;
                        QuickInputService.DeleteCustomSnippet(captured.Id, out msg);
                        _notifyStatus(msg);
                        RefreshSnippetList();
                    };
                    Grid.SetColumn(btnDel, 1);
                    actionGrid.Children.Add(btnDel);
                }

                itemStack.Children.Add(actionGrid);
                itemCard.Child = itemStack;
                _snippetListPanel.Children.Add(itemCard);
            }
        }
    }
}
