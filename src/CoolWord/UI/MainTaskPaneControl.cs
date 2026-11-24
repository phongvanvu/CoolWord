using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using CoolWord.Core;

namespace CoolWord.UI
{
    /// <summary>
    /// Giao diện bảng điều khiển tác vụ chính (TaskPane) của CoolWord.
    /// Thiết kế lại theo chuẩn Microsoft Fluent Design / Office 365 Add-in hiện đại:
    /// - Không còn hàng nút chen chúc chiếm diện tích.
    /// - Chế độ "Danh mục công cụ (Hub)": Tìm kiếm nhanh tức thì, phân loại 4 nhóm trực quan, thẻ tính năng đẹp mắt.
    /// - Chế độ "Thao tác chi tiết (Focus)": Tối ưu 100% chiều cao cho công cụ đang dùng, kèm thanh chọn nhanh tiện ích.
    /// </summary>
    public class MainTaskPaneControl : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Dictionary<string, UIElement> _moduleViews = new Dictionary<string, UIElement>();

        // Layout controls
        private TextBlock _headerSubtitle;
        private Button _btnToggleCatalog;

        private Grid _catalogPanel;
        private Grid _modulePanel;
        private ContentControl _moduleHost;
        private ComboBox _moduleSelector;
        private TextBlock _activeCategoryBadge;

        private TextBox _searchBox;
        private TextBlock _searchPlaceholder;
        private Button _btnClearSearch;
        private StackPanel _cardsStackPanel;
        private TextBlock _statusTextBlock;

        private string _activeModuleId;
        private string _selectedCategoryFilter = "all";
        private bool _isUpdatingSelector;

        public const string CategorySoanThao = "Soạn thảo & Thể thức";
        public const string CategoryKeToan = "Số & Kế toán";
        public const string CategoryChuyenDoi = "Chuyển đổi & Xuất bản";
        public const string CategoryTienIch = "Tiện ích mở rộng";

        public MainTaskPaneControl(IWordDocumentService wordService)
        {
            _wordService = wordService ?? new WordDocumentService();
            ModuleRegistry.EnsureInitialized();

            BuildLayout();

            // Mặc định khởi động vào công cụ "Đổi số ra chữ", người dùng có thể bấm [⊞ Danh mục] bất kỳ lúc nào
            SelectModule("number-to-text");
        }

        public string ActiveModuleId
        {
            get { return _activeModuleId; }
        }

        public bool IsInCatalogView
        {
            get { return _catalogPanel != null && _catalogPanel.Visibility == Visibility.Visible; }
        }

        /// <summary>
        /// Mở một module cụ thể và tối đa hóa không gian thao tác cho module đó.
        /// </summary>
        public void SelectModule(string moduleId)
        {
            if (string.Equals(moduleId, "catalog", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(moduleId, "home", StringComparison.OrdinalIgnoreCase))
            {
                ShowCatalog();
                return;
            }

            IToolModule module = ModuleRegistry.GetById(moduleId);
            if (module == null)
            {
                var all = ModuleRegistry.GetAllModules();
                if (all.Count == 0) return;
                module = all[0];
            }

            _activeModuleId = module.Id;

            // Lấy hoặc khởi tạo view của module
            UIElement view;
            if (!_moduleViews.TryGetValue(module.Id, out view))
            {
                view = module.CreateModuleView(_wordService, UpdateStatus);
                _moduleViews[module.Id] = view;
            }

            _moduleHost.Content = view;

            // Chuyển sang hiển thị Module Panel, ẩn Catalog Panel
            _catalogPanel.Visibility = Visibility.Collapsed;
            _modulePanel.Visibility = Visibility.Visible;

            // Cập nhật Header
            if (_headerSubtitle != null)
            {
                _headerSubtitle.Text = "Đang dùng: " + module.Title;
            }
            if (_btnToggleCatalog != null)
            {
                _btnToggleCatalog.Content = "⊞ Danh mục";
                _btnToggleCatalog.ToolTip = "Mở danh mục tất cả 15 công cụ";
            }

            // Cập nhật ComboBox bộ chọn nhanh (tránh lặp sự kiện)
            UpdateSelectorSelection(module.Id);

            // Cập nhật Badge danh mục bên cạnh bộ chọn nhanh
            if (_activeCategoryBadge != null)
            {
                string groupName = GetPrimaryGroup(module.Id);
                _activeCategoryBadge.Text = groupName;
                SolidColorBrush fg, bg, border;
                ThemeColors.GetCategoryBrushes(groupName, out fg, out bg, out border);
                _activeCategoryBadge.Foreground = fg;
                Border badgeParent = _activeCategoryBadge.Parent as Border;
                if (badgeParent != null)
                {
                    badgeParent.Background = bg;
                    badgeParent.BorderBrush = border;
                }
            }

            UpdateStatus("Đang mở: " + module.Title + " (" + module.Subtitle + ")");
        }

        /// <summary>
        /// Chuyển sang màn hình Danh mục tất cả công cụ (Hub/Catalog) với thanh tìm kiếm và lọc.
        /// </summary>
        public void ShowCatalog()
        {
            _modulePanel.Visibility = Visibility.Collapsed;
            _catalogPanel.Visibility = Visibility.Visible;

            if (_headerSubtitle != null)
            {
                _headerSubtitle.Text = "Danh mục 15 tiện ích Word";
            }
            if (_btnToggleCatalog != null)
            {
                _btnToggleCatalog.Content = !string.IsNullOrEmpty(_activeModuleId) ? "‹ Quay lại" : "⊞ Danh mục";
                _btnToggleCatalog.ToolTip = "Quay lại tiện ích đang mở";
            }

            UpdateStatus("Chọn một tiện ích bên dưới hoặc gõ từ khóa vào ô tìm kiếm.");
            RefreshCatalogCards();
        }

        /// <summary>
        /// Cập nhật thông báo lên thanh trạng thái dưới cùng.
        /// </summary>
        public void UpdateStatus(string message)
        {
            if (string.IsNullOrWhiteSpace(message)) return;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_statusTextBlock != null)
                {
                    _statusTextBlock.Text = message;
                }
            }));
        }

        #region Xây dựng giao diện Fluent UI

        private void BuildLayout()
        {
            Background = ThemeColors.AppBackground;

            Grid rootGrid = new Grid();
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 0: Top Header Banner (46px)
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // 1: Main Workspace (Catalog / Module)
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 2: Bottom Status Bar (28px)

            // 1. Top Header Banner
            rootGrid.Children.Add(CreateHeaderBanner());
            Grid.SetRow((UIElement)rootGrid.Children[0], 0);

            // 2. Main Workspace Container
            Grid workspaceGrid = new Grid();
            _catalogPanel = CreateCatalogPanel();
            _modulePanel = CreateModulePanel();

            workspaceGrid.Children.Add(_catalogPanel);
            workspaceGrid.Children.Add(_modulePanel);

            Grid.SetRow(workspaceGrid, 1);
            rootGrid.Children.Add(workspaceGrid);

            // 3. Bottom Status Bar
            Border statusBar = CreateStatusBar();
            Grid.SetRow(statusBar, 2);
            rootGrid.Children.Add(statusBar);

            Content = rootGrid;
        }

        private UIElement CreateHeaderBanner()
        {
            Border headerBorder = new Border
            {
                Background = ThemeColors.WordBlue,
                Padding = new Thickness(12, 9, 12, 9),
                BorderThickness = new Thickness(0, 0, 0, 1),
                BorderBrush = ThemeColors.BrushFromHex("#103F91")
            };

            Grid headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // Logo
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Titles
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // Catalog button

            // Logo Badge
            Border logoBadge = new Border
            {
                Width = 32,
                Height = 32,
                CornerRadius = new CornerRadius(6),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 10, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = "CW",
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 13,
                    FontWeight = FontWeights.ExtraBold,
                    Foreground = ThemeColors.WordBlue,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                }
            };
            Grid.SetColumn(logoBadge, 0);
            headerGrid.Children.Add(logoBadge);

            // App Title & Subtitle
            StackPanel titleStack = new StackPanel
            {
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0)
            };
            TextBlock titleTb = new TextBlock
            {
                Text = "CoolWord Studio",
                FontFamily = ThemeColors.MainFont,
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White
            };
            _headerSubtitle = new TextBlock
            {
                Text = "Danh mục 15 tiện ích Word",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11,
                Foreground = ThemeColors.BrushFromHex("#DCEBFA"),
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            titleStack.Children.Add(titleTb);
            titleStack.Children.Add(_headerSubtitle);
            Grid.SetColumn(titleStack, 1);
            headerGrid.Children.Add(titleStack);

            // Nút chuyển nhanh giữa Danh mục và Công cụ chi tiết
            _btnToggleCatalog = new Button
            {
                Content = "⊞ Danh mục",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brushes.White,
                Background = ThemeColors.BrushFromHex("#103F91"),
                BorderBrush = ThemeColors.BrushFromHex("#2A6CD6"),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(9, 5, 9, 5),
                Cursor = Cursors.Hand,
                VerticalAlignment = VerticalAlignment.Center,
                ToolTip = "Mở bảng danh mục 15 công cụ"
            };
            _btnToggleCatalog.Click += (s, e) =>
            {
                if (IsInCatalogView)
                {
                    if (!string.IsNullOrEmpty(_activeModuleId))
                    {
                        SelectModule(_activeModuleId);
                    }
                }
                else
                {
                    ShowCatalog();
                }
            };
            Grid.SetColumn(_btnToggleCatalog, 2);
            headerGrid.Children.Add(_btnToggleCatalog);

            headerBorder.Child = headerGrid;
            return headerBorder;
        }

        #endregion

        #region Màn hình Danh mục Công cụ (Hub / Catalog)

        private Grid CreateCatalogPanel()
        {
            Grid grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Search & Filters
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Quick Access Chips
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Cards List

            // 1. Ô tìm kiếm & Thanh lọc danh mục
            Border filterContainer = new Border
            {
                Background = Brushes.White,
                BorderBrush = ThemeColors.BorderBrush,
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(12, 10, 12, 10)
            };

            StackPanel filterStack = new StackPanel();

            // Search Bar
            Border searchBorder = new Border
            {
                Background = ThemeColors.SubtleBackground,
                BorderBrush = ThemeColors.InputBorderBrush,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(8, 5, 8, 5),
                Margin = new Thickness(0, 0, 0, 8)
            };

            Grid searchGrid = new Grid();
            searchGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // Search icon
            searchGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // TextBox
            searchGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // Clear button

            TextBlock searchIcon = new TextBlock
            {
                Text = "🔍",
                FontSize = 11,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 6, 0),
                Foreground = ThemeColors.TextMuted
            };
            Grid.SetColumn(searchIcon, 0);
            searchGrid.Children.Add(searchIcon);

            Grid inputGrid = new Grid();
            _searchBox = new TextBox
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 12,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                VerticalAlignment = VerticalAlignment.Center
            };
            _searchPlaceholder = new TextBlock
            {
                Text = "Tìm nhanh công cụ (mục lục, qr, số, in, mail...)",
                FontFamily = ThemeColors.MainFont,
                FontSize = 12,
                Foreground = ThemeColors.TextMuted,
                VerticalAlignment = VerticalAlignment.Center,
                IsHitTestVisible = false
            };
            inputGrid.Children.Add(_searchPlaceholder);
            inputGrid.Children.Add(_searchBox);
            Grid.SetColumn(inputGrid, 1);
            searchGrid.Children.Add(inputGrid);

            _btnClearSearch = new Button
            {
                Content = "✕",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11,
                Foreground = ThemeColors.TextMuted,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                Visibility = Visibility.Collapsed,
                Padding = new Thickness(4, 0, 4, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            _btnClearSearch.Click += (s, e) =>
            {
                _searchBox.Text = string.Empty;
                _searchBox.Focus();
            };
            Grid.SetColumn(_btnClearSearch, 2);
            searchGrid.Children.Add(_btnClearSearch);

            _searchBox.TextChanged += (s, e) =>
            {
                bool hasText = !string.IsNullOrEmpty(_searchBox.Text);
                _searchPlaceholder.Visibility = hasText ? Visibility.Collapsed : Visibility.Visible;
                _btnClearSearch.Visibility = hasText ? Visibility.Visible : Visibility.Collapsed;
                RefreshCatalogCards();
            };

            searchBorder.Child = searchGrid;
            filterStack.Children.Add(searchBorder);

            // Category Filter Pills
            WrapPanel pillWrap = new WrapPanel { Margin = new Thickness(0, 2, 0, 0) };
            pillWrap.Children.Add(CreateCategoryPill("all", "Tất cả (15)"));
            pillWrap.Children.Add(CreateCategoryPill("soan-thao", "📝 Soạn thảo (5)"));
            pillWrap.Children.Add(CreateCategoryPill("ke-toan", "💰 Kế toán (1)"));
            pillWrap.Children.Add(CreateCategoryPill("chuyen-doi", "🔄 Chuyển đổi (5)"));
            pillWrap.Children.Add(CreateCategoryPill("tien-ich", "⚡ Tiện ích (4)"));

            filterStack.Children.Add(pillWrap);
            filterContainer.Child = filterStack;

            Grid.SetRow(filterContainer, 0);
            grid.Children.Add(filterContainer);

            // 2. Thanh truy cập nhanh (Quick Chips)
            Border quickAccessBorder = new Border
            {
                Background = ThemeColors.BrushFromHex("#F0F6FF"),
                BorderBrush = ThemeColors.WordBlueBorder,
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(12, 6, 12, 6)
            };

            WrapPanel quickChipsWrap = new WrapPanel();
            quickChipsWrap.Children.Add(new TextBlock
            {
                Text = "Gợi ý nhanh:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11,
                FontWeight = FontWeights.Bold,
                Foreground = ThemeColors.WordBlue,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0)
            });
            quickChipsWrap.Children.Add(CreateQuickChip("number-to-text", "🔢 Đổi số ra chữ"));
            quickChipsWrap.Children.Add(CreateQuickChip("admin-decree30", "📑 Nghị định 30"));
            quickChipsWrap.Children.Add(CreateQuickChip("table-of-contents", "📋 Tạo Mục Lục"));
            quickChipsWrap.Children.Add(CreateQuickChip("qr-code", "📱 Mã QR Code"));

            quickAccessBorder.Child = quickChipsWrap;
            Grid.SetRow(quickAccessBorder, 1);
            grid.Children.Add(quickAccessBorder);

            // 3. Danh sách thẻ công cụ
            ScrollViewer scroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Padding = new Thickness(12, 10, 12, 10)
            };
            _cardsStackPanel = new StackPanel();
            scroll.Content = _cardsStackPanel;

            Grid.SetRow(scroll, 2);
            grid.Children.Add(scroll);

            return grid;
        }

        private Border CreateCategoryPill(string filterKey, string label)
        {
            bool isSelected = string.Equals(_selectedCategoryFilter, filterKey, StringComparison.OrdinalIgnoreCase);

            Border pill = new Border
            {
                CornerRadius = new CornerRadius(12),
                Padding = new Thickness(9, 4, 9, 4),
                Margin = new Thickness(0, 0, 5, 4),
                Background = isSelected ? ThemeColors.WordBlue : ThemeColors.SubtleBackground,
                BorderBrush = isSelected ? ThemeColors.WordBlue : ThemeColors.BorderBrush,
                BorderThickness = new Thickness(1),
                Cursor = Cursors.Hand,
                Tag = filterKey
            };

            TextBlock tb = new TextBlock
            {
                Text = label,
                FontFamily = ThemeColors.MainFont,
                FontSize = 11,
                FontWeight = isSelected ? FontWeights.Bold : FontWeights.Normal,
                Foreground = isSelected ? Brushes.White : ThemeColors.TextSecondary
            };
            pill.Child = tb;

            pill.MouseLeftButtonUp += (s, e) =>
            {
                _selectedCategoryFilter = filterKey;
                UpdatePillsVisual();
                RefreshCatalogCards();
            };

            return pill;
        }

        private void UpdatePillsVisual()
        {
            if (_catalogPanel == null) return;
            // Duyệt các pill trong WrapPanel để cập nhật visual
            Border filterContainer = _catalogPanel.Children[0] as Border;
            if (filterContainer == null) return;
            StackPanel filterStack = filterContainer.Child as StackPanel;
            if (filterStack == null || filterStack.Children.Count < 2) return;
            WrapPanel pillWrap = filterStack.Children[1] as WrapPanel;
            if (pillWrap == null) return;

            foreach (UIElement child in pillWrap.Children)
            {
                Border pill = child as Border;
                if (pill == null) continue;
                string key = pill.Tag as string;
                bool isSelected = string.Equals(_selectedCategoryFilter, key, StringComparison.OrdinalIgnoreCase);

                pill.Background = isSelected ? ThemeColors.WordBlue : ThemeColors.SubtleBackground;
                pill.BorderBrush = isSelected ? ThemeColors.WordBlue : ThemeColors.BorderBrush;
                TextBlock tb = pill.Child as TextBlock;
                if (tb != null)
                {
                    tb.FontWeight = isSelected ? FontWeights.Bold : FontWeights.Normal;
                    tb.Foreground = isSelected ? Brushes.White : ThemeColors.TextSecondary;
                }
            }
        }

        private Border CreateQuickChip(string moduleId, string text)
        {
            Border chip = new Border
            {
                CornerRadius = new CornerRadius(4),
                Background = Brushes.White,
                BorderBrush = ThemeColors.WordBlueBorder,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(7, 2, 7, 2),
                Margin = new Thickness(0, 0, 5, 0),
                Cursor = Cursors.Hand,
                ToolTip = "Mở nhanh tính năng này"
            };

            TextBlock tb = new TextBlock
            {
                Text = text,
                FontFamily = ThemeColors.MainFont,
                FontSize = 11,
                FontWeight = FontWeights.Medium,
                Foreground = ThemeColors.WordBlue
            };
            chip.Child = tb;

            chip.MouseEnter += (s, e) => chip.Background = ThemeColors.WordBlueLight;
            chip.MouseLeave += (s, e) => chip.Background = Brushes.White;
            chip.MouseLeftButtonUp += (s, e) => SelectModule(moduleId);

            return chip;
        }

        private void RefreshCatalogCards()
        {
            if (_cardsStackPanel == null) return;
            _cardsStackPanel.Children.Clear();

            string query = (_searchBox != null ? _searchBox.Text : string.Empty).Trim().ToLowerInvariant();
            var allModules = ModuleRegistry.GetAllModules();

            // Nhóm theo 4 nhóm chính
            string[] groupNames = new string[]
            {
                CategorySoanThao,
                CategoryKeToan,
                CategoryChuyenDoi,
                CategoryTienIch
            };

            int totalMatched = 0;

            foreach (string group in groupNames)
            {
                // Kiểm tra xem nhóm này có bị lọc bởi Category Pill không
                if (!IsGroupAllowedByFilter(group, _selectedCategoryFilter))
                    continue;

                List<IToolModule> modulesInGroup = new List<IToolModule>();
                foreach (IToolModule mod in allModules)
                {
                    if (string.Equals(GetPrimaryGroup(mod.Id), group, StringComparison.OrdinalIgnoreCase))
                    {
                        // Kiểm tra từ khóa tìm kiếm
                        if (string.IsNullOrEmpty(query) || MatchesSearch(mod, query))
                        {
                            modulesInGroup.Add(mod);
                        }
                    }
                }

                if (modulesInGroup.Count == 0) continue;

                totalMatched += modulesInGroup.Count;

                // Thêm Section Header cho nhóm
                _cardsStackPanel.Children.Add(CreateGroupSectionHeader(group, modulesInGroup.Count));

                // Thêm từng Card
                foreach (IToolModule mod in modulesInGroup)
                {
                    _cardsStackPanel.Children.Add(CreateModuleCard(mod, group));
                }
            }

            // Nếu tìm kiếm không thấy kết quả
            if (totalMatched == 0)
            {
                Border emptyCard = new Border
                {
                    Background = Brushes.White,
                    BorderBrush = ThemeColors.BorderBrush,
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(8),
                    Padding = new Thickness(20),
                    Margin = new Thickness(0, 20, 0, 0)
                };
                StackPanel emptyStack = new StackPanel { HorizontalAlignment = HorizontalAlignment.Center };
                emptyStack.Children.Add(new TextBlock
                {
                    Text = "🔍 Không tìm thấy tiện ích phù hợp",
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 13,
                    FontWeight = FontWeights.Bold,
                    Foreground = ThemeColors.TextPrimary,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 0, 0, 6)
                });
                emptyStack.Children.Add(new TextBlock
                {
                    Text = "Thử tìm với từ khóa khác (vd: 'số', 'qr', 'mục lục', 'in', 'word').",
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 11.5,
                    Foreground = ThemeColors.TextSecondary,
                    HorizontalAlignment = HorizontalAlignment.Center
                });
                emptyCard.Child = emptyStack;
                _cardsStackPanel.Children.Add(emptyCard);
            }
        }

        private bool IsGroupAllowedByFilter(string group, string filterKey)
        {
            if (string.IsNullOrEmpty(filterKey) || string.Equals(filterKey, "all", StringComparison.OrdinalIgnoreCase))
                return true;

            if (filterKey == "soan-thao" && group == CategorySoanThao) return true;
            if (filterKey == "ke-toan" && group == CategoryKeToan) return true;
            if (filterKey == "chuyen-doi" && group == CategoryChuyenDoi) return true;
            if (filterKey == "tien-ich" && group == CategoryTienIch) return true;

            return false;
        }

        private bool MatchesSearch(IToolModule mod, string query)
        {
            if (mod.Title != null && mod.Title.ToLowerInvariant().Contains(query)) return true;
            if (mod.Subtitle != null && mod.Subtitle.ToLowerInvariant().Contains(query)) return true;
            if (mod.IconText != null && mod.IconText.ToLowerInvariant().Contains(query)) return true;
            if (mod.Category != null && mod.Category.ToLowerInvariant().Contains(query)) return true;
            if (mod.Id != null && mod.Id.ToLowerInvariant().Contains(query)) return true;

            // Từ khóa tìm kiếm tiếng Việt thông minh
            if (query.Contains("qr") && mod.Id == "qr-code") return true;
            if (query.Contains("mục lục") && mod.Id == "table-of-contents") return true;
            if ((query.Contains("mail") || query.Contains("thư")) && mod.Id == "batch-email") return true;
            if (query.Contains("ảnh") && mod.Id == "image-tools") return true;
            if (query.Contains("dịch") && mod.Id == "google-translate") return true;
            if ((query.Contains("số") || query.Contains("tiền")) && mod.Id == "number-to-text") return true;
            if (query.Contains("30") && mod.Id == "admin-decree30") return true;
            if (query.Contains("excel") && mod.Id == "word-to-excel") return true;
            if (query.Contains("ppt") && mod.Id == "word-to-ppt") return true;
            if (query.Contains("pdf") && mod.Id == "pdf-utilities") return true;
            if (query.Contains("in") && mod.Id == "print-tools") return true;
            if (query.Contains("tách") && mod.Id == "name-split") return true;
            if (query.Contains("nhập") && mod.Id == "quick-input") return true;

            return false;
        }

        private UIElement CreateGroupSectionHeader(string groupName, int count)
        {
            Border header = new Border
            {
                Margin = new Thickness(0, 10, 0, 6),
                Padding = new Thickness(0, 0, 0, 4),
                BorderThickness = new Thickness(0, 0, 0, 1),
                BorderBrush = ThemeColors.BorderBrush
            };

            Grid g = new Grid();
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            g.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            TextBlock tb = new TextBlock
            {
                Text = groupName.ToUpperInvariant(),
                FontFamily = ThemeColors.MainFont,
                FontSize = 11,
                FontWeight = FontWeights.Bold,
                Foreground = ThemeColors.TextMuted
            };
            Grid.SetColumn(tb, 0);
            g.Children.Add(tb);

            TextBlock countTb = new TextBlock
            {
                Text = count.ToString() + " công cụ",
                FontFamily = ThemeColors.MainFont,
                FontSize = 10.5,
                Foreground = ThemeColors.TextMuted
            };
            Grid.SetColumn(countTb, 1);
            g.Children.Add(countTb);

            header.Child = g;
            return header;
        }

        private UIElement CreateModuleCard(IToolModule mod, string groupName)
        {
            string capturedId = mod.Id;
            bool isActive = string.Equals(mod.Id, _activeModuleId, StringComparison.OrdinalIgnoreCase);

            Border card = new Border
            {
                Background = Brushes.White,
                BorderBrush = isActive ? ThemeColors.WordBlue : ThemeColors.BorderBrush,
                BorderThickness = new Thickness(isActive ? 1.5 : 1),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(10, 8, 10, 8),
                Margin = new Thickness(0, 0, 0, 6),
                Cursor = Cursors.Hand,
                ToolTip = "Bấm để mở công cụ " + mod.Title
            };

            Grid grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // Badge icon
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Title & desc
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // Arrow button

            // Icon Badge
            SolidColorBrush fg, bg, border;
            ThemeColors.GetCategoryBrushes(groupName, out fg, out bg, out border);

            Border iconBadge = new Border
            {
                Width = 38,
                Height = 38,
                CornerRadius = new CornerRadius(6),
                Background = bg,
                BorderBrush = border,
                BorderThickness = new Thickness(1),
                Margin = new Thickness(0, 0, 10, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = mod.IconText,
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 10.5,
                    FontWeight = FontWeights.Bold,
                    Foreground = fg,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                }
            };
            Grid.SetColumn(iconBadge, 0);
            grid.Children.Add(iconBadge);

            // Title and Description
            StackPanel textStack = new StackPanel
            {
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0)
            };

            StackPanel titleRow = new StackPanel { Orientation = Orientation.Horizontal };
            titleRow.Children.Add(new TextBlock
            {
                Text = mod.Title,
                FontFamily = ThemeColors.MainFont,
                FontSize = 12.5,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeColors.TextPrimary
            });

            if (isActive)
            {
                Border activePill = new Border
                {
                    Background = ThemeColors.WordBlueLight,
                    CornerRadius = new CornerRadius(3),
                    Padding = new Thickness(4, 1, 4, 1),
                    Margin = new Thickness(6, 0, 0, 0),
                    VerticalAlignment = VerticalAlignment.Center,
                    Child = new TextBlock
                    {
                        Text = "Đang mở",
                        FontFamily = ThemeColors.MainFont,
                        FontSize = 9.5,
                        FontWeight = FontWeights.Bold,
                        Foreground = ThemeColors.WordBlue
                    }
                };
                titleRow.Children.Add(activePill);
            }
            textStack.Children.Add(titleRow);

            textStack.Children.Add(new TextBlock
            {
                Text = mod.Subtitle,
                FontFamily = ThemeColors.MainFont,
                FontSize = 11,
                Foreground = ThemeColors.TextMuted,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 2, 0, 0)
            });

            Grid.SetColumn(textStack, 1);
            grid.Children.Add(textStack);

            // Arrow button indicator
            Border arrowBorder = new Border
            {
                Width = 24,
                Height = 24,
                CornerRadius = new CornerRadius(12),
                Background = ThemeColors.SubtleBackground,
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = "➔",
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 11,
                    FontWeight = FontWeights.Bold,
                    Foreground = ThemeColors.WordBlue,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                }
            };
            Grid.SetColumn(arrowBorder, 2);
            grid.Children.Add(arrowBorder);

            card.Child = grid;

            // Hover effects
            card.MouseEnter += (s, e) =>
            {
                card.Background = ThemeColors.WordBlueLight;
                card.BorderBrush = ThemeColors.WordBlueBorder;
                arrowBorder.Background = ThemeColors.WordBlue;
                TextBlock arrowTb = arrowBorder.Child as TextBlock;
                if (arrowTb != null) arrowTb.Foreground = Brushes.White;
            };
            card.MouseLeave += (s, e) =>
            {
                card.Background = Brushes.White;
                card.BorderBrush = isActive ? ThemeColors.WordBlue : ThemeColors.BorderBrush;
                arrowBorder.Background = ThemeColors.SubtleBackground;
                TextBlock arrowTb = arrowBorder.Child as TextBlock;
                if (arrowTb != null) arrowTb.Foreground = ThemeColors.WordBlue;
            };

            card.MouseLeftButtonUp += (s, e) => SelectModule(capturedId);

            return card;
        }

        #endregion

        #region Màn hình Thao tác Chi tiết Module (Focus View)

        private Grid CreateModulePanel()
        {
            Grid grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Sub-Header (Quick Switcher)
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Module Host

            // Sub-Header Navigation
            Border subHeader = new Border
            {
                Background = Brushes.White,
                BorderBrush = ThemeColors.BorderBrush,
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(10, 6, 10, 6)
            };

            Grid subGrid = new Grid();
            subGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // Back to Catalog
            subGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // ComboBox Switcher
            subGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // Active Category Badge

            // Nút [‹ Danh mục]
            Button btnBackToCatalog = new Button
            {
                Content = "‹ Danh mục",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                FontWeight = FontWeights.SemiBold,
                Foreground = ThemeColors.WordBlue,
                Background = ThemeColors.WordBlueLight,
                BorderBrush = ThemeColors.WordBlueBorder,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(8, 4, 8, 4),
                Cursor = Cursors.Hand,
                ToolTip = "Trở về danh mục tất cả 15 công cụ",
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0)
            };
            btnBackToCatalog.Click += (s, e) => ShowCatalog();
            Grid.SetColumn(btnBackToCatalog, 0);
            subGrid.Children.Add(btnBackToCatalog);

            // ComboBox chọn nhanh công cụ khác ngay tại chỗ
            _moduleSelector = new ComboBox
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0)
            };
            PopulateModuleSelector();
            _moduleSelector.SelectionChanged += OnModuleSelectorChanged;
            Grid.SetColumn(_moduleSelector, 1);
            subGrid.Children.Add(_moduleSelector);

            // Active Category Badge
            Border categoryBadgeBorder = new Border
            {
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(6, 3, 6, 3),
                VerticalAlignment = VerticalAlignment.Center
            };
            _activeCategoryBadge = new TextBlock
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 10,
                FontWeight = FontWeights.Bold
            };
            categoryBadgeBorder.Child = _activeCategoryBadge;
            Grid.SetColumn(categoryBadgeBorder, 2);
            subGrid.Children.Add(categoryBadgeBorder);

            subHeader.Child = subGrid;
            Grid.SetRow(subHeader, 0);
            grid.Children.Add(subHeader);

            // Module View Host inside ScrollViewer
            ScrollViewer scroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Padding = new Thickness(12, 10, 12, 8)
            };
            _moduleHost = new ContentControl();
            scroll.Content = _moduleHost;

            Grid.SetRow(scroll, 1);
            grid.Children.Add(scroll);

            return grid;
        }

        private void PopulateModuleSelector()
        {
            if (_moduleSelector == null) return;
            _moduleSelector.Items.Clear();

            var all = ModuleRegistry.GetAllModules();
            foreach (var mod in all)
            {
                string group = GetPrimaryGroup(mod.Id);
                string text = string.Format("[{0}] {1}  ({2})", mod.IconText, mod.Title, group);
                _moduleSelector.Items.Add(new ModuleComboItem { Id = mod.Id, DisplayText = text });
            }
        }

        private void UpdateSelectorSelection(string moduleId)
        {
            if (_moduleSelector == null) return;
            _isUpdatingSelector = true;
            try
            {
                for (int i = 0; i < _moduleSelector.Items.Count; i++)
                {
                    ModuleComboItem item = _moduleSelector.Items[i] as ModuleComboItem;
                    if (item != null && string.Equals(item.Id, moduleId, StringComparison.OrdinalIgnoreCase))
                    {
                        _moduleSelector.SelectedIndex = i;
                        break;
                    }
                }
            }
            finally
            {
                _isUpdatingSelector = false;
            }
        }

        private void OnModuleSelectorChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isUpdatingSelector) return;
            ModuleComboItem selected = _moduleSelector.SelectedItem as ModuleComboItem;
            if (selected != null && !string.IsNullOrEmpty(selected.Id))
            {
                SelectModule(selected.Id);
            }
        }

        #endregion

        #region Thanh trạng thái dưới cùng (Status Bar)

        private Border CreateStatusBar()
        {
            Border statusBar = new Border
            {
                Background = Brushes.White,
                BorderBrush = ThemeColors.BorderBrush,
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(12, 6, 12, 6)
            };

            Grid statusGrid = new Grid();
            statusGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            statusGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            TextBlock dotIndicator = new TextBlock
            {
                Text = "●",
                FontSize = 9,
                Foreground = ThemeColors.SuccessText,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 6, 0)
            };
            Grid.SetColumn(dotIndicator, 0);
            statusGrid.Children.Add(dotIndicator);

            _statusTextBlock = new TextBlock
            {
                Text = "Sẵn sàng. Chọn công cụ để làm việc với tài liệu Word.",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11,
                Foreground = ThemeColors.TextSecondary,
                TextTrimming = TextTrimming.CharacterEllipsis,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(_statusTextBlock, 1);
            statusGrid.Children.Add(_statusTextBlock);

            statusBar.Child = statusGrid;
            return statusBar;
        }

        #endregion

        #region Helper Groups

        public static string GetPrimaryGroup(string moduleId)
        {
            if (string.IsNullOrEmpty(moduleId)) return CategoryTienIch;

            switch (moduleId)
            {
                case "admin-decree30":
                case "text-clean":
                case "quick-input":
                case "name-split":
                case "table-of-contents":
                    return CategorySoanThao;

                case "number-to-text":
                    return CategoryKeToan;

                case "word-to-word":
                case "word-to-excel":
                case "word-to-ppt":
                case "print-tools":
                case "pdf-utilities":
                    return CategoryChuyenDoi;

                case "batch-email":
                case "image-tools":
                case "google-translate":
                case "qr-code":
                    return CategoryTienIch;

                default:
                    return CategoryTienIch;
            }
        }

        private class ModuleComboItem
        {
            public string Id { get; set; }
            public string DisplayText { get; set; }

            public override string ToString()
            {
                return DisplayText;
            }
        }

        #endregion
    }
}
