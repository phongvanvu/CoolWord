using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using CoolWord.Core;

namespace CoolWord.UI
{
    public class MainTaskPaneControl : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Dictionary<string, UIElement> _moduleViews = new Dictionary<string, UIElement>();
        private readonly Dictionary<string, Border> _tabButtons = new Dictionary<string, Border>();

        private ContentControl _moduleHost;
        private TextBlock _statusTextBlock;
        private string _activeModuleId;

        public MainTaskPaneControl(IWordDocumentService wordService)
        {
            _wordService = wordService ?? new WordDocumentService();
            ModuleRegistry.EnsureInitialized();

            BuildLayout();
            SelectModule("number-to-text");
        }

        public void SelectModule(string moduleId)
        {
            IToolModule module = ModuleRegistry.GetById(moduleId);
            if (module == null)
            {
                var all = ModuleRegistry.GetAllModules();
                if (all.Count == 0) return;
                module = all[0];
            }

            _activeModuleId = module.Id;

            UIElement view;
            if (!_moduleViews.TryGetValue(module.Id, out view))
            {
                view = module.CreateModuleView(_wordService, UpdateStatus);
                _moduleViews[module.Id] = view;
            }

            _moduleHost.Content = view;
            UpdateTabVisuals();
            UpdateStatus("Đang mở công cụ: " + module.Title);
        }

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

        private void BuildLayout()
        {
            Background = ThemeColors.AppBackground;

            Grid mainGrid = new Grid();
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Header
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Module Nav
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Body
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Status bar

            // 1. Header Banner
            Border headerBorder = new Border
            {
                Background = ThemeColors.WordBlue,
                Padding = new Thickness(14, 12, 14, 12)
            };

            Grid headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Border logoBadge = new Border
            {
                Width = 34,
                Height = 34,
                CornerRadius = new CornerRadius(8),
                Background = Brushes.White,
                Margin = new Thickness(0, 0, 10, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = "CW",
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 14,
                    FontWeight = FontWeights.ExtraBold,
                    Foreground = ThemeColors.WordBlue,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                }
            };
            Grid.SetColumn(logoBadge, 0);
            headerGrid.Children.Add(logoBadge);

            StackPanel titleStack = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            titleStack.Children.Add(new TextBlock
            {
                Text = "CoolWord",
                FontFamily = ThemeColors.MainFont,
                FontSize = 15,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White
            });
            titleStack.Children.Add(new TextBlock
            {
                Text = "Tiện ích Tiếng Việt cho MS Word 2016+",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11,
                Foreground = ThemeColors.BrushFromHex("#DCEBFA")
            });
            Grid.SetColumn(titleStack, 1);
            headerGrid.Children.Add(titleStack);

            headerBorder.Child = headerGrid;
            Grid.SetRow(headerBorder, 0);
            mainGrid.Children.Add(headerBorder);

            // 2. Thanh chuyển đổi Module (Plugin Navigation Bar)
            Border navContainer = new Border
            {
                Background = Brushes.White,
                BorderBrush = ThemeColors.BorderBrush,
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(10, 8, 10, 8)
            };

            WrapPanel navWrap = new WrapPanel();
            foreach (IToolModule mod in ModuleRegistry.GetAllModules())
            {
                string capturedId = mod.Id;
                Border tab = new Border
                {
                    CornerRadius = new CornerRadius(6),
                    Padding = new Thickness(10, 6, 10, 6),
                    Margin = new Thickness(0, 0, 6, 0),
                    BorderThickness = new Thickness(1),
                    Cursor = System.Windows.Input.Cursors.Hand,
                    ToolTip = mod.Subtitle
                };

                StackPanel tabContent = new StackPanel { Orientation = Orientation.Horizontal };
                Border badge = new Border
                {
                    CornerRadius = new CornerRadius(4),
                    Background = ThemeColors.WordBlueLight,
                    Padding = new Thickness(5, 1, 5, 1),
                    Margin = new Thickness(0, 0, 6, 0),
                    Child = new TextBlock
                    {
                        Text = mod.IconText,
                        FontFamily = ThemeColors.MainFont,
                        FontSize = 10,
                        FontWeight = FontWeights.Bold,
                        Foreground = ThemeColors.WordBlue
                    }
                };
                tabContent.Children.Add(badge);

                TextBlock titleTb = new TextBlock
                {
                    Text = mod.Title,
                    FontFamily = ThemeColors.MainFont,
                    FontSize = 12,
                    FontWeight = FontWeights.SemiBold,
                    VerticalAlignment = VerticalAlignment.Center
                };
                tabContent.Children.Add(titleTb);

                tab.Child = tabContent;
                tab.MouseLeftButtonUp += (s, e) => SelectModule(capturedId);

                _tabButtons[mod.Id] = tab;
                navWrap.Children.Add(tab);
            }

            navContainer.Child = navWrap;
            Grid.SetRow(navContainer, 1);
            mainGrid.Children.Add(navContainer);

            // 3. Nội dung Module (ScrollViewer)
            ScrollViewer scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Padding = new Thickness(12, 12, 12, 8)
            };
            _moduleHost = new ContentControl();
            scrollViewer.Content = _moduleHost;
            Grid.SetRow(scrollViewer, 2);
            mainGrid.Children.Add(scrollViewer);

            // 4. Thanh trạng thái dưới cùng (Status Bar)
            Border statusBar = new Border
            {
                Background = Brushes.White,
                BorderBrush = ThemeColors.BorderBrush,
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(12, 7, 12, 7)
            };
            _statusTextBlock = new TextBlock
            {
                Text = "Sẵn sàng. Bôi đen số trong Word để chuyển đổi nhanh.",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            statusBar.Child = _statusTextBlock;
            Grid.SetRow(statusBar, 3);
            mainGrid.Children.Add(statusBar);

            Content = mainGrid;
        }

        private void UpdateTabVisuals()
        {
            foreach (var kvp in _tabButtons)
            {
                bool isActive = string.Equals(kvp.Key, _activeModuleId, StringComparison.OrdinalIgnoreCase);
                Border tab = kvp.Value;
                tab.Background = isActive ? ThemeColors.WordBlueLight : Brushes.Transparent;
                tab.BorderBrush = isActive ? ThemeColors.WordBlueBorder : Brushes.Transparent;

                StackPanel sp = tab.Child as StackPanel;
                if (sp != null && sp.Children.Count > 1)
                {
                    TextBlock tb = sp.Children[1] as TextBlock;
                    if (tb != null)
                    {
                        tb.Foreground = isActive ? ThemeColors.WordBlue : ThemeColors.TextSecondary;
                    }
                }
            }
        }
    }
}
