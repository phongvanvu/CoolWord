using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace CoolWord.UI
{
    public static class ThemeColors
    {
        public static readonly SolidColorBrush WordBlue = BrushFromHex("#185ABD");
        public static readonly SolidColorBrush WordBlueHover = BrushFromHex("#103F91");
        public static readonly SolidColorBrush WordBluePressed = BrushFromHex("#0A2E6E");
        public static readonly SolidColorBrush WordBlueLight = BrushFromHex("#EBF3FC");
        public static readonly SolidColorBrush WordBlueBorder = BrushFromHex("#B4D6FA");

        public static readonly SolidColorBrush AppBackground = BrushFromHex("#F4F6F9");
        public static readonly SolidColorBrush CardBackground = BrushFromHex("#FFFFFF");
        public static readonly SolidColorBrush SubtleBackground = BrushFromHex("#F8FAFC");
        public static readonly SolidColorBrush BorderBrush = BrushFromHex("#E2E8F0");
        public static readonly SolidColorBrush InputBorderBrush = BrushFromHex("#CBD5E1");

        public static readonly SolidColorBrush TextPrimary = BrushFromHex("#0F172A");
        public static readonly SolidColorBrush TextSecondary = BrushFromHex("#475569");
        public static readonly SolidColorBrush TextMuted = BrushFromHex("#64748B");

        public static readonly SolidColorBrush SuccessText = BrushFromHex("#047857");
        public static readonly SolidColorBrush SuccessBg = BrushFromHex("#ECFDF5");
        public static readonly SolidColorBrush SuccessBorder = BrushFromHex("#A7F3D0");

        public static readonly FontFamily MainFont = new FontFamily("Segoe UI");

        public static SolidColorBrush BrushFromHex(string hex)
        {
            Color color = (Color)ColorConverter.ConvertFromString(hex);
            SolidColorBrush brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }

        public static Border CreateCard(Thickness padding, Thickness margin)
        {
            return new Border
            {
                Background = CardBackground,
                BorderBrush = BorderBrush,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(8),
                Padding = padding,
                Margin = margin
            };
        }

        public static TextBlock CreateSectionHeader(string title, string subtitle = null)
        {
            TextBlock tb = new TextBlock
            {
                Text = title,
                FontFamily = MainFont,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = TextPrimary,
                Margin = new Thickness(0, 0, 0, string.IsNullOrEmpty(subtitle) ? 8 : 2)
            };
            return tb;
        }

        public static TextBlock CreateFieldLabel(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontFamily = MainFont,
                FontSize = 12,
                FontWeight = FontWeights.Medium,
                Foreground = TextSecondary,
                Margin = new Thickness(0, 0, 0, 4)
            };
        }

        public static Button CreatePrimaryButton(string text, string tooltip = null)
        {
            return CreateStyledButton(text, WordBlue, WordBlueHover, Brushes.White, WordBlue, tooltip, true);
        }

        public static Button CreateSecondaryButton(string text, string tooltip = null)
        {
            return CreateStyledButton(text, CardBackground, WordBlueLight, TextPrimary, InputBorderBrush, tooltip, false);
        }

        public static Button CreateAccentOutlineButton(string text, string tooltip = null)
        {
            return CreateStyledButton(text, WordBlueLight, BrushFromHex("#DCEBFA"), WordBlue, WordBlueBorder, tooltip, true);
        }

        private static Button CreateStyledButton(
            string text,
            Brush normalBg,
            Brush hoverBg,
            Brush fg,
            Brush borderBrush,
            string tooltip,
            bool semiBold)
        {
            Button btn = new Button
            {
                Content = text,
                FontFamily = MainFont,
                FontSize = 12.5,
                FontWeight = semiBold ? FontWeights.SemiBold : FontWeights.Normal,
                Foreground = fg,
                Background = normalBg,
                BorderBrush = borderBrush,
                BorderThickness = new Thickness(1),
                Padding = new Thickness(12, 7, 12, 7),
                Cursor = System.Windows.Input.Cursors.Hand,
                ToolTip = tooltip
            };

            ControlTemplate template = new ControlTemplate(typeof(Button));
            FrameworkElementFactory border = new FrameworkElementFactory(typeof(Border));
            border.Name = "BtnBorder";
            border.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Control.BackgroundProperty));
            border.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Control.BorderBrushProperty));
            border.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Control.BorderThicknessProperty));
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(6));
            border.SetValue(Border.PaddingProperty, new TemplateBindingExtension(Control.PaddingProperty));

            FrameworkElementFactory presenter = new FrameworkElementFactory(typeof(ContentPresenter));
            presenter.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            presenter.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            border.AppendChild(presenter);

            template.VisualTree = border;

            Trigger hoverTrigger = new Trigger
            {
                Property = UIElement.IsMouseOverProperty,
                Value = true
            };
            hoverTrigger.Setters.Add(new Setter(Border.BackgroundProperty, hoverBg, "BtnBorder"));
            template.Triggers.Add(hoverTrigger);

            btn.Template = template;
            return btn;
        }
    }
}
