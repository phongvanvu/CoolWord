using System;
using System.Windows;
using System.Windows.Controls;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Modules.GoogleTranslate
{
    public class GoogleTranslateModuleView : UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Action<string> _notifyStatus;

        private ComboBox _cboTargetLang;
        private CheckBox _chkReplaceSelection;

        public GoogleTranslateModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            _wordService = wordService;
            _notifyStatus = notifyStatus ?? (_ => { });

            BuildUi();
        }

        private void BuildUi()
        {
            StackPanel root = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };

            // 1. Card: Dịch qua Google Dịch
            Border transCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel trnStack = new StackPanel();

            trnStack.Children.Add(ThemeColors.CreateSectionHeader("Google Dịch (Google Translate)"));
            trnStack.Children.Add(new TextBlock
            {
                Text = "Dịch thuật văn bản tự động qua Google Dịch, hỗ trợ hơn 100 ngôn ngữ, hoạt động thông minh cả Online lẫn Offline:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            trnStack.Children.Add(new TextBlock { Text = "Ngôn ngữ đích (Dịch sang):", FontSize = 11, FontFamily = ThemeColors.MainFont, Margin = new Thickness(0, 2, 0, 2) });
            _cboTargetLang = new ComboBox { Height = 26, Margin = new Thickness(0, 0, 0, 6), FontFamily = ThemeColors.MainFont, FontSize = 12 };
            _cboTargetLang.Items.Add(new ComboBoxItem { Content = "🇻🇳 Tiếng Việt (Vietnamese)", Tag = "vi", IsSelected = true });
            _cboTargetLang.Items.Add(new ComboBoxItem { Content = "🇬🇧 Tiếng Anh (English)", Tag = "en" });
            _cboTargetLang.Items.Add(new ComboBoxItem { Content = "🇨🇳 Tiếng Trung (Chinese)", Tag = "zh-CN" });
            _cboTargetLang.Items.Add(new ComboBoxItem { Content = "🇯🇵 Tiếng Nhật (Japanese)", Tag = "ja" });
            _cboTargetLang.Items.Add(new ComboBoxItem { Content = "🇰🇷 Tiếng Hàn (Korean)", Tag = "ko" });
            _cboTargetLang.Items.Add(new ComboBoxItem { Content = "🇫🇷 Tiếng Pháp (French)", Tag = "fr" });
            trnStack.Children.Add(_cboTargetLang);

            _chkReplaceSelection = new CheckBox
            {
                Content = "Thay thế trực tiếp đoạn bôi đen bằng bản dịch",
                IsChecked = false,
                Margin = new Thickness(0, 2, 0, 8),
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5
            };
            trnStack.Children.Add(_chkReplaceSelection);

            Button btnTranslateSel = ThemeColors.CreatePrimaryButton(
                "🌐 Dịch Đoạn Văn Đang Chọn (Google Dịch)",
                "Dịch nhanh đoạn văn bản bôi đen trong Word");
            btnTranslateSel.Margin = new Thickness(0, 0, 0, 6);
            btnTranslateSel.Click += (s, e) =>
            {
                string msg;
                string tLang = GetSelectedLangTag();
                bool replace = _chkReplaceSelection.IsChecked == true;
                _wordService.TranslateSelectedText("auto", tLang, replace, out msg);
                _notifyStatus(msg);
            };
            trnStack.Children.Add(btnTranslateSel);

            Button btnTranslateDoc = ThemeColors.CreateSecondaryButton(
                "📄 Dịch Toàn Bộ Tài Liệu Sang File Mới",
                "Dịch toàn bộ văn bản và mở trong một tài liệu mới");
            btnTranslateDoc.Margin = new Thickness(0, 0, 0, 4);
            btnTranslateDoc.Click += (s, e) =>
            {
                string msg;
                string tLang = GetSelectedLangTag();
                _wordService.TranslateEntireDocument("auto", tLang, out msg);
                _notifyStatus(msg);
            };
            trnStack.Children.Add(btnTranslateDoc);

            transCard.Child = trnStack;
            root.Children.Add(transCard);

            // 2. Card: Nghiệp vụ Kế toán & Hợp đồng: Bảng Song Ngữ
            Border biCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 4));
            StackPanel biStack = new StackPanel();

            biStack.Children.Add(ThemeColors.CreateSectionHeader("Thuật Ngữ Kế Toán & Hợp Đồng"));
            biStack.Children.Add(new TextBlock
            {
                Text = "Từ điển thuật ngữ chuyên ngành chuẩn mực (VAS, IFRS, Incoterms, Hợp đồng). Tạo bảng biểu song ngữ Anh - Việt chuyên nghiệp:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            Button btnBilingual = ThemeColors.CreatePrimaryButton(
                "💼 Chèn Bảng Thuật Ngữ Kế Toán Song Ngữ (Anh - Việt)",
                "Chèn bảng đối chiếu Báo cáo tài chính & Hợp đồng song ngữ");
            btnBilingual.Margin = new Thickness(0, 0, 0, 4);
            btnBilingual.Click += (s, e) =>
            {
                string msg;
                _wordService.TranslateFinancialAndLegalTerms("en", true, out msg);
                _notifyStatus(msg);
            };
            biStack.Children.Add(btnBilingual);

            biCard.Child = biStack;
            root.Children.Add(biCard);

            Content = root;
        }

        private string GetSelectedLangTag()
        {
            if (_cboTargetLang != null)
            {
                ComboBoxItem item = _cboTargetLang.SelectedItem as ComboBoxItem;
                if (item != null && item.Tag != null)
                {
                    return item.Tag.ToString();
                }
            }
            return "vi";
        }
    }
}
