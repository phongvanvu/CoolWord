using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Modules.BatchEmail
{
    public class BatchEmailModuleView : System.Windows.Controls.UserControl
    {
        private readonly IWordDocumentService _wordService;
        private readonly Action<string> _notifyStatus;

        private System.Windows.Controls.TextBox _txtRecipient;
        private System.Windows.Controls.TextBox _txtSubject;
        private System.Windows.Controls.TextBox _txtBody;
        private System.Windows.Controls.CheckBox _chkAttachPdf;

        public BatchEmailModuleView(IWordDocumentService wordService, Action<string> notifyStatus)
        {
            _wordService = wordService;
            _notifyStatus = notifyStatus ?? (_ => { });

            BuildUi();
        }

        private void BuildUi()
        {
            StackPanel root = new StackPanel { Margin = new Thickness(0, 0, 0, 12) };

            // 1. Card: Gửi Tài Liệu Hiện Tại
            Border currentCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel curStack = new StackPanel();

            curStack.Children.Add(ThemeColors.CreateSectionHeader("Gửi Tài Liệu Word Qua Email"));
            curStack.Children.Add(new TextBlock
            {
                Text = "Gửi ngay tài liệu đang mở dưới dạng tệp đính kèm (.docx hoặc .pdf) qua Microsoft Outlook hoặc ứng dụng Mail mặc định:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            curStack.Children.Add(new TextBlock { Text = "Email người nhận:", FontSize = 11, FontFamily = ThemeColors.MainFont, Margin = new Thickness(0, 2, 0, 2) });
            _txtRecipient = new System.Windows.Controls.TextBox { FontFamily = ThemeColors.MainFont, FontSize = 12, Height = 26, Margin = new Thickness(0, 0, 0, 6) };
            curStack.Children.Add(_txtRecipient);

            curStack.Children.Add(new TextBlock { Text = "Tiêu đề thư:", FontSize = 11, FontFamily = ThemeColors.MainFont, Margin = new Thickness(0, 2, 0, 2) });
            _txtSubject = new System.Windows.Controls.TextBox { FontFamily = ThemeColors.MainFont, FontSize = 12, Height = 26, Text = "Kính gửi tài liệu Word từ CoolWord", Margin = new Thickness(0, 0, 0, 6) };
            curStack.Children.Add(_txtSubject);

            curStack.Children.Add(new TextBlock { Text = "Nội dung thư:", FontSize = 11, FontFamily = ThemeColors.MainFont, Margin = new Thickness(0, 2, 0, 2) });
            _txtBody = new System.Windows.Controls.TextBox { FontFamily = ThemeColors.MainFont, FontSize = 12, Height = 50, TextWrapping = TextWrapping.Wrap, AcceptsReturn = true, Text = "Kính gửi quý đối tác,\n\nĐính kèm tài liệu từ CoolWord.\nTrân trọng!", Margin = new Thickness(0, 0, 0, 6) };
            curStack.Children.Add(_txtBody);

            _chkAttachPdf = new System.Windows.Controls.CheckBox
            {
                Content = "Xuất đính kèm dạng file PDF (Không gửi file .docx gốc)",
                IsChecked = true,
                Margin = new Thickness(0, 2, 0, 8),
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5
            };
            curStack.Children.Add(_chkAttachPdf);

            System.Windows.Controls.Button btnSendCurrent = ThemeColors.CreatePrimaryButton(
                "📤 Soạn Thư & Gửi Tài Liệu Hiện Tại",
                "Mở cửa sổ soạn thư đính kèm tài liệu");
            btnSendCurrent.Margin = new Thickness(0, 0, 0, 4);
            btnSendCurrent.Click += (s, e) =>
            {
                string msg;
                bool asPdf = _chkAttachPdf.IsChecked == true;
                _wordService.SendCurrentDocumentViaEmail(_txtRecipient.Text, _txtSubject.Text, _txtBody.Text, asPdf, out msg);
                _notifyStatus(msg);
            };
            curStack.Children.Add(btnSendCurrent);

            currentCard.Child = curStack;
            root.Children.Add(currentCard);

            // 2. Card: Nghiệp vụ Kế toán: Gửi Phiếu Lương / Hóa Đơn Hàng Loạt
            Border voucherCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 10));
            StackPanel voucStack = new StackPanel();

            voucStack.Children.Add(ThemeColors.CreateSectionHeader("Nghiệp Vụ Kế Toán: Gửi Phiếu Lương / Chứng Từ"));
            voucStack.Children.Add(new TextBlock
            {
                Text = "Tự động trích xuất từng trang văn bản thành file PDF riêng biệt cho từng người và gửi kèm email riêng cho từng nhân sự / khách hàng:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            System.Windows.Controls.Button btnSendVouchers = ThemeColors.CreatePrimaryButton(
                "💼 Gửi Phiếu Lương / Chứng Từ Cho Từng Người...",
                "Chọn danh sách nhân sự (CSV/Excel) và tự động chia PDF gửi riêng");
            btnSendVouchers.Margin = new Thickness(0, 0, 0, 6);
            btnSendVouchers.Click += (s, e) =>
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Title = "Chọn file danh sách nhận chứng từ (CSV)";
                    ofd.Filter = "File danh sách CSV (*.csv)|*.csv|Tất cả tệp (*.*)|*.*";
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        string csvText = File.ReadAllText(ofd.FileName, System.Text.Encoding.UTF8);
                        var list = BatchEmailService.ParseCsvRecipients(csvText);
                        string msg;
                        _wordService.SendAccountingVouchersBatch(list, null, null, out msg);
                        _notifyStatus(msg);
                    }
                }
            };
            voucStack.Children.Add(btnSendVouchers);

            voucherCard.Child = voucStack;
            root.Children.Add(voucherCard);

            // 3. Card: Gửi Mail Trộn (Mail Merge)
            Border mergeCard = ThemeColors.CreateCard(new Thickness(12), new Thickness(0, 0, 0, 4));
            StackPanel mergeStack = new StackPanel();

            mergeStack.Children.Add(ThemeColors.CreateSectionHeader("Gửi Mail Trộn (Mail Merge)"));
            mergeStack.Children.Add(new TextBlock
            {
                Text = "Sử dụng các biến giữ chỗ {HoTen}, {Email}, {SoTien}, {PhongBan} để tự động cá nhân hóa nội dung thư gửi hàng loạt:",
                FontFamily = ThemeColors.MainFont,
                FontSize = 11.5,
                Foreground = ThemeColors.TextSecondary,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            System.Windows.Controls.Button btnBatchMerge = ThemeColors.CreateSecondaryButton(
                "📋 Tải Danh Sách Người Nhận (CSV) & Gửi Hàng Loạt...",
                "Chọn file danh sách và gửi thư hàng loạt qua Outlook");
            btnBatchMerge.Margin = new Thickness(0, 0, 0, 4);
            btnBatchMerge.Click += (s, e) =>
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Title = "Chọn file danh sách người nhận (CSV)";
                    ofd.Filter = "File danh sách CSV (*.csv)|*.csv|Tất cả tệp (*.*)|*.*";
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        string csvText = File.ReadAllText(ofd.FileName, System.Text.Encoding.UTF8);
                        var list = BatchEmailService.ParseCsvRecipients(csvText);
                        string msg;
                        _wordService.SendBatchMailMerge(list, "Thông báo: {HoTen}", "Kính gửi {HoTen},\n\nNội dung thông báo gửi bạn.\nTrân trọng!", true, out msg);
                        _notifyStatus(msg);
                    }
                }
            };
            mergeStack.Children.Add(btnBatchMerge);

            mergeCard.Child = mergeStack;
            root.Children.Add(mergeCard);

            Content = root;
        }
    }
}
