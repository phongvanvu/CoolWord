using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using CoolWord.AddIn;
using CoolWord.Core;
using CoolWord.UI;

namespace CoolWord.Studio
{
    public static class Program
    {
        [STAThread]
        public static int Main(string[] args)
        {
            if (args != null && args.Length > 0)
            {
                string cmd = args[0].Trim().ToLowerInvariant();
                if (cmd == "--install" || cmd == "/install")
                {
                    string msg;
                    bool ok = AddInInstaller.InstallForCurrentUser(out msg);
                    Console.WriteLine(msg);
                    return ok ? 0 : 1;
                }
                if (cmd == "--uninstall" || cmd == "/uninstall")
                {
                    string msg;
                    bool ok = AddInInstaller.UninstallForCurrentUser(out msg);
                    Console.WriteLine(msg);
                    return ok ? 0 : 1;
                }
            }

            Application app = new Application();
            Window win = CreateStudioWindow();
            app.Run(win);
            return 0;
        }

        private static Window CreateStudioWindow()
        {
            Window win = new Window
            {
                Title = "CoolWord Studio - Quản Lý Add-in & Đổi Số Ra Chữ Tiếng Việt",
                Width = 460,
                Height = 760,
                MinWidth = 380,
                MinHeight = 560,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                Background = ThemeColors.AppBackground
            };

            Grid rootGrid = new Grid();
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            WordDocumentService wordService = new WordDocumentService();
            MainTaskPaneControl mainControl = new MainTaskPaneControl(wordService);

            // Thanh quản lý cài đặt Add-in vào MS Word ở trên cùng của cửa sổ Studio
            Border installBar = new Border
            {
                Background = Brushes.White,
                BorderBrush = ThemeColors.BorderBrush,
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(12, 8, 12, 8)
            };

            Grid barGrid = new Grid();
            barGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            barGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            barGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            TextBlock statusBadge = new TextBlock
            {
                FontFamily = ThemeColors.MainFont,
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(statusBadge, 0);
            barGrid.Children.Add(statusBadge);

            Button btnInstallToggle = ThemeColors.CreatePrimaryButton("Cài vào MS Word");
            btnInstallToggle.Margin = new Thickness(8, 0, 6, 0);
            btnInstallToggle.Padding = new Thickness(10, 5, 10, 5);
            Grid.SetColumn(btnInstallToggle, 1);
            barGrid.Children.Add(btnInstallToggle);

            Button btnLaunchWord = ThemeColors.CreateSecondaryButton("Mở Word");
            btnLaunchWord.Padding = new Thickness(10, 5, 10, 5);
            btnLaunchWord.Click += (s, e) =>
            {
                try
                {
                    Process.Start("winword.exe");
                    mainControl.UpdateStatus("Đang khởi động Microsoft Word...");
                }
                catch (Exception ex)
                {
                    mainControl.UpdateStatus("Không thể khởi động Word: " + ex.Message);
                }
            };
            Grid.SetColumn(btnLaunchWord, 2);
            barGrid.Children.Add(btnLaunchWord);

            Action refreshInstallState = () =>
            {
                bool installed = AddInInstaller.IsInstalled();
                if (installed)
                {
                    statusBadge.Text = "✔ Add-in: Đã kích hoạt trong Word";
                    statusBadge.Foreground = ThemeColors.SuccessText;
                    btnInstallToggle.Content = "Gỡ khỏi Word";
                }
                else
                {
                    statusBadge.Text = "○ Add-in: Chưa đăng ký vào Word";
                    statusBadge.Foreground = ThemeColors.TextSecondary;
                    btnInstallToggle.Content = "Cài vào Word";
                }
            };

            btnInstallToggle.Click += (s, e) =>
            {
                string msg;
                if (AddInInstaller.IsInstalled())
                {
                    AddInInstaller.UninstallForCurrentUser(out msg);
                }
                else
                {
                    AddInInstaller.InstallForCurrentUser(out msg);
                }
                refreshInstallState();
                mainControl.UpdateStatus(msg);
            };

            refreshInstallState();

            installBar.Child = barGrid;
            Grid.SetRow(installBar, 0);
            rootGrid.Children.Add(installBar);

            Grid.SetRow(mainControl, 1);
            rootGrid.Children.Add(mainControl);

            win.Content = rootGrid;
            return win;
        }
    }
}
