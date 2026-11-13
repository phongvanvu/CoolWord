using System;
using System.IO;
using System.Reflection;
using Microsoft.Win32;

namespace CoolWord.AddIn
{
    public static class AddInInstaller
    {
        public const string AddInProgId = "CoolWord.AddIn";
        public const string AddInClsid = "{4D7E9A11-8B2C-4F6D-9E3A-1C2B3D4E5F6A}";
        public const string AddInClassName = "CoolWord.AddIn.CoolWordAddIn";

        public const string TaskPaneProgId = "CoolWord.TaskPaneHost";
        public const string TaskPaneClsid = "{7C5E8A12-3D4F-4E9B-8A1C-2F3E4D5C6B7A}";
        public const string TaskPaneClassName = "CoolWord.AddIn.TaskPaneHost";

        private const string WordAddInRegPath = @"Software\Microsoft\Office\Word\Addins\" + AddInProgId;
        private const string DotNetCategoryGuid = "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}";

        public static bool IsInstalled()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(WordAddInRegPath, false))
                {
                    if (key == null) return false;
                    object lb = key.GetValue("LoadBehavior");
                    return lb != null && Convert.ToInt32(lb) == 3;
                }
            }
            catch
            {
                return false;
            }
        }

        public static bool InstallForCurrentUser(out string message)
        {
            try
            {
                Assembly asm = typeof(CoolWordAddIn).Assembly;
                string dllPath = asm.Location;
                if (string.IsNullOrEmpty(dllPath) || !File.Exists(dllPath))
                {
                    message = "Không tìm thấy tệp CoolWord.dll tại đường dẫn thực thi.";
                    return false;
                }

                string codeBase = new Uri(dllPath).AbsoluteUri;
                string asmFullName = asm.FullName;
                string asmVersion = asm.GetName().Version.ToString();

                // Đăng ký COM Class cho cả kiến trúc 64-bit và 32-bit (Wow6432Node) trong HKCU\Software\Classes
                string[] classRoots = new[]
                {
                    @"Software\Classes",
                    @"Software\Classes\Wow6432Node"
                };

                foreach (string rootPath in classRoots)
                {
                    RegisterComClass(rootPath, AddInProgId, AddInClsid, AddInClassName, asmFullName, asmVersion, codeBase, false);
                    RegisterComClass(rootPath, TaskPaneProgId, TaskPaneClsid, TaskPaneClassName, asmFullName, asmVersion, codeBase, true);
                }

                // Đăng ký Add-in với Microsoft Word
                using (RegistryKey wordKey = Registry.CurrentUser.CreateSubKey(WordAddInRegPath))
                {
                    if (wordKey != null)
                    {
                        wordKey.SetValue("FriendlyName", "CoolWord - Bộ Công Cụ Tiếng Việt", RegistryValueKind.String);
                        wordKey.SetValue("Description", "Đổi số ra chữ Tiếng Việt & Chuẩn hóa văn bản cho MS Word 2016+", RegistryValueKind.String);
                        wordKey.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                    }
                }

                // Xóa trạng thái bị chặn (Resiliency / DisabledItems) nếu trước đó Word từng tắt Add-in
                message = "Đã đăng ký thành công Add-in CoolWord vào Microsoft Word (" + dllPath + "). Hãy khởi động lại MS Word để sử dụng!";
                return true;
            }
            catch (Exception ex)
            {
                message = "Lỗi khi đăng ký Add-in: " + ex.Message;
                return false;
            }
        }

        public static bool UninstallForCurrentUser(out string message)
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(WordAddInRegPath, false);

                string[] classRoots = new[]
                {
                    @"Software\Classes",
                    @"Software\Classes\Wow6432Node"
                };

                foreach (string rootPath in classRoots)
                {
                    Registry.CurrentUser.DeleteSubKeyTree(rootPath + @"\" + AddInProgId, false);
                    Registry.CurrentUser.DeleteSubKeyTree(rootPath + @"\CLSID\" + AddInClsid, false);

                    Registry.CurrentUser.DeleteSubKeyTree(rootPath + @"\" + TaskPaneProgId, false);
                    Registry.CurrentUser.DeleteSubKeyTree(rootPath + @"\CLSID\" + TaskPaneClsid, false);
                }

                message = "Đã gỡ cài đặt Add-in CoolWord khỏi Microsoft Word.";
                return true;
            }
            catch (Exception ex)
            {
                message = "Lỗi khi gỡ cài đặt Add-in: " + ex.Message;
                return false;
            }
        }

        private static void RegisterComClass(
            string classesRootPath,
            string progId,
            string clsid,
            string className,
            string asmFullName,
            string asmVersion,
            string codeBase,
            bool isActiveXControl)
        {
            using (RegistryKey progIdKey = Registry.CurrentUser.CreateSubKey(classesRootPath + @"\" + progId))
            {
                if (progIdKey != null)
                {
                    progIdKey.SetValue("", className, RegistryValueKind.String);
                    using (RegistryKey clsidSub = progIdKey.CreateSubKey("CLSID"))
                    {
                        if (clsidSub != null) clsidSub.SetValue("", clsid, RegistryValueKind.String);
                    }
                }
            }

            string clsidKeyPath = classesRootPath + @"\CLSID\" + clsid;
            using (RegistryKey clsidKey = Registry.CurrentUser.CreateSubKey(clsidKeyPath))
            {
                if (clsidKey == null) return;
                clsidKey.SetValue("", className, RegistryValueKind.String);

                using (RegistryKey inproc = clsidKey.CreateSubKey("InprocServer32"))
                {
                    if (inproc != null)
                    {
                        inproc.SetValue("", "mscoree.dll", RegistryValueKind.String);
                        inproc.SetValue("ThreadingModel", "Both", RegistryValueKind.String);
                        inproc.SetValue("Class", className, RegistryValueKind.String);
                        inproc.SetValue("Assembly", asmFullName, RegistryValueKind.String);
                        inproc.SetValue("RuntimeVersion", "v4.0.30319", RegistryValueKind.String);
                        inproc.SetValue("CodeBase", codeBase, RegistryValueKind.String);

                        using (RegistryKey verKey = inproc.CreateSubKey(asmVersion))
                        {
                            if (verKey != null)
                            {
                                verKey.SetValue("Class", className, RegistryValueKind.String);
                                verKey.SetValue("Assembly", asmFullName, RegistryValueKind.String);
                                verKey.SetValue("RuntimeVersion", "v4.0.30319", RegistryValueKind.String);
                                verKey.SetValue("CodeBase", codeBase, RegistryValueKind.String);
                            }
                        }
                    }
                }

                using (RegistryKey progIdSub = clsidKey.CreateSubKey("ProgId"))
                {
                    if (progIdSub != null) progIdSub.SetValue("", progId, RegistryValueKind.String);
                }

                using (RegistryKey catKey = clsidKey.CreateSubKey(@"Implemented Categories\" + DotNetCategoryGuid))
                {
                }

                if (isActiveXControl)
                {
                    using (clsidKey.CreateSubKey("Control")) { }
                    using (clsidKey.CreateSubKey("Programmable")) { }
                }
            }
        }
    }
}
