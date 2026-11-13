using System;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

namespace CoolWord.Core
{
    public static class SettingsService
    {
        private static readonly object SyncLock = new object();
        private static AppSettings _current;

        public static event EventHandler SettingsChanged;

        public static string SettingsFilePath
        {
            get
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string dir = Path.Combine(appData, "CoolWord");
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                return Path.Combine(dir, "settings.json");
            }
        }

        public static AppSettings Current
        {
            get
            {
                if (_current == null)
                {
                    lock (SyncLock)
                    {
                        if (_current == null)
                        {
                            _current = Load();
                        }
                    }
                }
                return _current;
            }
        }

        public static AppSettings Load()
        {
            try
            {
                string path = SettingsFilePath;
                if (File.Exists(path))
                {
                    byte[] bytes = File.ReadAllBytes(path);
                    using (MemoryStream ms = new MemoryStream(bytes))
                    {
                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(AppSettings));
                        AppSettings loaded = serializer.ReadObject(ms) as AppSettings;
                        if (loaded != null)
                        {
                            if (loaded.History == null)
                            {
                                loaded.History = new System.Collections.Generic.List<ConversionHistoryItem>();
                            }
                            return loaded;
                        }
                    }
                }
            }
            catch
            {
                // Khôi phục mặc định nếu file cấu hình lỗi
            }
            return new AppSettings();
        }

        public static void Save(AppSettings settings = null)
        {
            try
            {
                lock (SyncLock)
                {
                    if (settings != null)
                    {
                        _current = settings;
                    }
                    if (_current == null)
                    {
                        _current = new AppSettings();
                    }

                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(AppSettings));
                    using (MemoryStream ms = new MemoryStream())
                    {
                        serializer.WriteObject(ms, _current);
                        string json = Encoding.UTF8.GetString(ms.ToArray());
                        File.WriteAllText(SettingsFilePath, json, Encoding.UTF8);
                    }
                }

                EventHandler handler = SettingsChanged;
                if (handler != null)
                {
                    handler(null, EventArgs.Empty);
                }
            }
            catch
            {
                // Bỏ qua lỗi ghi file nếu bị khóa tạm thời
            }
        }
    }
}
