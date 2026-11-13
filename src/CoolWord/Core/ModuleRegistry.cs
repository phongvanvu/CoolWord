using System;
using System.Collections.Generic;
using System.Linq;
using CoolWord.Modules.AdminFormat;
using CoolWord.Modules.NameSplit;
using CoolWord.Modules.NumberToText;
using CoolWord.Modules.Pdf;
using CoolWord.Modules.Printing;
using CoolWord.Modules.QuickInput;
using CoolWord.Modules.TextUtilities;
using CoolWord.Modules.WordToWord;
using CoolWord.Modules.WordToExcel;
using CoolWord.Modules.WordToPowerPoint;
using CoolWord.Modules.BatchEmail;
using CoolWord.Modules.ImageTools;
using CoolWord.Modules.GoogleTranslate;

namespace CoolWord.Core
{
    public static class ModuleRegistry
    {
        private static readonly List<IToolModule> RegisteredModules = new List<IToolModule>();
        private static bool _initialized;

        public static void EnsureInitialized()
        {
            if (_initialized) return;
            _initialized = true;

            Register(new NumberToTextModule());
            Register(new AdminFormatModule());
            Register(new TextCleanModule());
            Register(new PrintModule());
            Register(new PdfModule());
            Register(new QuickInputModule());
            Register(new NameSplitModule());
            Register(new WordToWordModule());
            Register(new WordToExcelModule());
            Register(new WordToPowerPointModule());
            Register(new BatchEmailModule());
            Register(new ImageToolsModule());
            Register(new GoogleTranslateModule());
        }

        public static void Register(IToolModule module)
        {
            if (module == null) return;
            RegisteredModules.RemoveAll(m => string.Equals(m.Id, module.Id, StringComparison.OrdinalIgnoreCase));
            RegisteredModules.Add(module);
        }

        public static IReadOnlyList<IToolModule> GetAllModules()
        {
            EnsureInitialized();
            return RegisteredModules.OrderBy(m => m.Order).ToList();
        }

        public static IToolModule GetById(string id)
        {
            EnsureInitialized();
            return RegisteredModules.FirstOrDefault(m => string.Equals(m.Id, id, StringComparison.OrdinalIgnoreCase));
        }
    }
}
