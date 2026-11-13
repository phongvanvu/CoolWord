using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using CoolWord.Modules.NumberToText;
using CoolWord.Modules.QuickInput;

namespace CoolWord.Core
{
    [DataContract]
    public class ConversionHistoryItem
    {
        [DataMember]
        public string InputNumber { get; set; }

        [DataMember]
        public string OutputWords { get; set; }

        [DataMember]
        public string Timestamp { get; set; }
    }

    [DataContract]
    public class AppSettings
    {
        [DataMember]
        public int ReadingMode { get; set; }

        [DataMember]
        public int Dialect { get; set; }

        [DataMember]
        public int Capitalization { get; set; }

        [DataMember]
        public int InsertMode { get; set; }

        [DataMember]
        public bool AppendChan { get; set; }

        [DataMember]
        public bool AppendDotSlash { get; set; }

        [DataMember]
        public bool ItalicizeConvertedText { get; set; }

        [DataMember]
        public bool AutoDetectCurrencyFromSelection { get; set; }

        [DataMember]
        public string CustomMajorUnit { get; set; }

        [DataMember]
        public string CustomMinorUnit { get; set; }

        [DataMember]
        public bool LiveSyncSelectionToTaskPane { get; set; }

        [DataMember]
        public List<ConversionHistoryItem> History { get; set; }

        [DataMember]
        public List<QuickSnippet> CustomSnippets { get; set; }

        public AppSettings()
        {
            ReadingMode = (int)NumberReadingMode.CurrencyVnd;
            Dialect = (int)VietnameseDialect.MixedCommon;
            Capitalization = (int)TextCapitalization.SentenceCase;
            InsertMode = (int)WordInsertMode.AppendInParentheses;
            AppendChan = true;
            AppendDotSlash = false;
            ItalicizeConvertedText = true;
            AutoDetectCurrencyFromSelection = true;
            CustomMajorUnit = "đồng";
            CustomMinorUnit = "xu";
            LiveSyncSelectionToTaskPane = true;
            History = new List<ConversionHistoryItem>();
            CustomSnippets = new List<QuickSnippet>();
        }

        public NumberConvertOptions ToNumberConvertOptions()
        {
            return new NumberConvertOptions
            {
                ReadingMode = (NumberReadingMode)ReadingMode,
                Dialect = (VietnameseDialect)Dialect,
                Capitalization = (TextCapitalization)Capitalization,
                InsertMode = (WordInsertMode)InsertMode,
                AppendChan = AppendChan,
                AppendDotSlash = AppendDotSlash,
                ItalicizeConvertedText = ItalicizeConvertedText,
                AutoDetectCurrencyFromSelection = AutoDetectCurrencyFromSelection,
                CustomMajorUnit = CustomMajorUnit ?? "đồng",
                CustomMinorUnit = CustomMinorUnit ?? "xu"
            };
        }

        public void ApplyFromOptions(NumberConvertOptions options)
        {
            if (options == null) return;
            ReadingMode = (int)options.ReadingMode;
            Dialect = (int)options.Dialect;
            Capitalization = (int)options.Capitalization;
            InsertMode = (int)options.InsertMode;
            AppendChan = options.AppendChan;
            AppendDotSlash = options.AppendDotSlash;
            ItalicizeConvertedText = options.ItalicizeConvertedText;
            AutoDetectCurrencyFromSelection = options.AutoDetectCurrencyFromSelection;
            CustomMajorUnit = options.CustomMajorUnit;
            CustomMinorUnit = options.CustomMinorUnit;
        }

        public void AddHistory(string input, string output)
        {
            if (string.IsNullOrWhiteSpace(input) || string.IsNullOrWhiteSpace(output))
                return;

            if (History == null)
                History = new List<ConversionHistoryItem>();

            History.RemoveAll(h => string.Equals(h.InputNumber, input.Trim(), StringComparison.OrdinalIgnoreCase));
            History.Insert(0, new ConversionHistoryItem
            {
                InputNumber = input.Trim(),
                OutputWords = output,
                Timestamp = DateTime.Now.ToString("HH:mm dd/MM")
            });

            if (History.Count > 15)
            {
                History.RemoveRange(15, History.Count - 15);
            }
        }
    }
}
