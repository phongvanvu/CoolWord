using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using Word = Microsoft.Office.Interop.Word;

namespace CoolWord.Modules.GoogleTranslate
{
    /// <summary>
    /// Dịch vụ dịch thuật văn bản qua Google Dịch (Google Translate) kết hợp từ điển chuyên ngành Kế toán - Tài chính - Pháp lý.
    /// Hoạt động cả Online (Google Translate API) lẫn Offline (Từ điển song ngữ chuyên ngành).
    /// </summary>
    public static class GoogleTranslateService
    {
        private static readonly object Missing = Type.Missing;

        #region Bilingual Accounting & Financial Dictionary (Offline)

        private static readonly Dictionary<string, string> ViToEnDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // Báo cáo tài chính & Kế toán
            { "Bảng cân đối kế toán", "Balance Sheet" },
            { "Báo cáo kết quả hoạt động kinh doanh", "Income Statement / Profit & Loss" },
            { "Báo cáo lưu chuyển tiền tệ", "Cash Flow Statement" },
            { "Thuyết minh báo cáo tài chính", "Notes to Financial Statements" },
            { "Tài sản ngắn hạn", "Current Assets" },
            { "Tài sản dài hạn", "Non-current Assets" },
            { "Tiền và các khoản tương đương tiền", "Cash and Cash Equivalents" },
            { "Các khoản đầu tư tài chính ngắn hạn", "Short-term Financial Investments" },
            { "Phải thu ngắn hạn của khách hàng", "Short-term Accounts Receivable" },
            { "Hàng tồn kho", "Inventories" },
            { "Tài sản cố định hữu hình", "Tangible Fixed Assets" },
            { "Hao mòn lũy kế", "Accumulated Depreciation" },
            { "Nợ phải trả", "Liabilities" },
            { "Nợ ngắn hạn", "Current Liabilities" },
            { "Phải trả người bán ngắn hạn", "Short-term Accounts Payable" },
            { "Thuế và các khoản phải nộp Nhà nước", "Taxes and Other Payables to the State" },
            { "Vốn chủ sở hữu", "Owner's Equity" },
            { "Vốn góp của chủ sở hữu", "Paid-in Capital / Contributed Capital" },
            { "Lợi nhuận sau thuế chưa phân phối", "Undistributed Earnings after Tax" },
            { "Doanh thu bán hàng và cung cấp dịch vụ", "Revenue from Sales and Services" },
            { "Các khoản giảm trừ doanh thu", "Revenue Deductions" },
            { "Doanh thu thuần", "Net Revenue" },
            { "Giá vốn hàng bán", "Cost of Goods Sold (COGS)" },
            { "Lợi nhuận gộp", "Gross Profit" },
            { "Doanh thu hoạt động tài chính", "Financial Income" },
            { "Chi phí tài chính", "Financial Expenses" },
            { "Chi phí lãi vay", "Interest Expense" },
            { "Chi phí bán hàng", "Selling Expenses" },
            { "Chi phí quản lý doanh nghiệp", "General & Administrative Expenses" },
            { "Lợi nhuận thuần từ hoạt động kinh doanh", "Operating Profit" },
            { "Thu nhập khác", "Other Income" },
            { "Chi phí khác", "Other Expenses" },
            { "Tổng lợi nhuận kế toán trước thuế", "Accounting Profit before Tax" },
            { "Chi phí thuế thu nhập doanh nghiệp", "Corporate Income Tax Expense" },
            { "Lợi nhuận sau thuế thu nhập doanh nghiệp", "Net Profit after Tax" },
            { "Lãi cơ bản trên cổ phiếu", "Basic Earnings per Share (EPS)" },

            // Hợp đồng & Pháp lý
            { "Cộng hòa xã hội chủ nghĩa Việt Nam", "Socialist Republic of Vietnam" },
            { "Độc lập - Tự do - Hạnh phúc", "Independence - Freedom - Happiness" },
            { "Hợp đồng kinh tế", "Economic Contract" },
            { "Hợp đồng mua bán hàng hóa", "Sale and Purchase Agreement" },
            { "Hợp đồng cung cấp dịch vụ", "Service Agreement" },
            { "Hợp đồng lao động", "Employment Contract" },
            { "Bên A", "Party A" },
            { "Bên B", "Party B" },
            { "Người đại diện theo pháp luật", "Legal Representative" },
            { "Điều khoản thanh toán", "Payment Terms" },
            { "Phương thức thanh toán", "Payment Method" },
            { "Tạm ứng", "Advance Payment" },
            { "Bảo lãnh thực hiện hợp đồng", "Performance Guarantee" },
            { "Sự kiện bất khả kháng", "Force Majeure" },
            { "Bảo mật thông tin", "Confidentiality / Non-Disclosure" },
            { "Phạt vi phạm hợp đồng", "Penalty for Breach of Contract" },
            { "Bồi thường thiệt hại", "Indemnification / Damages" },
            { "Giải quyết tranh chấp", "Dispute Resolution" },
            { "Trọng tài thương mại", "Commercial Arbitration" },
            { "Hiệu lực thi hành", "Validity and Execution" }
        };

        private static readonly Dictionary<string, string> EnToViDict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        static GoogleTranslateService()
        {
            foreach (var kvp in ViToEnDict)
            {
                if (!EnToViDict.ContainsKey(kvp.Value))
                {
                    EnToViDict[kvp.Value] = kvp.Key;
                }
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Dịch đoạn văn bản đang được chọn trong tài liệu Word.
        /// </summary>
        public static bool TranslateSelectedText(
            Word.Application wordApp,
            string sourceLang,
            string targetLang,
            bool replaceSelection,
            out string statusMessage)
        {
            statusMessage = string.Empty;
            if (wordApp == null || wordApp.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                Word.Selection sel = wordApp.Selection;
                if (sel == null || string.IsNullOrWhiteSpace(sel.Text))
                {
                    statusMessage = "Vui lòng bôi đen đoạn văn bản cần dịch.";
                    return false;
                }

                string raw = sel.Text.Trim();
                string sLang = string.IsNullOrEmpty(sourceLang) ? "auto" : sourceLang;
                string tLang = string.IsNullOrEmpty(targetLang) ? "vi" : targetLang;

                string translated = TranslateText(raw, sLang, tLang);

                if (string.IsNullOrWhiteSpace(translated))
                {
                    statusMessage = "Không thể dịch đoạn văn bản.";
                    return false;
                }

                if (replaceSelection)
                {
                    sel.Text = translated;
                }
                else
                {
                    Word.Range endRange = sel.Range.Duplicate;
                    object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                    endRange.Collapse(ref collapseEnd);
                    endRange.Text = "\n[Bản dịch]: " + translated;
                }

                statusMessage = "Đã dịch văn bản thành công qua Google Dịch.";
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi dịch văn bản: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Dịch toàn bộ nội dung tài liệu Word sang ngôn ngữ đích và mở trong tài liệu mới.
        /// </summary>
        public static bool TranslateEntireDocument(
            Word.Application wordApp,
            string sourceLang,
            string targetLang,
            out string statusMessage)
        {
            statusMessage = string.Empty;
            if (wordApp == null || wordApp.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                Word.Document doc = wordApp.ActiveDocument;
                string text = doc.Content.Text;
                if (string.IsNullOrWhiteSpace(text))
                {
                    statusMessage = "Tài liệu trống, không có nội dung.";
                    return false;
                }

                string sLang = string.IsNullOrEmpty(sourceLang) ? "auto" : sourceLang;
                string tLang = string.IsNullOrEmpty(targetLang) ? "vi" : targetLang;

                // Cắt theo từng đoạn văn bản để dịch
                string[] paragraphs = text.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                List<string> translatedParas = new List<string>();

                foreach (string p in paragraphs)
                {
                    string trimmed = p.Trim();
                    if (!string.IsNullOrWhiteSpace(trimmed))
                    {
                        translatedParas.Add(TranslateText(trimmed, sLang, tLang));
                    }
                }

                Word.Document newDoc = wordApp.Documents.Add();
                Word.Range r = newDoc.Content;
                foreach (string tp in translatedParas)
                {
                    r.InsertAfter(tp + "\n");
                }

                statusMessage = string.Format("Đã dịch toàn bộ tài liệu ({0} đoạn) sang ngôn ngữ mới.", translatedParas.Count);
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi dịch tài liệu: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Nghiệp Vụ Kế Toán & Hợp Đồng: Tạo Bảng Song Ngữ (Anh - Việt) hoặc dịch thuật ngữ tài chính chuẩn mực.
        /// </summary>
        public static bool TranslateFinancialAndLegalTerms(
            Word.Application wordApp,
            string targetLang,
            bool createBilingualTable,
            out string statusMessage)
        {
            statusMessage = string.Empty;
            if (wordApp == null || wordApp.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                Word.Document doc = wordApp.ActiveDocument;
                Word.Selection sel = wordApp.Selection;

                if (createBilingualTable)
                {
                    // Tạo một bảng song ngữ mẫu cho Báo cáo tài chính / Hợp đồng
                    Word.Range r = (sel != null && sel.Start != sel.End) ? sel.Range : doc.Content;
                    object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                    r.Collapse(ref collapseEnd);

                    List<KeyValuePair<string, string>> sampleTerms = new List<KeyValuePair<string, string>>
                    {
                        new KeyValuePair<string, string>("TIẾNG VIỆT", "TIẾNG ANH (ENGLISH)"),
                        new KeyValuePair<string, string>("Bảng cân đối kế toán", "Balance Sheet"),
                        new KeyValuePair<string, string>("Doanh thu thuần", "Net Revenue"),
                        new KeyValuePair<string, string>("Giá vốn hàng bán (COGS)", "Cost of Goods Sold"),
                        new KeyValuePair<string, string>("Lợi nhuận gộp", "Gross Profit"),
                        new KeyValuePair<string, string>("Chi phí quản lý doanh nghiệp", "General & Administrative Expenses"),
                        new KeyValuePair<string, string>("Tổng lợi nhuận trước thuế", "Accounting Profit before Tax"),
                        new KeyValuePair<string, string>("Lợi nhuận sau thuế (Lãi ròng)", "Net Profit after Tax")
                    };

                    Word.Table table = doc.Tables.Add(r, sampleTerms.Count, 2);
                    table.Borders.Enable = 1;

                    for (int i = 0; i < sampleTerms.Count; i++)
                    {
                        table.Cell(i + 1, 1).Range.Text = sampleTerms[i].Key;
                        table.Cell(i + 1, 2).Range.Text = sampleTerms[i].Value;

                        if (i == 0)
                        {
                            table.Rows[1].Range.Font.Bold = 1;
                            table.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            table.Rows[1].Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;
                        }
                    }

                    statusMessage = "Đã chèn Bảng Thuật Ngữ Kế Toán Song Ngữ (Anh - Việt) chuẩn mực vào tài liệu.";
                    return true;
                }
                else
                {
                    // Dịch đoạn bôi đen dựa trên từ điển chuyên ngành
                    return TranslateSelectedText(wordApp, "auto", targetLang, false, out statusMessage);
                }
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi xử lý dịch thuật kế toán: " + ex.Message;
                return false;
            }
        }

        #endregion

        #region Translation Engine (Online + Offline Fallback)

        /// <summary>
        /// Dịch chuỗi văn bản: Ưu tiên tra từ điển chuyên ngành trước, nếu không có thì gọi Google Translate Online.
        /// </summary>
        public static string TranslateText(string input, string sourceLang, string targetLang)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;

            string trimmed = input.Trim();

            // 1. Kiểm tra từ điển chuyên ngành trước
            if (targetLang.Equals("en", StringComparison.OrdinalIgnoreCase))
            {
                string matchEn;
                if (ViToEnDict.TryGetValue(trimmed, out matchEn))
                {
                    return matchEn;
                }
            }
            else if (targetLang.Equals("vi", StringComparison.OrdinalIgnoreCase))
            {
                string matchVi;
                if (EnToViDict.TryGetValue(trimmed, out matchVi))
                {
                    return matchVi;
                }
            }

            // 2. Gọi Google Translate API miễn phí (Endpoint gtx)
            try
            {
                string sLang = string.IsNullOrEmpty(sourceLang) ? "auto" : sourceLang;
                string tLang = string.IsNullOrEmpty(targetLang) ? "vi" : targetLang;

                string url = string.Format(
                    "https://translate.googleapis.com/translate_a/single?client=gtx&sl={0}&tl={1}&dt=t&q={2}",
                    sLang,
                    tLang,
                    Uri.EscapeDataString(trimmed));

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "GET";
                request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) CoolWord/1.0";
                request.Timeout = 4000; // 4 giây timeout

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    string json = reader.ReadToEnd();
                    return ParseGoogleTranslateJson(json);
                }
            }
            catch
            {
                // Fallback nếu mất mạng hoặc Google API bị giới hạn:
                // Thay thế các từ khóa xuất hiện trong từ điển
                string fallback = trimmed;
                if (targetLang.Equals("en", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (var kvp in ViToEnDict)
                    {
                        if (fallback.IndexOf(kvp.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            fallback = Regex.Replace(fallback, Regex.Escape(kvp.Key), kvp.Value, RegexOptions.IgnoreCase);
                        }
                    }
                }
                else
                {
                    foreach (var kvp in EnToViDict)
                    {
                        if (fallback.IndexOf(kvp.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            fallback = Regex.Replace(fallback, Regex.Escape(kvp.Key), kvp.Value, RegexOptions.IgnoreCase);
                        }
                    }
                }
                return fallback;
            }
        }

        /// <summary>
        /// Phân tích JSON trả về từ Google Translate gtx endpoint:
        /// [[["Bản dịch","Văn bản gốc",null,null,1],...],...]
        /// </summary>
        public static string ParseGoogleTranslateJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return string.Empty;

            StringBuilder sb = new StringBuilder();
            MatchCollection matches = Regex.Matches(json, @"\[\""(?<trans>.*?)\"",\""(?<orig>.*?)\""");

            foreach (Match m in matches)
            {
                if (m.Success)
                {
                    string part = m.Groups["trans"].Value;
                    part = Regex.Unescape(part);
                    sb.Append(part);
                }
            }

            string result = sb.ToString().Trim();
            if (string.IsNullOrEmpty(result))
            {
                // Fallback trích xuất chuỗi đầu tiên
                int start = json.IndexOf("\"");
                if (start >= 0)
                {
                    int end = json.IndexOf("\"", start + 1);
                    if (end > start)
                    {
                        result = json.Substring(start + 1, end - start - 1);
                    }
                }
            }

            return result;
        }

        #endregion
    }
}
