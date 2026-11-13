using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace CoolWord.Modules.WordToExcel
{
    /// <summary>
    /// Dịch vụ xuất dữ liệu từ Word sang Microsoft Excel.
    /// Hỗ trợ xuất bảng biểu, xuất bảng kế toán tự động chèn công thức SUM, xuất danh sách theo dòng.
    /// </summary>
    public static class WordToExcelService
    {
        private static readonly object Missing = Type.Missing;

        #region Public Methods

        /// <summary>
        /// Xuất tất cả các bảng biểu trong tài liệu Word sang file Excel (.xlsx).
        /// </summary>
        public static bool ExportAllTablesToExcel(Word.Application wordApp, string outputPath, bool openAfterExport, bool separateSheets, out string statusMessage)
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
                if (doc.Tables.Count == 0)
                {
                    statusMessage = "Tài liệu Word không có bảng biểu nào.";
                    return false;
                }

                List<List<List<string>>> allTablesData = new List<List<List<string>>>();
                for (int t = 1; t <= doc.Tables.Count; t++)
                {
                    Word.Table table = doc.Tables[t];
                    allTablesData.Add(ExtractTableGrid(table));
                }

                string targetPath = GetTargetPath(doc, outputPath, "_BangBieu.xlsx");
                bool ok = WriteTablesToExcelWorkbook(allTablesData, targetPath, separateSheets, false, openAfterExport, out statusMessage);
                if (ok && string.IsNullOrEmpty(statusMessage))
                {
                    statusMessage = string.Format("Đã xuất thành công {0} bảng biểu sang Excel: {1}", doc.Tables.Count, Path.GetFileName(targetPath));
                }
                return ok;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi xuất bảng sang Excel: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Xuất bảng biểu đang được con trỏ bôi đen / chọn trong Word sang Excel.
        /// </summary>
        public static bool ExportSelectedTableToExcel(Word.Application wordApp, string outputPath, bool openAfterExport, out string statusMessage)
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
                Word.Table targetTable = null;

                if (sel.Tables.Count > 0)
                {
                    targetTable = sel.Tables[1];
                }
                else
                {
                    Word.Document doc = wordApp.ActiveDocument;
                    if (doc.Tables.Count > 0)
                    {
                        targetTable = doc.Tables[1];
                    }
                }

                if (targetTable == null)
                {
                    statusMessage = "Vui lòng đặt con trỏ vào bên trong một bảng biểu cần xuất.";
                    return false;
                }

                List<List<string>> tableData = ExtractTableGrid(targetTable);
                List<List<List<string>>> tables = new List<List<List<string>>> { tableData };

                string targetPath = GetTargetPath(wordApp.ActiveDocument, outputPath, "_BangChon.xlsx");
                bool ok = WriteTablesToExcelWorkbook(tables, targetPath, false, false, openAfterExport, out statusMessage);
                if (ok && string.IsNullOrEmpty(statusMessage))
                {
                    statusMessage = "Đã xuất bảng đang chọn sang Excel: " + Path.GetFileName(targetPath);
                }
                return ok;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi xuất bảng sang Excel: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Tính năng Kế Toán: Xuất Bảng Kê / Báo Cáo Tài Chính sang Excel.
        /// Tự động nhận diện cột số tiền, áp dụng định dạng số kế toán (#,##0) và chèn hàng TỔNG CỘNG với công thức =SUM(...).
        /// </summary>
        public static bool ExportAccountingTableToExcel(Word.Application wordApp, string outputPath, bool openAfterExport, out string statusMessage)
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
                Word.Table targetTable = null;

                if (sel.Tables.Count > 0)
                {
                    targetTable = sel.Tables[1];
                }
                else
                {
                    Word.Document doc = wordApp.ActiveDocument;
                    if (doc.Tables.Count > 0)
                    {
                        targetTable = doc.Tables[1];
                    }
                }

                if (targetTable == null)
                {
                    statusMessage = "Không tìm thấy bảng biểu chứng từ/kế toán nào trong tài liệu.";
                    return false;
                }

                List<List<string>> tableData = ExtractTableGrid(targetTable);
                List<List<List<string>>> tables = new List<List<List<string>>> { tableData };

                string targetPath = GetTargetPath(wordApp.ActiveDocument, outputPath, "_KeToan_ChungTu.xlsx");
                bool ok = WriteTablesToExcelWorkbook(tables, targetPath, false, true, openAfterExport, out statusMessage);
                if (ok && string.IsNullOrEmpty(statusMessage))
                {
                    statusMessage = "Đã xuất Bảng Kê / Báo Cáo Kế Toán chuẩn sang Excel kèm công thức SUM: " + Path.GetFileName(targetPath);
                }
                return ok;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi xuất bảng kế toán sang Excel: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Xuất danh sách các đoạn văn bản (Paragraphs) hoặc danh mục theo dòng sang cột A của Excel.
        /// </summary>
        public static bool ExportParagraphsToExcel(Word.Application wordApp, string outputPath, bool openAfterExport, out string statusMessage)
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
                Word.Range range = (wordApp.Selection != null && wordApp.Selection.Start != wordApp.Selection.End)
                    ? wordApp.Selection.Range
                    : doc.Content;

                List<string> lines = new List<string>();
                foreach (Word.Paragraph p in range.Paragraphs)
                {
                    string txt = CleanCellText(p.Range.Text);
                    if (!string.IsNullOrWhiteSpace(txt))
                    {
                        lines.Add(txt);
                    }
                }

                if (lines.Count == 0)
                {
                    statusMessage = "Không có đoạn văn bản nào để xuất.";
                    return false;
                }

                List<List<string>> tableGrid = new List<List<string>>();
                tableGrid.Add(new List<string> { "STT", "Nội Dung Văn Bản" });
                for (int i = 0; i < lines.Count; i++)
                {
                    tableGrid.Add(new List<string> { (i + 1).ToString(), lines[i] });
                }

                List<List<List<string>>> tables = new List<List<List<string>>> { tableGrid };
                string targetPath = GetTargetPath(doc, outputPath, "_DanhSach.xlsx");
                bool ok = WriteTablesToExcelWorkbook(tables, targetPath, false, false, openAfterExport, out statusMessage);
                if (ok && string.IsNullOrEmpty(statusMessage))
                {
                    statusMessage = string.Format("Đã xuất {0} dòng danh sách sang Excel: {1}", lines.Count, Path.GetFileName(targetPath));
                }
                return ok;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi xuất danh sách sang Excel: " + ex.Message;
                return false;
            }
        }

        #endregion

        #region Helper & Data Extraction Methods

        /// <summary>
        /// Trích xuất cấu trúc dữ liệu 2 chiều (Row x Column) từ Word Table, làm sạch ký tự \r\a.
        /// </summary>
        public static List<List<string>> ExtractTableGrid(Word.Table table)
        {
            List<List<string>> grid = new List<List<string>>();
            if (table == null) return grid;

            try
            {
                int rowCount = table.Rows.Count;
                for (int r = 1; r <= rowCount; r++)
                {
                    Word.Row row = table.Rows[r];
                    List<string> rowData = new List<string>();
                    for (int c = 1; c <= row.Cells.Count; c++)
                    {
                        try
                        {
                            Word.Cell cell = row.Cells[c];
                            rowData.Add(CleanCellText(cell.Range.Text));
                        }
                        catch
                        {
                            rowData.Add(string.Empty);
                        }
                    }
                    grid.Add(rowData);
                }
            }
            catch
            {
                // Đối với các bảng có ô gộp phức tạp, duyệt qua danh sách Range Cells
                try
                {
                    int maxCol = 1;
                    Dictionary<int, List<string>> rowsDict = new Dictionary<int, List<string>>();
                    foreach (Word.Cell cell in table.Range.Cells)
                    {
                        int r = cell.RowIndex;
                        if (!rowsDict.ContainsKey(r))
                        {
                            rowsDict[r] = new List<string>();
                        }
                        rowsDict[r].Add(CleanCellText(cell.Range.Text));
                        if (rowsDict[r].Count > maxCol) maxCol = rowsDict[r].Count;
                    }
                    foreach (var kvp in rowsDict)
                    {
                        grid.Add(kvp.Value);
                    }
                }
                catch { }
            }

            return grid;
        }

        /// <summary>
        /// Làm sạch chuỗi từ cell Word: loại bỏ \r\a, \a và khoảng trắng thừa.
        /// </summary>
        public static string CleanCellText(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return string.Empty;
            return raw.Replace("\r\a", "")
                      .Replace("\a", "")
                      .Replace("\r", " ")
                      .Replace("\n", " ")
                      .Trim();
        }

        /// <summary>
        /// Kiểm tra chuỗi có phải là số kế toán / tiền tệ (ví dụ: "1.250.000", "(500.000)", "15%")
        /// </summary>
        public static bool TryParseAccountingNumber(string text, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;

            string s = text.Trim();
            bool isNegative = false;
            if (s.StartsWith("(") && s.EndsWith(")"))
            {
                isNegative = true;
                s = s.Substring(1, s.Length - 2).Trim();
            }

            // Bỏ các ký hiệu tiền tệ hoặc %
            s = s.Replace("đ", "").Replace("VNĐ", "").Replace("VND", "").Replace("$", "").Replace("%", "").Trim();

            // Xử lý dấu phân cách hàng nghìn kiểu Việt Nam: 1.250.000 hoặc 1,250,000
            if (Regex.IsMatch(s, @"^\-?\d{1,3}(\.\d{3})+(,\d+)?$"))
            {
                s = s.Replace(".", "").Replace(",", ".");
            }
            else if (Regex.IsMatch(s, @"^\-?\d{1,3}(,\d{3})+(\.\d+)?$"))
            {
                s = s.Replace(",", "");
            }

            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
            {
                if (isNegative) value = -value;
                return true;
            }
            return false;
        }

        private static string GetTargetPath(Word.Document doc, string specifiedPath, string defaultSuffix)
        {
            if (!string.IsNullOrEmpty(specifiedPath)) return specifiedPath;

            string fullName = string.Empty;
            try { fullName = doc.FullName; } catch { }

            string dir = string.IsNullOrEmpty(fullName) || !File.Exists(fullName)
                ? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                : Path.GetDirectoryName(fullName);

            string baseName = string.IsNullOrEmpty(fullName)
                ? "CoolWord_Export_" + DateTime.Now.ToString("yyyyMMdd_HHmmss")
                : Path.GetFileNameWithoutExtension(fullName);

            return Path.Combine(dir, baseName + defaultSuffix);
        }

        /// <summary>
        /// Ghi danh sách bảng vào Workbook Excel qua COM Interop.
        /// Tự động áp dụng header style, nhận diện số kế toán và công thức SUM nếu isAccountingMode = true.
        /// </summary>
        private static bool WriteTablesToExcelWorkbook(
            List<List<List<string>>> tables,
            string targetPath,
            bool separateSheets,
            bool isAccountingMode,
            bool openAfterExport,
            out string statusMessage)
        {
            statusMessage = string.Empty;
            if (tables == null || tables.Count == 0)
            {
                statusMessage = "Không có dữ liệu bảng để ghi.";
                return false;
            }

            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;
                workbook = excelApp.Workbooks.Add(Missing);

                int sheetIndex = 1;

                if (separateSheets && tables.Count > 1)
                {
                    for (int t = 0; t < tables.Count; t++)
                    {
                        Excel.Worksheet ws;
                        if (sheetIndex <= workbook.Sheets.Count)
                        {
                            ws = (Excel.Worksheet)workbook.Sheets[sheetIndex];
                        }
                        else
                        {
                            ws = (Excel.Worksheet)workbook.Sheets.Add(Missing, workbook.Sheets[workbook.Sheets.Count], 1, Missing);
                        }
                        sheetIndex++;
                        try { ws.Name = "Bảng " + (t + 1); } catch { }

                        WriteSingleTableToWorksheet(ws, tables[t], 1, isAccountingMode);
                    }
                }
                else
                {
                    Excel.Worksheet ws = (Excel.Worksheet)workbook.Sheets[1];
                    try { ws.Name = isAccountingMode ? "Kế Toán" : "Dữ Liệu Word"; } catch { }

                    int startRow = 1;
                    for (int t = 0; t < tables.Count; t++)
                    {
                        if (tables.Count > 1)
                        {
                            Excel.Range titleRange = ws.Cells[startRow, 1];
                            titleRange.Value2 = string.Format("BẢNG {0}", t + 1);
                            titleRange.Font.Bold = true;
                            titleRange.Font.Size = 12;
                            startRow++;
                        }

                        int rowsWritten = WriteSingleTableToWorksheet(ws, tables[t], startRow, isAccountingMode);
                        startRow += rowsWritten + 2; // Cách 2 dòng giữa các bảng
                    }
                }

                if (!string.IsNullOrEmpty(targetPath))
                {
                    string dir = Path.GetDirectoryName(targetPath);
                    if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                    object fmt = Excel.XlFileFormat.xlOpenXMLWorkbook;
                    workbook.SaveAs(targetPath, fmt, Missing, Missing, Missing, Missing,
                                    Excel.XlSaveAsAccessMode.xlNoChange, Missing, Missing, Missing, Missing, Missing);
                }

                if (openAfterExport)
                {
                    excelApp.Visible = true;
                    excelApp.UserControl = true;
                }
                else
                {
                    workbook.Close(false, Missing, Missing);
                    excelApp.Quit();
                }

                return true;
            }
            catch (Exception ex)
            {
                // Fallback xuất định dạng CSV nếu Excel COM không khả dụng
                try
                {
                    if (workbook != null) workbook.Close(false, Missing, Missing);
                    if (excelApp != null) excelApp.Quit();
                }
                catch { }

                bool csvOk = WriteTablesToCsvFallback(tables, targetPath, out statusMessage);
                if (csvOk)
                {
                    statusMessage = "Đã xuất bảng ra tệp CSV (tương thích Excel hoàn toàn): " + Path.GetFileName(targetPath);
                    return true;
                }

                statusMessage = "Lỗi khi ghi dữ liệu Excel: " + ex.Message;
                return false;
            }
        }

        private static int WriteSingleTableToWorksheet(Excel.Worksheet ws, List<List<string>> grid, int startRow, bool isAccountingMode)
        {
            if (grid == null || grid.Count == 0) return 0;

            int rowCount = grid.Count;
            int colCount = 0;
            foreach (var r in grid)
            {
                if (r.Count > colCount) colCount = r.Count;
            }
            if (colCount == 0) return 0;

            // Xác định các cột là cột số tiền
            HashSet<int> numericColumns = new HashSet<int>();
            if (isAccountingMode && rowCount > 1)
            {
                for (int c = 0; c < colCount; c++)
                {
                    int numericCells = 0;
                    double dummyVal;
                    for (int r = 1; r < rowCount; r++)
                    {
                        if (c < grid[r].Count && TryParseAccountingNumber(grid[r][c], out dummyVal))
                        {
                            numericCells++;
                        }
                    }
                    // Nếu > 50% số dòng là số, coi cột này là cột số
                    if (numericCells >= (rowCount - 1) / 2 && numericCells > 0)
                    {
                        numericColumns.Add(c + 1);
                    }
                }
            }

            // Ghi dữ liệu
            for (int r = 0; r < rowCount; r++)
            {
                int curRow = startRow + r;
                for (int c = 0; c < grid[r].Count; c++)
                {
                    int curCol = c + 1;
                    string cellVal = grid[r][c];
                    Excel.Range cell = ws.Cells[curRow, curCol];

                    double num;
                    if (r > 0 && TryParseAccountingNumber(cellVal, out num))
                    {
                        cell.Value2 = num;
                        if (isAccountingMode || numericColumns.Contains(curCol))
                        {
                            cell.NumberFormat = "#,##0";
                            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        }
                    }
                    else
                    {
                        cell.Value2 = cellVal;
                    }
                }
            }

            // Định dạng dòng tiêu đề (Header row)
            Excel.Range headerRange = ws.Range[ws.Cells[startRow, 1], ws.Cells[startRow, colCount]];
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(220, 235, 252));
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // Nếu là kế toán: Thêm dòng TỔNG CỘNG với công thức SUM
            int totalRowsAdded = rowCount;
            if (isAccountingMode && numericColumns.Count > 0 && rowCount > 1)
            {
                int sumRow = startRow + rowCount;
                totalRowsAdded++;

                Excel.Range labelCell = ws.Cells[sumRow, 1];
                labelCell.Value2 = "TỔNG CỘNG";
                labelCell.Font.Bold = true;

                foreach (int colIdx in numericColumns)
                {
                    Excel.Range sumCell = ws.Cells[sumRow, colIdx];
                    int startDataRow = startRow + 1;
                    int endDataRow = startRow + rowCount - 1;

                    string colLetter = GetExcelColumnName(colIdx);
                    sumCell.Formula = string.Format("=SUM({0}{1}:{0}{2})", colLetter, startDataRow, endDataRow);
                    sumCell.Font.Bold = true;
                    sumCell.NumberFormat = "#,##0";
                    sumCell.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    sumCell.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDouble;
                }
            }

            // Kẻ viền bảng (Borders) và Tự động căn chỉnh độ rộng cột (AutoFit)
            try
            {
                Excel.Range fullTableRange = ws.Range[ws.Cells[startRow, 1], ws.Cells[startRow + totalRowsAdded - 1, colCount]];
                fullTableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                fullTableRange.Columns.AutoFit();
            }
            catch { }

            return totalRowsAdded;
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            string colName = string.Empty;
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                colName = Convert.ToChar('A' + modulo) + colName;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return colName;
        }

        private static bool WriteTablesToCsvFallback(List<List<List<string>>> tables, string targetPath, out string statusMessage)
        {
            statusMessage = string.Empty;
            try
            {
                string csvPath = Path.ChangeExtension(targetPath, ".csv");
                using (StreamWriter sw = new StreamWriter(csvPath, false, Encoding.UTF8))
                {
                    // Ghi UTF-8 BOM để Excel hiển thị đúng tiếng Việt
                    for (int t = 0; t < tables.Count; t++)
                    {
                        if (tables.Count > 1)
                        {
                            sw.WriteLine(string.Format("\"=== BẢNG {0} ===\"", t + 1));
                        }
                        foreach (var row in tables[t])
                        {
                            List<string> escaped = new List<string>();
                            foreach (var cell in row)
                            {
                                string v = cell ?? "";
                                escaped.Add("\"" + v.Replace("\"", "\"\"") + "\"");
                            }
                            sw.WriteLine(string.Join(",", escaped.ToArray()));
                        }
                        sw.WriteLine();
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = ex.Message;
                return false;
            }
        }

        #endregion
    }
}
