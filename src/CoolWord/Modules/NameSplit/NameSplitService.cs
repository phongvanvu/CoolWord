using System;
using System.Text.RegularExpressions;
using CoolWord.Core;
using Word = Microsoft.Office.Interop.Word;

namespace CoolWord.Modules.NameSplit
{
    public static class NameSplitService
    {
        private static object Missing = Type.Missing;

        /// <summary>
        /// Tách cột Họ và tên trong bảng Word thành 2 cột: Cột 1 (Họ và tên đệm), Cột 2 (Tên).
        /// </summary>
        public static bool SplitTableColumn2(Word.Application app, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                Word.Selection sel = app.Selection;
                if (sel == null || sel.Tables.Count == 0)
                {
                    statusMessage = "Vui lòng đặt con trỏ chuột vào cột Họ và tên trong Bảng cần tách.";
                    return false;
                }

                Word.Table table = sel.Tables[1];
                int colIndex = 1;
                try
                {
                    colIndex = sel.Cells[1].ColumnIndex;
                }
                catch
                {
                    statusMessage = "Không thể xác định vị trí cột trong bảng. Vui lòng đặt con trỏ vào 1 ô cụ thể.";
                    return false;
                }

                StartUndoRecord(app, "CoolWord: Tách Họ - Tên (2 cột)");
                try
                {
                    Word.Column newCol;
                    try
                    {
                        if (colIndex < table.Columns.Count)
                        {
                            object beforeCol = table.Columns[colIndex + 1];
                            newCol = table.Columns.Add(ref beforeCol);
                        }
                        else
                        {
                            newCol = table.Columns.Add(ref Missing);
                        }
                    }
                    catch (Exception ex)
                    {
                        statusMessage = "Không thể chèn cột mới (bảng có thể có ô bị gộp): " + ex.Message;
                        return false;
                    }

                    // Điều chỉnh kích thước cột
                    try
                    {
                        float origWidth = table.Columns[colIndex].Width;
                        table.Columns[colIndex].Width = origWidth * 0.65f; // Họ đệm chiếm 65%
                        newCol.Width = origWidth * 0.35f;                 // Tên chiếm 35%
                    }
                    catch { }

                    int rowCount = table.Rows.Count;
                    int processedRows = 0;

                    for (int r = 1; r <= rowCount; r++)
                    {
                        Word.Cell cellOld = null;
                        Word.Cell cellNew = null;

                        try
                        {
                            cellOld = table.Cell(r, colIndex);
                            cellNew = table.Cell(r, colIndex + 1);
                        }
                        catch { continue; }

                        if (cellOld == null || cellNew == null) continue;

                        string raw = CleanCellText(cellOld.Range.Text);
                        if (string.IsNullOrWhiteSpace(raw)) continue;

                        // Nếu là dòng tiêu đề đầu tiên
                        if (r == 1 && IsHeaderCell(raw))
                        {
                            cellOld.Range.Text = "Họ và tên đệm";
                            cellNew.Range.Text = "Tên";
                            processedRows++;
                            continue;
                        }

                        ParsedVietnameseName parsed = VietnameseNameSplitter.Split(raw);
                        cellOld.Range.Text = parsed.HoDem;
                        cellNew.Range.Text = parsed.Ten;
                        processedRows++;
                    }

                    statusMessage = string.Format("Đã tách thành công {0} dòng trong bảng thành 2 cột (Họ đệm & Tên).", processedRows);
                    return true;
                }
                finally
                {
                    EndUndoRecord(app);
                }
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi tách Họ - Tên trong bảng: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Tách cột Họ và tên trong bảng Word thành 3 cột: Cột 1 (Họ), Cột 2 (Tên đệm), Cột 3 (Tên).
        /// </summary>
        public static bool SplitTableColumn3(Word.Application app, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                Word.Selection sel = app.Selection;
                if (sel == null || sel.Tables.Count == 0)
                {
                    statusMessage = "Vui lòng đặt con trỏ chuột vào cột Họ và tên trong Bảng cần tách.";
                    return false;
                }

                Word.Table table = sel.Tables[1];
                int colIndex = 1;
                try
                {
                    colIndex = sel.Cells[1].ColumnIndex;
                }
                catch
                {
                    statusMessage = "Không thể xác định vị trí cột trong bảng.";
                    return false;
                }

                StartUndoRecord(app, "CoolWord: Tách Họ - Tên (3 cột)");
                try
                {
                    Word.Column col2, col3;
                    try
                    {
                        if (colIndex < table.Columns.Count)
                        {
                            object beforeCol = table.Columns[colIndex + 1];
                            col2 = table.Columns.Add(ref beforeCol);
                            col3 = table.Columns.Add(ref beforeCol);
                        }
                        else
                        {
                            col2 = table.Columns.Add(ref Missing);
                            col3 = table.Columns.Add(ref Missing);
                        }
                    }
                    catch (Exception ex)
                    {
                        statusMessage = "Không thể chèn cột mới (bảng có thể có ô bị gộp): " + ex.Message;
                        return false;
                    }

                    try
                    {
                        float origWidth = table.Columns[colIndex].Width;
                        table.Columns[colIndex].Width = origWidth * 0.35f;     // Họ
                        col2.Width = origWidth * 0.35f;                        // Đệm
                        col3.Width = origWidth * 0.30f;                        // Tên
                    }
                    catch { }

                    int rowCount = table.Rows.Count;
                    int processedRows = 0;

                    for (int r = 1; r <= rowCount; r++)
                    {
                        Word.Cell c1 = null, c2 = null, c3 = null;
                        try
                        {
                            c1 = table.Cell(r, colIndex);
                            c2 = table.Cell(r, colIndex + 1);
                            c3 = table.Cell(r, colIndex + 2);
                        }
                        catch { continue; }

                        if (c1 == null || c2 == null || c3 == null) continue;

                        string raw = CleanCellText(c1.Range.Text);
                        if (string.IsNullOrWhiteSpace(raw)) continue;

                        if (r == 1 && IsHeaderCell(raw))
                        {
                            c1.Range.Text = "Họ";
                            c2.Range.Text = "Tên đệm";
                            c3.Range.Text = "Tên";
                            processedRows++;
                            continue;
                        }

                        ParsedVietnameseName parsed = VietnameseNameSplitter.Split(raw);
                        c1.Range.Text = parsed.Ho;
                        c2.Range.Text = parsed.Dem;
                        c3.Range.Text = parsed.Ten;
                        processedRows++;
                    }

                    statusMessage = string.Format("Đã tách thành công {0} dòng trong bảng thành 3 cột (Họ | Đệm | Tên).", processedRows);
                    return true;
                }
                finally
                {
                    EndUndoRecord(app);
                }
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi tách Họ - Tên: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Gộp 2 cột liền kề (Cột Họ đệm và Cột Tên) thành 1 cột Họ và tên duy nhất.
        /// </summary>
        public static bool MergeTableColumns(Word.Application app, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                Word.Selection sel = app.Selection;
                if (sel == null || sel.Tables.Count == 0)
                {
                    statusMessage = "Vui lòng đặt con trỏ chuột vào cột Họ đệm (cột bên trái cột Tên) trong Bảng.";
                    return false;
                }

                Word.Table table = sel.Tables[1];
                int colIndex = 1;
                try
                {
                    colIndex = sel.Cells[1].ColumnIndex;
                }
                catch
                {
                    statusMessage = "Không thể xác định vị trí cột trong bảng.";
                    return false;
                }

                if (colIndex >= table.Columns.Count)
                {
                    statusMessage = "Cột đang chọn là cột cuối cùng của bảng, không có cột kế tiếp để gộp.";
                    return false;
                }

                StartUndoRecord(app, "CoolWord: Gộp Họ và Tên");
                try
                {
                    int rowCount = table.Rows.Count;
                    int processedRows = 0;

                    for (int r = 1; r <= rowCount; r++)
                    {
                        Word.Cell c1 = null, c2 = null;
                        try
                        {
                            c1 = table.Cell(r, colIndex);
                            c2 = table.Cell(r, colIndex + 1);
                        }
                        catch { continue; }

                        if (c1 == null || c2 == null) continue;

                        string t1 = CleanCellText(c1.Range.Text);
                        string t2 = CleanCellText(c2.Range.Text);

                        if (r == 1 && (IsHeaderCell(t1) || IsHeaderCell(t2)))
                        {
                            c1.Range.Text = "Họ và tên";
                            processedRows++;
                            continue;
                        }

                        c1.Range.Text = VietnameseNameSplitter.Merge(t1, t2);
                        processedRows++;
                    }

                    // Xóa cột thứ 2 (cột Tên cũ sau khi đã gộp vào cột 1)
                    try
                    {
                        table.Columns[colIndex + 1].Delete();
                    }
                    catch { }

                    statusMessage = string.Format("Đã gộp thành công {0} dòng thành cột \"Họ và tên\" hoàn chỉnh.", processedRows);
                    return true;
                }
                finally
                {
                    EndUndoRecord(app);
                }
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi gộp cột Họ và Tên: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Tách danh sách Họ và tên trong vùng văn bản đang bôi đen (ngăn cách bằng dấu Tab hoặc gạch đứng |).
        /// </summary>
        public static bool SplitSelectedTextLines(Word.Application app, string delimiter, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                Word.Selection sel = app.Selection;
                if (sel == null || sel.Range == null || sel.Range.Start == sel.Range.End)
                {
                    statusMessage = "Vui lòng bôi đen danh sách họ tên cần tách.";
                    return false;
                }

                string rawText = sel.Range.Text;
                if (string.IsNullOrWhiteSpace(rawText))
                {
                    statusMessage = "Vùng chọn trống.";
                    return false;
                }

                string sep = delimiter ?? "\t";
                string[] lines = rawText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                int count = 0;

                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i];
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    ParsedVietnameseName parsed = VietnameseNameSplitter.Split(line);
                    if (!string.IsNullOrEmpty(parsed.Ten))
                    {
                        lines[i] = string.IsNullOrEmpty(parsed.HoDem)
                            ? parsed.Ten
                            : parsed.HoDem + sep + parsed.Ten;
                        count++;
                    }
                }

                string replaced = string.Join("\r\n", lines);
                StartUndoRecord(app, "CoolWord: Tách Họ - Tên danh sách");
                try
                {
                    sel.Range.Text = replaced;
                }
                finally
                {
                    EndUndoRecord(app);
                }

                statusMessage = string.Format("Đã tách Họ - Tên cho {0} dòng văn bản.", count);
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi tách Họ - Tên: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Chuẩn hóa viết hoa danh từ riêng cho toàn bộ họ và tên trong vùng chọn hoặc bảng.
        /// </summary>
        public static bool NormalizeSelectedNames(Word.Application app, out string statusMessage)
        {
            statusMessage = string.Empty;
            if (app == null || app.Documents.Count == 0)
            {
                statusMessage = "Không tìm thấy tài liệu Word đang mở.";
                return false;
            }

            try
            {
                Word.Selection sel = app.Selection;
                if (sel == null || sel.Range == null || sel.Range.Start == sel.Range.End)
                {
                    statusMessage = "Vui lòng bôi đen danh sách tên cần chuẩn hóa.";
                    return false;
                }

                string rawText = sel.Range.Text;
                if (string.IsNullOrWhiteSpace(rawText))
                {
                    statusMessage = "Vùng chọn trống.";
                    return false;
                }

                string[] lines = rawText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                int count = 0;

                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i];
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    string normalized = VietnameseNameSplitter.NormalizeName(line);
                    if (!string.Equals(line, normalized, StringComparison.Ordinal))
                    {
                        lines[i] = normalized;
                        count++;
                    }
                }

                string replaced = string.Join("\r\n", lines);
                StartUndoRecord(app, "CoolWord: Chuẩn hóa Họ Tên");
                try
                {
                    sel.Range.Text = replaced;
                }
                finally
                {
                    EndUndoRecord(app);
                }

                statusMessage = string.Format("Đã chuẩn hóa viết hoa danh từ riêng cho {0} tên.", count);
                return true;
            }
            catch (Exception ex)
            {
                statusMessage = "Lỗi khi chuẩn hóa họ tên: " + ex.Message;
                return false;
            }
        }

        private static string CleanCellText(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            return text.TrimEnd('\r', '\n', '\a', '\t', ' ');
        }

        private static bool IsHeaderCell(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            string lower = text.Trim().ToLowerInvariant();
            return lower.Contains("họ") || lower.Contains("tên") || lower.Contains("họ và tên") || lower.Contains("họ tên");
        }

        private static void StartUndoRecord(Word.Application app, string name)
        {
            try
            {
                if (app.UndoRecord != null && !app.UndoRecord.IsRecordingCustomRecord)
                {
                    app.UndoRecord.StartCustomRecord(name);
                }
            }
            catch { }
        }

        private static void EndUndoRecord(Word.Application app)
        {
            try
            {
                if (app.UndoRecord != null && app.UndoRecord.IsRecordingCustomRecord)
                {
                    app.UndoRecord.EndCustomRecord();
                }
            }
            catch { }
        }
    }
}
