using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OXSpreadsheetExample
{
    public class OXSpreadsheet
    {
        private SpreadsheetDocument _document = null;
        private Dictionary<string, Sheet> _sheets = null;

        /// <summary>
        /// 建立檔案
        /// </summary>
        /// <param name="Filename"></param>
        /// <returns></returns>
        public bool Create(string Filename)
        {
            try
            {
                _document = SpreadsheetDocument.Create(Filename, SpreadsheetDocumentType.Workbook);
                _document.AddWorkbookPart();
                _document.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                _document.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Excel", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                this.Close();
                return false;
            }
        }

        /// <summary>
        /// 建立工作表
        /// </summary>
        /// <param name="SheetName"></param>
        /// <returns></returns>
        public bool CreateSheet(string SheetName)
        {
            try
            {
                var sheetPart = _document.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = _document.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                string relationshipId = _document.WorkbookPart.GetIdOfPart(sheetPart);

                uint sheetId = 1;
                if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                {
                    sheetId =
                        sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = SheetName };
                sheets.Append(sheet);
                return true;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 匯出資料
        /// </summary>
        /// <param name="SheetName"></param>
        /// <param name="Table"></param>
        /// <returns></returns>
        public bool ExportFromDataTable(string SheetName, DataTable Table)
        {
            WorksheetPart _worksheet_part = GetWorksheetPart(SheetName);
            if (_worksheet_part == null) return false;
            Worksheet _worksheet = _worksheet_part.Worksheet;
            SheetData _sheet_data = _worksheet.GetFirstChild<SheetData>();

            DocumentFormat.OpenXml.Spreadsheet.Row _header_row = new DocumentFormat.OpenXml.Spreadsheet.Row();
            List<String> _columns = new List<string>();
            foreach (System.Data.DataColumn column in Table.Columns)
            {
                _columns.Add(column.ColumnName);

                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                _header_row.AppendChild(cell);
            }
            _sheet_data.AppendChild(_header_row);

            foreach (System.Data.DataRow row in Table.Rows)
            {
                DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                foreach (String col in _columns)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    //cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;

                    if (row[col] == DBNull.Value)
                    {
                        cell.DataType = CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("");
                    }
                    else
                    {
                        cell.DataType = CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(row[col].ToString());
                    }

                    newRow.AppendChild(cell);
                }
                _sheet_data.AppendChild(newRow);
            }

            return true;
        }
        /// <summary>
        /// 開啟Excel文件
        /// </summary>
        /// <param name="Filename"></param>
        /// <param name="IsReadOnly"></param>
        /// <returns></returns>
        public bool Open(string Filename, bool isEditable)
        {
            try
            {
                _document = SpreadsheetDocument.Open(Filename, isEditable);
                //_document = SpreadsheetDocument.Open(new System.IO.FileStream(Filename, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite), !IsReadOnly);
                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Excel", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                this.Close();
                return false;
            }
        }
        /// <summary>
        /// 關閉文件
        /// </summary>
        public void Close()
        {
            if (_sheets != null)
            {
                _sheets.Clear();
                _sheets = null;
            }
            if (_document != null)
            {
                _document.Close();
                _document.Dispose();
            }
            _document = null;
        }
        /// <summary>
        /// 讀取Excel指定Sheet
        /// </summary>
        /// <param name="SheetName"></param>
        /// <returns></returns>
        public bool LoadSheet(string SheetName)
        {
            if (_document == null)
            {
                System.Windows.Forms.MessageBox.Show("尚未開啟文件", "ExcelReader", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
            _sheets = new Dictionary<string, Sheet>();

            //取得所有 Sheet
            List<Sheet> _list1 = _document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().ToList();

            //檢查指定的 Sheet
            bool _sheet_exists = false;
            foreach (Sheet sheet in _list1)
            {
                uint _sheetid = sheet.SheetId;
                //取得 Sheet 名稱
                //string _sheetname = _document.WorkbookPart.Workbook.Descendants<Sheet>().ElementAt((int)_sheetid - 1).Name;
                string _sheetname = sheet.Name.Value;
                //比較
                if (string.Compare(_sheetname, SheetName, true) == 0)
                {
                    _sheets.Add(_sheetname, sheet);
                    _sheet_exists = true;
                    break;
                }
            }

            if (_sheet_exists == true) return true;
            else
            {
                _sheets.Clear();
                _sheets = null;
                System.Windows.Forms.MessageBox.Show("指定的Sheet不存在", "ExcelReader", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }
        /// <summary>
        /// 儲存格寫入值
        /// </summary>
        /// <param name="SheetName"></param>
        /// <param name="ColumnName"></param>
        /// <param name="RowIndex"></param>
        /// <param name="Text"></param>
        /// <returns></returns>
        public bool WriteText(string SheetName, string ColumnName, int RowIndex, string Text)
        {
            int shared_string_index = InsertSharedStringItem(Text);

            WorksheetPart worksheetpart = GetWorksheetPart(SheetName);
            if (worksheetpart == null) return false;

            Worksheet worksheet = worksheetpart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = ColumnName + RowIndex;

            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == RowIndex).Count() != 0)
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == RowIndex).First();
            else
            {
                row = new Row() { RowIndex = (uint)RowIndex };
                sheetData.Append(row);
            }
            //===============================
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == ColumnName + RowIndex).Count() > 0)
            {
                Cell cell = row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
                cell.CellValue = new CellValue(shared_string_index.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                return true;
            }
            else
            {
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);
                worksheet.Save();
                newCell.CellValue = new CellValue(shared_string_index.ToString());
                newCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                return true;
            }
        }
        private WorksheetPart GetWorksheetPart(string SheetName)
        {
            if (_document == null) return null;
            string relId = _document.WorkbookPart.Workbook.Descendants<Sheet>().First(s => SheetName.Equals(s.Name)).Id;
            return (WorksheetPart)_document.WorkbookPart.GetPartById(relId);
        }

        private int InsertSharedStringItem(string Text)
        {
            SharedStringTablePart shareStringPart;
            if (_document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                shareStringPart = _document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            else
                shareStringPart = _document.WorkbookPart.AddNewPart<SharedStringTablePart>();

            if (shareStringPart.SharedStringTable == null)
                shareStringPart.SharedStringTable = new SharedStringTable();


            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == Text)
                {
                    return i;
                }

                i++;
            }

            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(Text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }
        /// <summary>
        /// Inserts a new row at the desired index. If one already exists, then it is
        /// returned. If an insertRow is provided, then it is inserted into the desired
        /// rowIndex
        /// 於下方插入列
        /// </summary>
        /// <param name="rowIndex">Row Index</param>
        /// <param name="worksheetPart">Worksheet Part</param>
        /// <param name="insertRow">Row to insert</param>
        /// <param name="isLastRow">Optional parameter - True, you can guarantee that this row is the last row (not replacing an existing last row) in the sheet to insert; false it is not</param>
        /// <returns>Inserted Row</returns>
        public Row InsertRow(string sheetName, uint rowIndex, bool isNewLastRow = false)
        {
            WorksheetPart worksheetPart = (WorksheetPart)_document.WorkbookPart.GetPartById(_document.WorkbookPart.Workbook.Descendants<Sheet>().First((s) => s.Name == sheetName).Id);
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            Row currentRow;
            Row cloneRow;
            //將指定列取出
            currentRow = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            //將指定列複製
            cloneRow = (Row)currentRow.Clone();

            Row retRow = !isNewLastRow ? sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex) : null;

            // If the worksheet does not contain a row with the specified row index, insert one.
            if (retRow != null)
            {
                // if retRow is not null and we are inserting a new row, then move all existing rows down.
                if (cloneRow != null)
                {
                    UpdateRowIndexes(worksheetPart, rowIndex, false);
                    UpdateMergedCellReferences(worksheetPart, rowIndex, false);
                    UpdateHyperlinkReferences(worksheetPart, rowIndex, false);

                    // actually insert the new row into the sheet
                    retRow = sheetData.InsertBefore(cloneRow, retRow);  // at this point, retRow still points to the row that had the insert rowIndex

                    string curIndex = retRow.RowIndex.ToString();
                    string newIndex = rowIndex.ToString();

                    foreach (Cell cell in retRow.Elements<Cell>())
                    {
                        // Update the references for the rows cells.
                        cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curIndex, newIndex));
                    }

                    // Update the row index.
                    retRow.RowIndex = rowIndex;
                }
            }
            else
            {
                // Row doesn't exist yet, shifting not needed.
                // Rows must be in sequential order according to RowIndex. Determine where to insert the new row.
                Row refRow = !isNewLastRow ? sheetData.Elements<Row>().FirstOrDefault(row => row.RowIndex > rowIndex) : null;

                // use the insert row if it exists
                retRow = cloneRow ?? new Row() { RowIndex = rowIndex };

                IEnumerable<Cell> cellsInRow = retRow.Elements<Cell>();

                if (cellsInRow.Any())
                {
                    string curIndex = retRow.RowIndex.ToString();
                    string newIndex = rowIndex.ToString();

                    foreach (Cell cell in cellsInRow)
                    {
                        // Update the references for the rows cells.
                        cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curIndex, newIndex));
                    }

                    // Update the row index.
                    retRow.RowIndex = rowIndex;
                }

                sheetData.InsertBefore(retRow, refRow);
            }

            worksheet.Save();
            return retRow;
        }
        /// <summary>
        /// Updates all of the Row indexes and the child Cells' CellReferences whenever
        /// a row is inserted or deleted.
        /// 每當插入或刪除行時，更新所有 Row 索引和子 Cells 的 CellReferences。
        /// </summary>
        /// <param name="worksheetPart">Worksheet Part</param>
        /// <param name="rowIndex">Row Index being inserted or deleted</param>
        /// <param name="isDeletedRow">True if row was deleted, otherwise false</param>
        private static void UpdateRowIndexes(WorksheetPart worksheetPart, uint rowIndex, bool isDeletedRow)
        {
            // Get all the rows in the worksheet with equal or higher row index values than the one being inserted/deleted for reindexing.
            IEnumerable<Row> rows = worksheetPart.Worksheet.Descendants<Row>().Where(r => r.RowIndex.Value >= rowIndex);

            foreach (Row row in rows)
            {
                //超過最大列數則跳離
                if (row.RowIndex > worksheetPart.Worksheet.Descendants<Row>().ToList().Last().RowIndex) break;

                uint newIndex = (isDeletedRow ? row.RowIndex - 1 : row.RowIndex + 1);
                string curRowIndex = row.RowIndex.ToString();
                string newRowIndex = newIndex.ToString();

                foreach (Cell cell in row.Elements<Cell>())
                {
                    // Update the references for the rows cells.
                    cell.CellReference = new StringValue(cell.CellReference.Value.Replace(curRowIndex, newRowIndex));
                }

                // Update the row index.
                row.RowIndex = newIndex;
            }
        }

        /// <summary>
        /// Updates the MergedCelss reference whenever a new row is inserted or deleted. It will simply take the
        /// row index and either increment or decrement the cell row index in the merged cell reference based on
        /// if the row was inserted or deleted.
        /// 每當插入或刪除新行時更新 MergedCelss 引用。 它將簡單地採用行索引，並根據行是插入還是刪除來增加或減少合併單元格引用中的單元格行索引。
        /// </summary>
        /// <param name="worksheetPart">Worksheet Part</param>
        /// <param name="rowIndex">Row Index being inserted or deleted</param>
        /// <param name="isDeletedRow">True if row was deleted, otherwise false</param>
        private static void UpdateMergedCellReferences(WorksheetPart worksheetPart, uint rowIndex, bool isDeletedRow)
        {
            if (worksheetPart.Worksheet.Elements<MergeCells>().Count() > 0)
            {
                MergeCells mergeCells = worksheetPart.Worksheet.Elements<MergeCells>().FirstOrDefault();

                if (mergeCells != null)
                {
                    // Grab all the merged cells that have a merge cell row index reference equal to or greater than the row index passed in
                    List<MergeCell> mergeCellsList = mergeCells.Elements<MergeCell>().Where(r => r.Reference.HasValue)
                                                                                     .Where(r => GetRowIndex(r.Reference.Value.Split(':').ElementAt(0)) >= rowIndex ||
                                                                                                 GetRowIndex(r.Reference.Value.Split(':').ElementAt(1)) >= rowIndex).ToList();

                    // Need to remove all merged cells that have a matching rowIndex when the row is deleted
                    if (isDeletedRow)
                    {
                        List<MergeCell> mergeCellsToDelete = mergeCellsList.Where(r => GetRowIndex(r.Reference.Value.Split(':').ElementAt(0)) == rowIndex ||
                                                                                       GetRowIndex(r.Reference.Value.Split(':').ElementAt(1)) == rowIndex).ToList();

                        // Delete all the matching merged cells
                        foreach (MergeCell cellToDelete in mergeCellsToDelete)
                        {
                            cellToDelete.Remove();
                        }

                        // Update the list to contain all merged cells greater than the deleted row index
                        mergeCellsList = mergeCells.Elements<MergeCell>().Where(r => r.Reference.HasValue)
                                                                         .Where(r => GetRowIndex(r.Reference.Value.Split(':').ElementAt(0)) > rowIndex ||
                                                                                     GetRowIndex(r.Reference.Value.Split(':').ElementAt(1)) > rowIndex).ToList();
                    }

                    // Either increment or decrement the row index on the merged cell reference
                    foreach (MergeCell mergeCell in mergeCellsList)
                    {
                        string[] cellReference = mergeCell.Reference.Value.Split(':');

                        if (GetRowIndex(cellReference.ElementAt(0)) >= rowIndex)
                        {
                            string columnName = GetColumnName(cellReference.ElementAt(0));
                            cellReference[0] = isDeletedRow ? columnName + (GetRowIndex(cellReference.ElementAt(0)) - 1).ToString() : IncrementCellReference(cellReference.ElementAt(0), CellReferencePartEnum.Row);
                        }

                        if (GetRowIndex(cellReference.ElementAt(1)) >= rowIndex)
                        {
                            string columnName = GetColumnName(cellReference.ElementAt(1));
                            cellReference[1] = isDeletedRow ? columnName + (GetRowIndex(cellReference.ElementAt(1)) - 1).ToString() : IncrementCellReference(cellReference.ElementAt(1), CellReferencePartEnum.Row);
                        }

                        mergeCell.Reference = new StringValue(cellReference[0] + ":" + cellReference[1]);
                    }
                }
            }
        }

        /// <summary>
        /// Updates all hyperlinks in the worksheet when a row is inserted or deleted.
        /// 插入或刪除行時更新工作表中的所有超鏈接。
        /// </summary>
        /// <param name="worksheetPart">Worksheet Part</param>
        /// <param name="rowIndex">Row Index being inserted or deleted</param>
        /// <param name="isDeletedRow">True if row was deleted, otherwise false</param>
        private static void UpdateHyperlinkReferences(WorksheetPart worksheetPart, uint rowIndex, bool isDeletedRow)
        {
            Hyperlinks hyperlinks = worksheetPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();

            if (hyperlinks != null)
            {
                Match hyperlinkRowIndexMatch;
                uint hyperlinkRowIndex;

                foreach (Hyperlink hyperlink in hyperlinks.Elements<Hyperlink>())
                {
                    hyperlinkRowIndexMatch = Regex.Match(hyperlink.Reference.Value, "[0-9]+");
                    if (hyperlinkRowIndexMatch.Success && uint.TryParse(hyperlinkRowIndexMatch.Value, out hyperlinkRowIndex) && hyperlinkRowIndex >= rowIndex)
                    {
                        // if being deleted, hyperlink needs to be removed or moved up
                        if (isDeletedRow)
                        {
                            // if hyperlink is on the row being removed, remove it
                            if (hyperlinkRowIndex == rowIndex)
                            {
                                hyperlink.Remove();
                            }
                            // else hyperlink needs to be moved up a row
                            else
                            {
                                hyperlink.Reference.Value = hyperlink.Reference.Value.Replace(hyperlinkRowIndexMatch.Value, (hyperlinkRowIndex - 1).ToString());

                            }
                        }
                        // else row is being inserted, move hyperlink down
                        else
                        {
                            hyperlink.Reference.Value = hyperlink.Reference.Value.Replace(hyperlinkRowIndexMatch.Value, (hyperlinkRowIndex + 1).ToString());
                        }
                    }
                }

                // Remove the hyperlinks collection if none remain
                if (hyperlinks.Elements<Hyperlink>().Count() == 0)
                {
                    hyperlinks.Remove();
                }
            }
        }

        /// <summary>
        /// Given a cell name, parses the specified cell to get the row index.
        /// </summary>
        /// <param name="cellReference">Address of the cell (ie. B2)</param>
        /// <returns>Row Index (ie. 2)</returns>
        public static uint GetRowIndex(string cellReference)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellReference);

            return uint.Parse(match.Value);
        }

        /// <summary>
        /// Increments the reference of a given cell.  This reference comes from the CellReference property
        /// on a Cell.
        /// </summary>
        /// <param name="reference">reference string</param>
        /// <param name="cellRefPart">indicates what is to be incremented</param>
        /// <returns></returns>
        public static string IncrementCellReference(string reference, CellReferencePartEnum cellRefPart)
        {
            string newReference = reference;

            if (cellRefPart != CellReferencePartEnum.None && !String.IsNullOrEmpty(reference))
            {
                string[] parts = Regex.Split(reference, "([A-Z]+)");

                if (cellRefPart == CellReferencePartEnum.Column || cellRefPart == CellReferencePartEnum.Both)
                {
                    List<char> col = parts[1].ToCharArray().ToList();
                    bool needsIncrement = true;
                    int index = col.Count - 1;

                    do
                    {
                        // increment the last letter
                        col[index] = Letters[Letters.IndexOf(col[index]) + 1];

                        // if it is the last letter, then we need to roll it over to 'A'
                        if (col[index] == Letters[Letters.Count - 1])
                        {
                            col[index] = Letters[0];
                        }
                        else
                        {
                            needsIncrement = false;
                        }

                    } while (needsIncrement && --index >= 0);

                    // If true, then we need to add another letter to the mix. Initial value was something like "ZZ"
                    if (needsIncrement)
                    {
                        col.Add(Letters[0]);
                    }

                    parts[1] = new String(col.ToArray());
                }

                if (cellRefPart == CellReferencePartEnum.Row || cellRefPart == CellReferencePartEnum.Both)
                {
                    // Increment the row number. A reference is invalid without this componenet, so we assume it will always be present.
                    parts[2] = (int.Parse(parts[2]) + 1).ToString();
                }

                newReference = parts[1] + parts[2];
            }

            return newReference;
        }
        /// <summary>
        /// Given a cell name, parses the specified cell to get the column name.
        /// </summary>
        /// <param name="cellReference">Address of the cell (ie. B2)</param>
        /// <returns>Column name (ie. A2)</returns>
        private static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }
        public enum CellReferencePartEnum
        {
            None,
            Column,
            Row,
            Both
        }
        private static List<char> Letters = new List<char>() { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', ' ' };

    }
}
