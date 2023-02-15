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
                WorkbookPart wbp = _document.AddWorkbookPart();
                _document.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                _document.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                WorkbookStylesPart wbsp = wbp.AddNewPart<WorkbookStylesPart>();
                // add styles to sheet
                wbsp.Stylesheet = CreateStylesheet();
                wbsp.Stylesheet.Save();

                //========== 定義 WorkBook 相容層級
                //WorkbookStylesPart workbookStylesPart = _document.WorkbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");
                //Stylesheet stylesheet = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
                //stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                //stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

                ////========== 宣告字型
                //Fonts fonts1 = new Fonts() { Count = (UInt32Value)1U, KnownFonts = true };
                //Font font1 = new Font();
                ////========== 定義字型大小
                //FontSize fontSize1 = new FontSize() { Val = new DoubleValue(_font_size).Value };
                ////========== 定義字型大小
                //Color color1 = new Color() { Theme = (UInt32Value)1U };
                ////========== 定義字型
                //FontName fontName1 = new FontName() { Val = _font.Name };
                //FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
                //FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };
                //font1.Append(fontSize1);
                //font1.Append(color1);
                //font1.Append(fontName1);
                //font1.Append(fontFamilyNumbering1);
                //font1.Append(fontScheme1);
                //fonts1.Append(font1);

                ////========== 
                //Fill fill3 = new Fill();
                //PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
                //ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "FFFF0000" };
                //BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };
                //patternFill3.Append(foregroundColor1);
                //patternFill3.Append(backgroundColor1);
                //fill3.Append(patternFill3);
                //stylesheet.Append(fonts1);
                //workbookStylesPart.Stylesheet = stylesheet;

                //Sheets sheets = _document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

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
        public bool CreateColumns(string SheetName, Dictionary<string, double> ColumnsWidth)
        {
            WorksheetPart worksheetpart = GetWorksheetPart(SheetName);
            if (worksheetpart == null) return false;

            // Create custom widths for columns
            Columns lstColumns = worksheetpart.Worksheet.GetFirstChild<Columns>();
            bool needToInsertColumns = false;
            if (lstColumns == null)
            {
                lstColumns = new Columns();
                needToInsertColumns = true;
            }

            // Min = 1, Max = 1 ==> Apply this to column 1 (A)
            // Min = 2, Max = 2 ==> Apply this to column 2 (B)
            // Width = 25 ==> Set the width to 25
            // CustomWidth = true ==> Tell Excel to use the custom width
            int colIndex = 1;
            foreach (KeyValuePair<string, double> keyValue in ColumnsWidth)
            {
                lstColumns.Append(new Column() { Min = (uint)colIndex, Max = (uint)colIndex, Width = keyValue.Value, CustomWidth = true });
                colIndex++;
            }

            // Only insert the columns if we had to create a new columns element
            if (needToInsertColumns)
                worksheetpart.Worksheet.InsertAt(lstColumns, 0);

            worksheetpart.Worksheet.Save();
            return true;
        }
        /// <summary>
        /// 開啟Excel文件
        /// </summary>
        /// <param name="Filename"></param>
        /// <param name="IsReadOnly"></param>
        /// <returns></returns>
        public bool Open(string Filename, bool IsReadOnly = true)
        {
            try
            {
                _document = SpreadsheetDocument.Open(Filename, !IsReadOnly);
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
                cell.StyleIndex = 0;
                return true;
            }
            else
            {
                //Cell refCell = null;
                //foreach (Cell cell in row.Elements<Cell>())
                //{
                //    if (cell.CellReference.Value.Length == cellReference.Length)
                //    {
                //        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                //        {
                //            refCell = cell;
                //            break;
                //        }
                //    }
                //}

                Cell newCell = new Cell();
                newCell.CellReference = cellReference;
                newCell.CellValue = new CellValue(shared_string_index.ToString());
                newCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                newCell.StyleIndex = 0;
                //row.InsertBefore(newCell, refCell);
                row.Append(newCell);

                //worksheet.Save();
                return true;
            }
        }

        /// <summary>
        /// 儲存格寫入值
        /// 日期格式
        /// </summary>
        /// <param name="SheetName"></param>
        /// <param name="ColumnName"></param>
        /// <param name="RowIndex"></param>
        /// <param name="Text"></param>
        /// <returns></returns>
        public bool WriteDate(string SheetName, string ColumnName, int RowIndex, string dateTime)
        {
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
                cell.CellValue = new CellValue(dateTime);
                cell.DataType = new EnumValue<CellValues>(CellValues.Date);
                cell.StyleIndex = 0;
                return true;
            }
            else
            {
                //Cell refCell = null;
                //foreach (Cell cell in row.Elements<Cell>())
                //{
                //    if (cell.CellReference.Value.Length == cellReference.Length)
                //    {
                //        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                //        {
                //            refCell = cell;
                //            break;
                //        }
                //    }
                //}

                Cell newCell = new Cell();
                newCell.CellReference = cellReference;
                newCell.CellValue = new CellValue(dateTime);
                newCell.DataType = new EnumValue<CellValues>(CellValues.Date);
                newCell.StyleIndex = 0;
                //row.InsertBefore(newCell, refCell);
                row.Append(newCell);

                //worksheet.Save();
                return true;
            }
        }

        public bool SetCellStyle(string SheetName, string ColumnName, int RowIndex, CustomizeStyle customizeStyle)
        {
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
                //設定格式
                if (customizeStyle == CustomizeStyle.藍底黑框)
                    cell.StyleIndex = (UInt32Value)1U;
                else if (customizeStyle == CustomizeStyle.白底黑框)
                    cell.StyleIndex = (UInt32Value)2U;
                else if (customizeStyle == CustomizeStyle.上下藍框)
                    cell.StyleIndex = (UInt32Value)3U;
                else
                    cell.StyleIndex = (UInt32Value)0U;
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
                //worksheet.Save();
                //設定格式
                if (customizeStyle == CustomizeStyle.藍底黑框)
                    newCell.StyleIndex = (UInt32Value)1U;
                else if (customizeStyle == CustomizeStyle.白底黑框)
                    newCell.StyleIndex = (UInt32Value)2U;
                else if (customizeStyle == CustomizeStyle.上下藍框)
                    newCell.StyleIndex = (UInt32Value)3U;
                else
                    newCell.StyleIndex = (UInt32Value)0U;
                return true;
            }
        }
        public bool WorkSheetSave(string SheetName)
        {
            WorksheetPart worksheetpart = GetWorksheetPart(SheetName);
            if (worksheetpart == null) return false;

            Worksheet worksheet = worksheetpart.Worksheet;
            worksheet.Save();
            return true;
        }
        private WorksheetPart GetWorksheetPart(string SheetName)
        {
            if (_document == null) return null;
            string relId = _document.WorkbookPart.Workbook.Descendants<Sheet>().First(s => SheetName.Equals(s.Name)).Id;
            return (WorksheetPart)_document.WorkbookPart.GetPartById(relId);
        }
        /// <summary>
        /// 取得 Excel 欄位索引名稱)
        /// </summary>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public string GetExcelColumnName(int columnIndex)
        {
            //  例:  (0) should return "A"
            //       (1) should return "B"
            //       (25) should return "Z"
            //       (26) should return "AA"
            //       (27) should return "AB"
            //       ..etc..
            char firstChar;
            char secondChar;
            char thirdChar;

            if (columnIndex < 26)
            {
                return ((char)('A' + columnIndex)).ToString();
            }

            if (columnIndex < 702)
            {
                firstChar = (char)('A' + (columnIndex / 26) - 1);
                secondChar = (char)('A' + (columnIndex % 26));

                return string.Format("{0}{1}", firstChar, secondChar);
            }

            int firstInt = columnIndex / 26 / 26;
            int secondInt = (columnIndex - firstInt * 26 * 26) / 26;
            if (secondInt == 0)
            {
                secondInt = 26;
                firstInt = firstInt - 1;
            }
            int thirdInt = (columnIndex - firstInt * 26 * 26 - secondInt * 26);

            firstChar = (char)('A' + firstInt - 1);
            secondChar = (char)('A' + secondInt - 1);
            thirdChar = (char)('A' + thirdInt);

            return string.Format("{0}{1}{2}", firstChar, secondChar, thirdChar);
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

        private static List<char> Letters = new List<char>() { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', ' ' };
        /// <summary>
        /// 建立預設格式
        /// </summary>
        /// <returns></returns>
        private static Stylesheet CreateStylesheet()
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)1U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 12D };//字型大小
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "微軟正黑體" };//字型
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.None }; //字型方案

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            fonts1.Append(font1);

            Fills fills1 = new Fills() { Count = (UInt32Value)5U };

            // FillId = 0
            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };
            fill1.Append(patternFill1);

            // FillId = 1
            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };
            fill2.Append(patternFill2);

            // FillId = 2,藍色 DCE6F1
            Fill fill3 = new Fill();
            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = new HexBinaryValue("DCE6F1") };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);
            fill3.Append(patternFill3);

            // FillId = 3,BLUE
            Fill fill4 = new Fill();
            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "FF0070C0" };
            BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill4.Append(foregroundColor2);
            patternFill4.Append(backgroundColor2);
            fill4.Append(patternFill4);

            // FillId = 4,YELLO
            Fill fill5 = new Fill();
            PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor3 = new ForegroundColor() { Rgb = "FFFFFF00" };
            BackgroundColor backgroundColor3 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill5.Append(foregroundColor3);
            patternFill5.Append(backgroundColor3);
            fill5.Append(patternFill5);

            fills1.Append(fill1);
            fills1.Append(fill2);
            fills1.Append(fill3);
            fills1.Append(fill4);
            fills1.Append(fill5);

            Borders borders1 = new Borders() { Count = (UInt32Value)3U };

            //Border = 0,無框線
            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            borders1.Append(border1);

            //Border = 1,上下左右實線 黑色
            Border border2 = new Border();
            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };//實線
            Color leftBorder2Color = new Color() { Indexed = (UInt32Value)64U };
            leftBorder2.Append(leftBorder2Color);
            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };//實線
            Color rightBorder2Color = new Color() { Indexed = (UInt32Value)64U };
            rightBorder2.Append(rightBorder2Color);
            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };//實線
            Color topBorder2Color = new Color() { Indexed = (UInt32Value)64U };
            topBorder2.Append(topBorder2Color);
            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };//實線
            Color bottomBorder2Color = new Color() { Indexed = (UInt32Value)64U };
            bottomBorder2.Append(bottomBorder2Color);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border2);

            //Border = 2,上實線 下雙實線 藍色 #4F81BD
            Border border3 = new Border();
            LeftBorder leftBorder3 = new LeftBorder();
            RightBorder rightBorder3 = new RightBorder();
            TopBorder topBorder3 = new TopBorder() { Style = BorderStyleValues.Thin };//實線
            Color topBorder3Color = new Color() { Rgb = new HexBinaryValue("4F81BD") };
            topBorder3.Append(topBorder3Color);
            BottomBorder bottomBorder3 = new BottomBorder() { Style = BorderStyleValues.Double };//雙實線
            Color bottomBorder3Color = new Color() { Rgb = new HexBinaryValue("4F81BD") };
            bottomBorder3.Append(bottomBorder3Color);
            DiagonalBorder diagonalBorder3 = new DiagonalBorder();

            border3.Append(leftBorder3);
            border3.Append(rightBorder3);
            border3.Append(topBorder3);
            border3.Append(bottomBorder3);
            border3.Append(diagonalBorder3);

            borders1.Append(border3);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)4U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFill = true };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyNumberFormat = true };
            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyFill = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);
            cellFormats1.Append(cellFormat5);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleMedium9" };

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            return stylesheet1;
        }
        /// <summary>
        /// 日期測試
        /// </summary>
        /// <param name="filename"></param>
        public void Datetest(string filename)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                //fluff to generate the workbook etc
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet" };
                sheets.Append(sheet);

                workbookPart.Workbook.Save();

                var sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                //add the style
                Stylesheet styleSheet = new Stylesheet();

                CellFormat cf = new CellFormat();
                //數字格式
                cf.NumberFormatId = 14;
                //數字格式啟用
                cf.ApplyNumberFormat = true;

                CellFormats cfs = new CellFormats();
                cfs.Append(cf);
                styleSheet.CellFormats = cfs;

                styleSheet.Borders = new Borders();
                styleSheet.Borders.Append(new Border());
                styleSheet.Fills = new Fills();
                styleSheet.Fills.Append(new Fill());
                styleSheet.Fonts = new Fonts();
                styleSheet.Fonts.Append(new Font());

                workbookPart.AddNewPart<WorkbookStylesPart>();
                workbookPart.WorkbookStylesPart.Stylesheet = styleSheet;

                CellStyles css = new CellStyles();
                CellStyle cs = new CellStyle();
                cs.FormatId = 0;
                cs.BuiltinId = 0;
                css.Append(cs);
                css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);
                styleSheet.Append(css);

                // Create custom widths for columns
                Columns lstColumns = worksheetPart.Worksheet.GetFirstChild<Columns>();
                Boolean needToInsertColumns = false;
                if (lstColumns == null)
                {
                    lstColumns = new Columns();
                    needToInsertColumns = true;
                }
                // Min = 1, Max = 1 ==> Apply this to column 1 (A)
                // Min = 2, Max = 2 ==> Apply this to column 2 (B)
                // Width = 25 ==> Set the width to 25
                // CustomWidth = true ==> Tell Excel to use the custom width
                lstColumns.Append(new Column() { Min = 1, Max = 1, Width = 6.38, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 2, Max = 2, Width = 6.38, CustomWidth = true });
                lstColumns.Append(new Column() { Min = 3, Max = 3, Width = 6.38, CustomWidth = true });
                //lstColumns.Append(new Column() { Min = 4, Max = 4, Width = 8.38, CustomWidth = true });
                //lstColumns.Append(new Column() { Min = 5, Max = 5, Width = 13, CustomWidth = true });
                //lstColumns.Append(new Column() { Min = 6, Max = 6, Width = 17, CustomWidth = true });
                //lstColumns.Append(new Column() { Min = 7, Max = 7, Width = 12, CustomWidth = true });
                // Only insert the columns if we had to create a new columns element
                if (needToInsertColumns)
                    worksheetPart.Worksheet.InsertAt(lstColumns, 0);

                Row row = new Row();

                DateTime date = new DateTime(2017, 6, 24);

                /*** Date code here ***/
                //write an OADate with type of Number
                Cell cell1 = new Cell();
                cell1.CellReference = "A1";
                cell1.CellValue = new CellValue(date.ToOADate().ToString());
                cell1.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell1.StyleIndex = 0;
                row.Append(cell1);

                //write an OADate with no type (defaults to Number)
                Cell cell2 = new Cell();
                cell2.CellReference = "B1";
                cell2.CellValue = new CellValue(date.ToOADate().ToString());
                cell1.StyleIndex = 0;
                row.Append(cell2);

                //write an ISO 8601 date with type of Date
                Cell cell3 = new Cell();
                cell3.CellReference = "C1";
                cell3.CellValue = new CellValue(date.ToString("yyyy-MM-dd"));
                cell3.DataType = new EnumValue<CellValues>(CellValues.Date);
                cell1.StyleIndex = 0;
                row.Append(cell3);

                sheetData.AppendChild(row);

                worksheetPart.Worksheet.Save();
            }
        }
        /// <summary>
        /// 有錯誤，順序問題
        /// </summary>
        /// <param name="SheetName"></param>
        public void AutoSize(string SheetName)
        {
            WorksheetPart worksheetpart = GetWorksheetPart(SheetName);
            if (worksheetpart == null) return;

            Worksheet worksheet = worksheetpart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            var maxColWidth = GetMaxCharacterWidth(sheetData);

            Columns lstColumns = worksheet.GetFirstChild<Columns>();

            int colIndex = 0;
            foreach (Column col in lstColumns)
            {
                col.Width = maxColWidth[colIndex++];
            }

            worksheet.Save();
        }
        private Dictionary<int, double> GetMaxCharacterWidth(SheetData sheetData)
        {
            //iterate over all cells getting a max char value for each column
            Dictionary<int, double> maxColWidth = new Dictionary<int, double>();
            var rows = sheetData.Elements<Row>();
            UInt32[] numberStyles = new UInt32[] { 5, 6, 7, 8 }; //styles that will add extra chars
            UInt32[] boldStyles = new UInt32[] { 1, 2, 3, 4, 6, 7, 8 }; //styles that will bold
            foreach (var r in rows)
            {
                var cells = r.Elements<Cell>().ToArray();

                //using cell index as my column
                for (int i = 0; i < cells.Length; i++)
                {
                    var cell = cells[i];
                    var cellValue = cell.CellValue == null ? string.Empty : cell.CellValue.InnerText;
                    var cellTextLength = 8.38;

                    if (cell.DataType == null || cell.DataType.Value == CellValues.Date)
                    {
                        cellTextLength = 8.88;
                    }
                    else if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                    {
                        SharedStringTablePart sharedStringTablePart = _document.WorkbookPart.SharedStringTablePart;
                        string str = sharedStringTablePart.SharedStringTable.ChildElements[int.Parse(cellValue)].InnerText;
                        cellTextLength = GetWidth(new System.Drawing.Font("微軟正黑體", 12), str);
                    }
                    else
                    {
                        cellTextLength = 8.38;
                    }
                    //if (cell.StyleIndex != null && numberStyles.Contains(cell.StyleIndex))
                    //{
                    //    int thousandCount = (int)Math.Truncate((double)cellTextLength / 4);

                    //    //add 3 for '.00' 
                    //    cellTextLength += (3 + thousandCount);
                    //}

                    //if (cell.StyleIndex != null && boldStyles.Contains(cell.StyleIndex))
                    //{
                    //    //add an extra char for bold - not 100% acurate but good enough for what i need.
                    //    cellTextLength += 1;
                    //}

                    if (maxColWidth.ContainsKey(i))
                    {
                        var current = maxColWidth[i];
                        if (cellTextLength > current)
                        {
                            maxColWidth[i] = cellTextLength;
                        }
                    }
                    else
                    {
                        maxColWidth.Add(i, cellTextLength);
                    }
                }
            }

            return maxColWidth;
        }
        private static double GetWidth(System.Drawing.Font stringFont, string text)
        {
            // This formula calculates width. For better desired outputs try to change 0.5M to something else

            System.Drawing.Size textSize = System.Windows.Forms.TextRenderer.MeasureText(text, stringFont);
            double width = (double)(((textSize.Width / (double)7) * 256) - (128 / 7)) / 256;
            width = (double)decimal.Round((decimal)width + 0.5M, 2);

            return width;
        }
        /// <summary>
        /// 加入圖片
        /// </summary>
        /// <param name="SheetName"></param>
        /// <param name="imageStream"></param>
        /// <param name="imgDesc"></param>
        /// <param name="colNumber"></param>
        /// <param name="rowNumber"></param>
        /// <param name="imageWidth">圖片寬度(公分)</param>
        /// <param name="imageHeight">圖片高度(公分)</param>
        /// <param name="imageColOffset">圖片列偏移(公分)</param>
        /// <param name="imageRowOffset">圖片行偏移(公分)</param>
        public void AddImage(string SheetName,
                             Stream imageStream, string imgDesc,
                             int colNumber, int rowNumber,
                             double imageWidth, double imageHeight,
                             double imageColOffset, double imageRowOffset)
        {
            WorksheetPart worksheetPart = GetWorksheetPart(SheetName);
            if (worksheetPart == null) return;

            // We need the image stream more than once, thus we create a memory copy
            MemoryStream imageMemStream = new MemoryStream();
            imageStream.Position = 0;
            imageStream.CopyTo(imageMemStream);
            imageStream.Position = 0;

            var drawingsPart = worksheetPart.DrawingsPart;
            if (drawingsPart == null)
                drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

            if (!worksheetPart.Worksheet.ChildElements.OfType<Drawing>().Any())
            {
                worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            if (drawingsPart.WorksheetDrawing == null)
            {
                drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
            }

            var worksheetDrawing = drawingsPart.WorksheetDrawing;

            System.Drawing.Bitmap bm = new System.Drawing.Bitmap(imageMemStream);
            var imagePart = drawingsPart.AddImagePart(GetImagePartTypeByBitmap(bm));
            imagePart.FeedData(imageStream);

            //設定圖片高度、寬度
            A.Extents extents = new A.Extents();
            //var extentsCx = bm.Width * (long)(914400 / bm.HorizontalResolution);
            //var extentsCy = bm.Height * (long)(914400 / bm.VerticalResolution);
            //公分轉英吋 1 cm = 0.393701 inch
            //英吋 * 914400 單位emu
            var extentsCx = (long)(imageWidth * 0.393701 * 914400);
            var extentsCy = (long)(imageHeight * 0.393701 * 914400);
            bm.Dispose();

            var colOffset = (long)(imageColOffset * 0.393701 * 914400);
            var rowOffset = (long)(imageRowOffset * 0.393701 * 914400);

            var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
            var nvpId = nvps.Count() > 0
                ? (UInt32Value)worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().Max(p => p.Id.Value) + 1
                : 1U;

            var oneCellAnchor = new Xdr.OneCellAnchor(
                new Xdr.FromMarker
                {
                    ColumnId = new Xdr.ColumnId((colNumber - 1).ToString()),
                    RowId = new Xdr.RowId((rowNumber - 1).ToString()),
                    ColumnOffset = new Xdr.ColumnOffset(colOffset.ToString()),
                    RowOffset = new Xdr.RowOffset(rowOffset.ToString())
                },
                new Xdr.Extent { Cx = extentsCx, Cy = extentsCy },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = "Picture " + nvpId, Description = imgDesc },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = false })
                    ),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = A.BlipCompressionValues.Print },
                        new A.Stretch(new A.FillRectangle())
                    ),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0, Y = 0 },
                            new A.Extents { Cx = extentsCx, Cy = extentsCy }
                        ),
                        new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }
                    )
                ),
                new Xdr.ClientData()
            );

            worksheetDrawing.Append(oneCellAnchor);
        }
        public static ImagePartType GetImagePartTypeByBitmap(System.Drawing.Bitmap image)
        {
            if (ImageFormat.Bmp.Equals(image.RawFormat))
                return ImagePartType.Bmp;
            else if (ImageFormat.Gif.Equals(image.RawFormat))
                return ImagePartType.Gif;
            else if (ImageFormat.Png.Equals(image.RawFormat))
                return ImagePartType.Png;
            else if (ImageFormat.Tiff.Equals(image.RawFormat))
                return ImagePartType.Tiff;
            else if (ImageFormat.Icon.Equals(image.RawFormat))
                return ImagePartType.Icon;
            else if (ImageFormat.Jpeg.Equals(image.RawFormat))
                return ImagePartType.Jpeg;
            else if (ImageFormat.Emf.Equals(image.RawFormat))
                return ImagePartType.Emf;
            else if (ImageFormat.Wmf.Equals(image.RawFormat))
                return ImagePartType.Wmf;
            else
                throw new Exception("Image type could not be determined.");
        }
        private static Stylesheet CreateStylesheet2()
        {
            Stylesheet ss = new Stylesheet();

            Fonts fts = new Fonts();

            //字型 Font 0
            DocumentFormat.OpenXml.Spreadsheet.Font ft = new DocumentFormat.OpenXml.Spreadsheet.Font();
            FontName ftn = new FontName();
            ftn.Val = StringValue.FromString("微軟正黑體");
            FontSize ftsz = new FontSize();
            ftsz.Val = DoubleValue.FromDouble(12);
            FontScheme ftsh = new FontScheme();
            ftsh.Val = FontSchemeValues.None;
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            ft.FontScheme = ftsh;
            fts.Append(ft);

            //Fonts Count
            fts.Count = UInt32Value.FromUInt32((uint)fts.ChildElements.Count);

            Fills fills = new Fills();
            //填滿 Fill 0
            Fill fill;
            PatternFill patternFill;
            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.None;
            fill.PatternFill = patternFill;
            fills.Append(fill);

            //填滿 Fill 1
            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Gray125;
            fill.PatternFill = patternFill;
            fills.Append(fill);

            //填滿 Fill 2 自訂藍色
            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Solid;
            patternFill.ForegroundColor = new ForegroundColor();
            patternFill.ForegroundColor.Rgb = HexBinaryValue.FromString("DCE6F1");
            patternFill.BackgroundColor = new BackgroundColor();
            patternFill.BackgroundColor.Rgb = patternFill.ForegroundColor.Rgb;
            fill.PatternFill = patternFill;
            fills.Append(fill);

            fills.Count = UInt32Value.FromUInt32((uint)fills.ChildElements.Count);

            Borders borders = new Borders();
            //框線 Border 0
            Border border = new Border();
            border.LeftBorder = new LeftBorder();
            border.RightBorder = new RightBorder();
            border.TopBorder = new TopBorder();
            border.BottomBorder = new BottomBorder();
            border.DiagonalBorder = new DiagonalBorder();
            borders.Append(border);

            //框線 Border 1 上下左右實線 黑色
            border = new Border();
            border.LeftBorder = new LeftBorder();
            border.LeftBorder.Style = BorderStyleValues.Thin;
            border.RightBorder = new RightBorder();
            border.RightBorder.Style = BorderStyleValues.Thin;
            border.TopBorder = new TopBorder();
            border.TopBorder.Style = BorderStyleValues.Thin;
            border.BottomBorder = new BottomBorder();
            border.BottomBorder.Style = BorderStyleValues.Thin;
            border.DiagonalBorder = new DiagonalBorder();
            borders.Append(border);
            borders.Count = UInt32Value.FromUInt32((uint)borders.ChildElements.Count);

            CellStyleFormats csfs = new CellStyleFormats();
            CellFormat cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            csfs.Append(cf);
            csfs.Count = UInt32Value.FromUInt32((uint)csfs.ChildElements.Count);

            NumberingFormats nfs = new NumberingFormats();
            CellFormats cfs = new CellFormats();
            // index 0
            cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cfs.Append(cf);

            //日期
            NumberingFormat nfDateTime = new NumberingFormat();
            nfDateTime.NumberFormatId = UInt32Value.FromUInt32(14);
            nfDateTime.FormatCode = StringValue.FromString("yyyy/mm/dd");
            nfs.Append(nfDateTime);

            //數字 小數點4位
            NumberingFormat nf4decimal = new NumberingFormat();
            nf4decimal.NumberFormatId = UInt32Value.FromUInt32(4);
            nf4decimal.FormatCode = StringValue.FromString("#,##0.0000");
            nfs.Append(nf4decimal);

            //數字 小數點2位 #,##0.00 is also Excel style index 4
            NumberingFormat nf2decimal = new NumberingFormat();
            nf2decimal.NumberFormatId = UInt32Value.FromUInt32(4);
            nf2decimal.FormatCode = StringValue.FromString("#,##0.00");
            nfs.Append(nf2decimal);

            //通用格式 @ is also Excel style index 49
            NumberingFormat nfForcedText = new NumberingFormat();
            nfForcedText.NumberFormatId = UInt32Value.FromUInt32(49);
            nfForcedText.FormatCode = StringValue.FromString("@");
            nfs.Append(nfForcedText);

            Alignment alignment1 = new Alignment();
            alignment1.Horizontal = HorizontalAlignmentValues.Center;//水平置中
            alignment1.Vertical = VerticalAlignmentValues.Center;//垂直置中

            Alignment alignment2 = new Alignment();
            alignment2.Horizontal = HorizontalAlignmentValues.Center;//水平置中
            alignment2.Vertical = VerticalAlignmentValues.Center;//垂直置中

            Alignment alignment3 = new Alignment();
            alignment3.Horizontal = HorizontalAlignmentValues.Center;//水平置中
            alignment3.Vertical = VerticalAlignmentValues.Center;//垂直置中

            // index 1 文字 藍色底黑框 
            cf = new CellFormat();
            cf.Alignment = alignment1;
            cf.NumberFormatId = nfForcedText.NumberFormatId;
            cf.FontId = 0;
            cf.FillId = 2;
            cf.BorderId = 1;
            cf.FormatId = 0;
            cf.ApplyFill = BooleanValue.FromBoolean(true);
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cfs.Append(cf);

            // index 2 文字 白色底黑框
            cf = new CellFormat();
            cf.Alignment = alignment2;
            cf.NumberFormatId = nfForcedText.NumberFormatId;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 1;
            cf.FormatId = 0;
            cf.ApplyFill = BooleanValue.FromBoolean(true);
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cfs.Append(cf);

            // index 3 日期 白色底黑框
            cf = new CellFormat();
            cf.Alignment = alignment3;
            cf.NumberFormatId = nfDateTime.NumberFormatId;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 1;
            cf.FormatId = 0;
            cf.ApplyFill = BooleanValue.FromBoolean(true);
            cf.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cfs.Append(cf);

            nfs.Count = UInt32Value.FromUInt32((uint)nfs.ChildElements.Count);
            cfs.Count = UInt32Value.FromUInt32((uint)cfs.ChildElements.Count);

            ss.Append(nfs);
            ss.Append(fts);
            ss.Append(fills);
            ss.Append(borders);
            ss.Append(csfs);
            ss.Append(cfs);

            CellStyles css = new CellStyles();
            CellStyle cs = new CellStyle();
            cs.Name = StringValue.FromString("Normal");
            cs.FormatId = 0;
            cs.BuiltinId = 0;
            css.Append(cs);
            css.Count = UInt32Value.FromUInt32((uint)css.ChildElements.Count);
            ss.Append(css);

            DifferentialFormats dfs = new DifferentialFormats();
            dfs.Count = 0;
            ss.Append(dfs);

            TableStyles tss = new TableStyles();
            tss.Count = 0;
            tss.DefaultTableStyle = StringValue.FromString("TableStyleMedium9");
            tss.DefaultPivotStyle = StringValue.FromString("PivotStyleLight16");
            ss.Append(tss);

            return ss;
        }
    }
    public enum CellReferencePartEnum
    {
        None,
        Column,
        Row,
        Both
    }
    /// <summary>
    /// 自訂格式
    /// </summary>
    public enum CustomizeStyle
    {
        藍底黑框 = 1,
        白底黑框 = 2,
        上下藍框 = 3
    }
}
