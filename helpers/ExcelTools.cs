using System;
using System.IO;
using CsvHelper;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Collections.Generic;
using System.Linq;

namespace exceltools.helpers
{
    public class ExcelTools
    {
        public ExcelTools()
        {
        }

        # region CSV TO Excel
        public DataTable csv2dt(string filepathIn)
        {
            if (!File.Exists(filepathIn))
            {
                throw new Exception("File not found!");
            }

            var table = new DataTable();
            using (var textReader = new StreamReader(filepathIn))
            {
                using (var csv = new CsvReader(textReader))
                {
                    csv.Configuration.HasHeaderRecord = true;

                    csv.Read();
                    csv.ReadHeader();
                    foreach (var header in csv.Context.HeaderRecord)
                    {
                        table.Columns.Add(header);
                    }


                    while (csv.Read())
                    {
                        var row = table.NewRow();
                        foreach (DataColumn column in table.Columns)
                        {
                            row[column.ColumnName] = csv.GetField(column.DataType, column.ColumnName);
                        }
                        table.Rows.Add(row);
                    }
                }
            }

            return table;
        }

        public NumberingFormats createCustomNumberFormats()
        {
            var NumberingFormats = new NumberingFormats();
            var nf2decimal = new NumberingFormat();
            nf2decimal.NumberFormatId = 500;
            nf2decimal.FormatCode = StringValue.FromString("dd/mm/yyyy");
            NumberingFormats.Append(nf2decimal);

            return NumberingFormats;
        }

        public CellFormats registerCellFormats()
        {
            CellFormat cellFormat;
            var cellFormats = new CellFormats();

            // 0   General
            cellFormats.AppendChild(new CellFormat());

            // 1    HEADER
            cellFormat = new CellFormat();
            cellFormat.FillId = 1;
            cellFormat.ApplyFill = true;
            cellFormats.AppendChild(cellFormat);

            // 2   0         
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = 1;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.AppendChild(cellFormat);

            // 3   0.00
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = 2;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.AppendChild(cellFormat);

            // 4    #,##0
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = 3;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.AppendChild(cellFormat);

            // 5    #,##0.00
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = 4;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.AppendChild(cellFormat);

            // 6    0 %
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = 9;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.AppendChild(cellFormat);

            // 7    0.00 %
            cellFormat = new CellFormat();
            cellFormat.NumberFormatId = 10;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.AppendChild(cellFormat);

            // 8    dd/mm/yyyy
            cellFormat = new CellFormat();
            cellFormat.FormatId = 0;
            cellFormat.NumberFormatId = 500;
            cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);
            cellFormats.AppendChild(cellFormat);

            cellFormats.Count = UInt32Value.FromUInt32((uint)cellFormats.ChildElements.Count);
            return cellFormats;
        }

        public Cell createCell(string value, converterToExcelSettings settings = null)
        {
            var cell = new Cell();

            if (settings != null)
            {
                switch (settings.Type)
                {
                    case 0:
                        cell.CellValue = new CellValue(value);
                        cell.DataType = CellValues.String;
                        cell.StyleIndex = (uint)settings.Type;
                        break;
                    case 2:
                    case 3:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                        cell.CellValue = new CellValue(value);
                        cell.DataType = CellValues.Number;
                        cell.StyleIndex = (uint)settings.Type;
                        break;
                    case 8:
                        DateTime auxDate;
                        if (DateTime.TryParse(value, out auxDate))
                        {
                            cell.CellValue = new CellValue(auxDate.ToOADate().ToString());
                            cell.DataType = CellValues.Number;
                            cell.StyleIndex = (uint)settings.Type;
                        }
                        break;
                }
            }

            if (cell.DataType == null)
            {
                cell.CellValue = new CellValue(value);
                cell.DataType = CellValues.String;
            }

            return cell;
        }

        public void csv2excel(string inFile, string outFile, converterToExcelSettings[] settings = null)
        {
            DataTable table = csv2dt(inFile);

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(outFile, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet();
                stylesPart.Stylesheet.NumberingFormats = createCustomNumberFormats();
                stylesPart.Stylesheet.CellFormats = registerCellFormats();
                stylesPart.Stylesheet.Fills = new Fills(
                        new Fill(new PatternFill() { PatternType = PatternValues.None }),
                        new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "dddddd" } }) { PatternType = PatternValues.Solid })
                    );
                stylesPart.Stylesheet.Save();

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                var sheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };

                sheets.Append(sheet);

                var headerRow = new Row();

                List<String> columns = new List<string>();
                foreach (DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);

                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(column.ColumnName);
                    cell.StyleIndex = 1;
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (DataRow dsrow in table.Rows)
                {
                    var newRow = new Row();
                    int i = 0;
                    foreach (String col in columns)
                    {
                        converterToExcelSettings settingsLocal = null;
                        if (settings != null)
                        {
                            try
                            {
                                settingsLocal = settings[i++] ?? null;
                            }
                            catch (Exception e) { }
                        }

                        var cell = createCell(dsrow[col].ToString(), settingsLocal);

                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }

                if (settings != null)
                {
                    Columns lstColumns = worksheetPart.Worksheet.GetFirstChild<Columns>();
                    Boolean needToInsertColumns = false;
                    if (lstColumns == null)
                    {
                        lstColumns = new Columns();
                        needToInsertColumns = true;
                    }


                    foreach (converterToExcelSettings item in settings)
                    {
                        if (item.Width == 0)
                        {
                            //TODO: Autofit
                        }
                        else if (item.Width > 0)
                        {
                            lstColumns.Append(new Column()
                            {
                                Min = (uint)item.Index,
                                Max = (uint)item.Index,
                                Width = item.Width,
                                CustomWidth = true
                            });
                        }
                    }

                    if (needToInsertColumns)
                    {
                        worksheetPart.Worksheet.InsertAt(lstColumns, 0);
                    }
                }

                workbookPart.Workbook.Save();
            }
        }
        #endregion

        #region Excel to CSV
        public DataTable excel2dt(string filepathIn, converterToCsvSettings settings)
        {
            if (!File.Exists(filepathIn))
            {
                throw new Exception("File not found!");
            }

            DataTable dataTable = new DataTable();
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(filepathIn, false))
            {
                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                //string relationshipId = sheets.First().Id.Value;

                Boolean runHeader = true;
                foreach (Sheet sheet in sheets)
                {
                    if (settings != null)
                    {
                        if (settings.Sheets != null && settings.Sheets is Array && settings.Sheets.Length > 0)
                        {
                            if (!settings.Sheets.Contains(sheet.Name.Value))
                            {
                                continue;
                            }
                        }
                        if (settings.SkipHidden != null && settings.SkipHidden == true)
                        {
                            if (sheet.State == null || sheet.State.Value == SheetStateValues.Visible);
                            else {
                                continue;
                            }
                        }
                    }

                    string relationshipId = sheet.Id.Value;
                    WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                    Worksheet workSheet = worksheetPart.Worksheet;
                    SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                    IEnumerable<Row> rows = sheetData.Descendants<Row>();

                    if (runHeader)
                    {
                        foreach (Cell cell in rows.ElementAt(0))
                        {
                            dataTable.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                        }
                        runHeader = false;
                    }

                    foreach (Row row in rows)
                    {
                        DataRow dataRow = dataTable.NewRow();
                        for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                        {
                            dataRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                        }

                        dataTable.Rows.Add(dataRow);
                    }
                }

            }

            if (dataTable.Rows.Count > 0)
            {
                dataTable.Rows.RemoveAt(0);
            }

            return dataTable;
        }

        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            try {
                SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
                string value = cell.CellValue.InnerXml;

                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                }
                else
                {
                    return value;
                }
            } catch (Exception e){
                return null;
            }
        }

        public void excel2csv(string inFile, string outFile, converterToCsvSettings settings = null)
        {
            DataTable table = excel2dt(inFile, settings);

            StreamWriter sw = new StreamWriter(outFile, false);
            //headers   
            for (int i = 0; i < table.Columns.Count; i++)
            {
                sw.Write(table.Columns[i].ToString().Trim());
                if (i < table.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in table.Rows)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString().Trim();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString().Trim());
                        }
                    }
                    if (i < table.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }
        #endregion
    }
}
