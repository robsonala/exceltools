using System;
using System.IO;
using CsvHelper;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Collections.Generic;

namespace exceltools.helpers
{
    public class ExcelTools
    {
        public ExcelTools()
        {
        }

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

		public Cell createCell(string value, converterSettings settings = null)
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

		public void csv2excel(string inFile, string outFile, converterSettings[] settings = null)
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
				var sheet = new Sheet() {
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
						converterSettings settingsLocal = null;
						if (settings != null)
						{
							try {
                                settingsLocal = settings[i++] ?? null;
							} catch (Exception e){}
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


					foreach (converterSettings item in settings)
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
    }
}
