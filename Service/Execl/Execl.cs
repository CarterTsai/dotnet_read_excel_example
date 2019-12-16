using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using dotnet_read_excel_example.Service.Model;

namespace dotnet_read_excel_example.Service.Execl
{
    public class ExeclService: IService
    {
        public void Read(string filepath) {
            var data = ReadModelList(filepath);
            foreach (var item in data)
            {
                Console.WriteLine($"{item.Key}\t{item.Column}\t{item.DataType}\t{item.IsNull}\t{item.ColumnName}\t{item.DefaultValue}\t{item.Comment}");
            }
        }

        public static IEnumerable<TablesModel> ReadModelList(string filepath)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filepath, false))
            {
                WorkbookPart wbPart = spreadsheetDocument.WorkbookPart;
                var theSheets = wbPart.Workbook;
                var tableData = new List<TablesModel>{};

                var _sheet = new Dictionary<string, string>(); // sheet id, sheet name

                foreach (Sheet item in theSheets.Sheets)
                {
                    System.Console.WriteLine($"Sheet Name: {item.Name}");

                    IEnumerable<Sheet> _sheets =
                    wbPart.Workbook.GetFirstChild<Sheets>().
                    Elements<Sheet>().Where(s => s.Name == item.Name);

                    if (_sheets?.Count() == 0)
                    {
                        //return null;
                    }

                    string relationshipId = _sheets?.First().Id.Value;

                    WorksheetPart parts = (WorksheetPart)wbPart.GetPartById(relationshipId);

                    var _parts = new List<WorksheetPart>();
                    _parts.Add(parts);

                    foreach (WorksheetPart WSP in _parts)
                    {
                        //find sheet data
                        IEnumerable<SheetData> sheetData = WSP.Worksheet.Elements<SheetData>();

                        // Iterate through every sheet inside Excel sheet
                        foreach (SheetData SD in sheetData)
                        {
                            IEnumerable<Row> row = SD.Elements<Row>(); // Get the row IEnumerator
                            var rowData = row.Where(o => !string.IsNullOrWhiteSpace(o.InnerText));
                            foreach (var r in rowData)
                            {
                                var _cell = r.Descendants<Cell>()
                                         .Select(o => GetCellText(o, wbPart.SharedStringTablePart.SharedStringTable)).ToList();
                                var t = new TablesModel{
                                    Key = (_cell[1] != null)?_cell[1].Trim():"",
                                    Column = (_cell[2] != null)?_cell[2].Trim():"",
                                    DataType = (_cell[3] != null)?_cell[3].Trim():"",
                                    IsNull  = (_cell[4] != null)?_cell[4].Trim():"",
                                    ColumnName = (_cell[5] != null)?_cell[5].Trim():"",
                                    DefaultValue = (_cell[6] != null)?_cell[6].Trim():"",
                                    Comment = (_cell[7] != null)?_cell[7].Trim():"",
                                };

                                tableData.Add(t);
                                //Console.WriteLine(t);
                                //Console.WriteLine(String.Join("\t", _cell));
                            }
                        }
                    }
                }
                return tableData;
            }
        }

        //https://blog.darkthread.net/blog/open-xml-sdk-read-excel-string/
        public static string GetCellText(Cell cell, SharedStringTable strTable)
        {
            if (cell.ChildElements.Count == 0)
                return null;
            string val = cell.CellValue.InnerText;
            //若為共享字串時的處理邏輯
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                val = strTable.ChildElements[int.Parse(val)].InnerText;
            return val;
        }
    }
}