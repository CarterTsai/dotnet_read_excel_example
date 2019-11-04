using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace read_execl
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadModelList();
        }

        public static void ReadModelList()
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(@"test.xlsx", false))
            {
                WorkbookPart wbPart = spreadsheetDocument.WorkbookPart;
                var theSheets = wbPart.Workbook;

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

                                Console.WriteLine(String.Join("\t", _cell));
                            }
                        }
                    }
                }
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
