using System.Collections.Generic;
using dotnet_read_excel_example.Service;
using dotnet_read_excel_example.Service.Model;

namespace dotnet_read_excel_example.Service.Csv
{
    public class CsvService: IService
    {
        public IEnumerable<TablesModel> Read(string filepath, bool debug = false) {
            System.Console.WriteLine("csv");
            return null;
        }
    }
}