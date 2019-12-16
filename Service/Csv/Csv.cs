using dotnet_read_excel_example.Service;

namespace dotnet_read_excel_example.Service.Csv
{
    public class CsvService: IService
    {
        public void Read(string filepath) {
            System.Console.WriteLine("csv");
        }
    }
}