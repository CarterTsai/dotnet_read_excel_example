using System;
using dotnet_read_excel_example.Service.Execl;

namespace read_execl
{ 
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            var e = new ExeclService();
            var table = e.Read(@"test.xlsx");
        }
    }
}
