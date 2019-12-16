using dotnet_read_excel_example.Service.Execl;

namespace read_execl
{ 
    class Program
    {
        static void Main(string[] args)
        {
            var e = new ExeclService();
            e.Read(@"test.xlsx");
        }
    }
}
