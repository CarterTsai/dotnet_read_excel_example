using System.Collections.Generic;
using dotnet_read_excel_example.Service.Model;

namespace dotnet_read_excel_example.Service
{
    public interface IService
    {
        IEnumerable<TablesModel> Read(string filepath, bool debug = false);
    }
}