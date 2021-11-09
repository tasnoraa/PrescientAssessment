using Repository.Model;
using System.Collections.Generic;

namespace Repository
{
    public interface IExcelReader
    {
        System.Data.DataTable ReadExcel(string fileName);

        List<Excel> GetExcelFiles(System.Data.DataTable dataTable);
    }
}
