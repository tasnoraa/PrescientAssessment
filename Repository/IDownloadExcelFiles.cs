using System.Collections.Generic;

namespace Repository
{
    public interface IDownloadExcelFiles
    {
        void downloadExcelFiles(List<string> excelList);
    }
}
