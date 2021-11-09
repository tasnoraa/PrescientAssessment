using System.Collections.Generic;

namespace PrescientAssessment.Service
{
    public interface IExcelService
    {
        List<string> GetDownloadedList();

        void ReadExcelFiles();
    }
}
