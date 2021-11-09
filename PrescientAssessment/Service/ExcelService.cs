using Repository;
using System.Collections.Generic;

namespace PrescientAssessment.Service
{
    public class ExcelService: IExcelService
    {
        private IDownloadExcelFiles _excelFiles;
        private IExcelDocuments _documents;
        private IExcelReader _reader;
        private IExcelRepo _excelRepo;
        private List<string> documentsToDownload { get; set; }
        public ExcelService(IDownloadExcelFiles excelFiles, IExcelDocuments documents, IExcelReader reader, IExcelRepo excelRepo)
        {
            _excelFiles = excelFiles;
            _documents = documents;
            _reader = reader;
            _excelRepo = excelRepo;
            documentsToDownload = new List<string>();
        }
        public List<string> GetDownloadedList() 
        {
            documentsToDownload = _documents.GetExcelFileNames();
            _excelFiles.downloadExcelFiles(documentsToDownload);
            return documentsToDownload;
        }

        public void ReadExcelFiles() 
        {
            //var documentsToDownload = new List<string>() { @"C:\Users\User\Documents\Assessment\test\20211105_D_Daily MTM Report.xls" };

            if (documentsToDownload.Count == 0 || documentsToDownload == null)
            {
                documentsToDownload = _documents.GetExcelFileNames();
            }

            foreach (var document in documentsToDownload)
            {
                if (document.EndsWith("xls"))
                {
                    var excelDataTable = _reader.ReadExcel(document);
                    var excelList = _reader.GetExcelFiles(excelDataTable);
                    foreach (var excel in excelList)
                    {
                        _excelRepo.AddDataToTable(excel);
                    }
                }
            }

        }
    }
}
