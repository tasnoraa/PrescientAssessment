using System.Collections.Generic;
using System.IO;
using System.Net.Http;

namespace Repository
{
    public class DownloadExcelFiles: IDownloadExcelFiles
    {
        public async void downloadExcelFiles(List<string> excelList) 
        {
            var uri = "https://clientportal.jse.co.za/downloadable-files?RequestNode=/YieldX/Derivatives/Docs_DMTM";

            var httpClient = new HttpClient();
            foreach (var file in excelList)
            {
                if (!File.Exists(@"C:\ExcelFiles\" + file)) {
                    using (var stream = await httpClient.GetStreamAsync(uri + @"/" + file))
                    {
                        using (var fileStream = new FileStream(@"C:\ExcelFiles\" + file, FileMode.OpenOrCreate))
                        {
                            stream.CopyToAsync(fileStream).Wait();
                        }
                    } 
                }
                
            }
        }
    }
}
