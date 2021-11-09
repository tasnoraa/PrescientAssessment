using HtmlAgilityPack;
using System.Collections.Generic;
using System.Net;

namespace Repository
{
    public class ExcelDocuments: IExcelDocuments
    {
        private List<string> dontDownloadFiles = new List<string> 
        {
            "Parent","2005","2006","2007","2008","2009","2010","2011",
            "2012","2013","2014","2015","2016","2017","2018","2019",
            "2020"
        };
        public List<string> GetExcelFileNames()
        {
            var excelList = new List<string>();
            string html;
            using (WebClient client = new WebClient())
            {
                html = client.DownloadString("https://clientportal.jse.co.za/downloadable-files?RequestNode=/YieldX/Derivatives/Docs_DMTM");
            }
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(html);
            foreach (HtmlNode excel in doc.DocumentNode.SelectNodes("//a[@class='inline']"))
            {
                if(!dontDownloadFiles.Contains(excel.InnerText))
                    excelList.Add(excel.InnerHtml);
            }

            return excelList;
        }
    }
}
