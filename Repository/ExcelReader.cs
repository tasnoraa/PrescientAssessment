using Repository.Model;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Globalization;

namespace Repository
{
    public class ExcelReader: IExcelReader
    {
        public System.Data.DataTable ReadExcel(string fileName)
        {
            string conn = string.Empty;
            System.Data.DataTable dtexcel = new System.Data.DataTable();
            conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
            }
            GetExcelFiles(dtexcel);
            return dtexcel;
        }

        public List<Excel> GetExcelFiles(System.Data.DataTable dataTable) 
        {
            CultureInfo provider = CultureInfo.InvariantCulture;
            var excelList = new List<Excel>();
            var summaryForDate = new DateTime(2015, 12, 31);

            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                var excel = new Excel();
                if (i > 4)
                {
                    var premimiumOption = 
                    excel.FileDate = summaryForDate;
                    excel.Contract = dataTable.Rows[i].ItemArray[0].ToString();
                    excel.ExpiryDate = DateTime.Now;//DateTime.ParseExact(dataTable.Rows[i].ItemArray[2].ToString(), "mm/dd/yyyy", provider); ;
                    excel.Classification = dataTable.Rows[i].ItemArray[3].ToString();
                    excel.Strike = dataTable.Rows[i].ItemArray[4] == DBNull.Value ? 0 : Convert.ToSingle(dataTable.Rows[i].ItemArray[4]);
                    excel.CallPut = dataTable.Rows[i].ItemArray[5].ToString();
                    excel.MTMYield = dataTable.Rows[i].ItemArray[6] == DBNull.Value ? 0 : Convert.ToSingle(dataTable.Rows[i].ItemArray[6]);
                    excel.MarkPrice = dataTable.Rows[i].ItemArray[7] == DBNull.Value ? 0 : Convert.ToSingle(dataTable.Rows[i].ItemArray[7]);
                    excel.SpotRate = dataTable.Rows[i].ItemArray[8] == DBNull.Value ? 0 : Convert.ToSingle(dataTable.Rows[i].ItemArray[8]);
                    excel.PreviousMTM = dataTable.Rows[i].ItemArray[9] == DBNull.Value ? 0 : Convert.ToSingle(dataTable.Rows[i].ItemArray[9]);
                    excel.PreviousPrice = dataTable.Rows[i].ItemArray[10] == DBNull.Value ? 0 : Convert.ToSingle(dataTable.Rows[i].ItemArray[10]);
                    excel.PremiumOnOption = dataTable.Rows[i].ItemArray[11] == DBNull.Value ? 0: Convert.ToSingle(dataTable.Rows[i].ItemArray[11]);
                    excel.Volatility = dataTable.Rows[i].ItemArray[12] == DBNull.Value ? 0 : Convert.ToSingle(dataTable.Rows[i].ItemArray[12]);
                    excel.Delta = dataTable.Rows[i].ItemArray[13] == DBNull.Value ? 0 : Convert.ToSingle(dataTable.Rows[i].ItemArray[13]);
                    excel.DeltaValue = dataTable.Rows[i].ItemArray[14] == DBNull.Value ? 0 : Convert.ToSingle(dataTable.Rows[i].ItemArray[14]);
                    excel.ContractsTraded = dataTable.Rows[i].ItemArray[15] == DBNull.Value ? 0 : Convert.ToSingle(dataTable.Rows[i].ItemArray[15]);
                    excel.OpenInterest = dataTable.Rows[i].ItemArray[16] == DBNull.Value ? 0 : Convert.ToSingle(dataTable.Rows[i].ItemArray[16]);

                    excelList.Add(excel);
                }
            }

            return excelList;
        }

    }
}
