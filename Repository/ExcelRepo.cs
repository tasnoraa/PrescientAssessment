using Dapper;
using Repository.Model;
using System.Data.SqlClient;

namespace Repository
{
    public class ExcelRepo: IExcelRepo
    {
        public int AddDataToTable(Excel excel) 
        {
            int affectedRows = 0;
            string sql = InsertQuery();

            using (var connection = new SqlConnection("Data Source=.;Initial Catalog=master;Integrated Security=true;"))
            {
                affectedRows = connection.Execute(sql, InsertParam(excel));
            }

            return affectedRows;
        }

        public string InsertQuery() 
        {
            return @"
            INSERT [dbo].[DailyMTM] (
                FileDate,
                Contract, 
                ExpiryDate, 
                Classification, 
                Strike, 
                CallPut, 
                MTMYield, 
                MarkPrice, 
                SpotRate, 
                PreviousMTM, 
                PreviousPrice, 
                PremiumOnOption, 
                Volatility, 
                Delta, 
                DeltaValue, 
                ContractsTraded, 
                OpenInterest)
            VALUES (@FileDate,
                @Contract, 
                @ExpiryDate, 
                @Classification, 
                @Strike, 
                @CallPut, 
                @MTMYield, 
                @MarkPrice, 
                @SpotRate, 
                @PreviousMTM, 
                @PreviousPrice, 
                @PremiumOnOption, 
                @Volatility, 
                @Delta, 
                @DeltaValue, 
                @ContractsTraded, 
                @OpenInterest)
            ";
        }

        public object InsertParam(Excel excel) 
        {
            return new 
            {
                FileDate = excel.FileDate,
                Contract = excel.Contract,
                ExpiryDate = excel.ExpiryDate,
                Classification = excel.Classification,
                Strike = excel.Strike,
                CallPut = excel.CallPut,
                MTMYield = excel.MTMYield,
                MarkPrice = excel.MarkPrice,
                SpotRate = excel.SpotRate,
                PreviousMTM = excel.PreviousMTM,
                PreviousPrice = excel.PreviousPrice,
                PremiumOnOption = excel.PremiumOnOption,
                Volatility = excel.Volatility,
                Delta = excel.Delta, 
                DeltaValue = excel.DeltaValue,
                ContractsTraded = excel.ContractsTraded,
                OpenInterest = excel.OpenInterest
            };
        }
    }
}
