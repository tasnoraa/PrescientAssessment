
USE master;  
GO  
CREATE PROCEDURE SP_Total_Contracts_Traded_Report  
    @DateFrom date,   
    @DateTo date   
AS   

    SET NOCOUNT ON;  
    SELECT [FileDate],[Contract],[ContractsTraded]  
    FROM [master].[dbo].[DailyMTM] 
    WHERE FileDate = @DateFrom AND ExpiryDate = @DateTo  AND [ContractsTraded] > 0
 
GO
