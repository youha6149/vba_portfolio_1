SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE InsertIntoMarketTable
    @date DATE,
    @price_date NVARCHAR(50),
    @price_hhmm NVARCHAR(10),
    @au_ny_end DECIMAL(10,2),
    @pt_ny_end DECIMAL(10,2),
    @ag_ny_end DECIMAL(10,2),
    @ny_exchange_rate DECIMAL(10,2),
    @ny_end INT,
    @tokyo_exchange_rate DECIMAL(10,2),
    @tokyo_start INT,
    @au_buy DECIMAL(10,2),
    @au_buy_diff INT,
    @pt_buy DECIMAL(10,2),
    @pt_buy_diff INT,
    @au_sell DECIMAL(10,2),
    @au_sell_diff INT,
    @pt_sell DECIMAL(10,2),
    @pt_sell_diff INT
AS
BEGIN
    INSERT INTO dbo.market_table
    (date, price_date, price_hhmm, au_ny_end, pt_ny_end, ag_ny_end, ny_exchange_rate, ny_end, tokyo_exchange_rate, tokyo_start, au_buy, au_buy_diff, pt_buy, pt_buy_diff, au_sell, au_sell_diff, pt_sell, pt_sell_diff)
    VALUES 
    (@date, @price_date, @price_hhmm, @au_ny_end, @pt_ny_end, @ag_ny_end, @ny_exchange_rate, @ny_end, @tokyo_exchange_rate, @tokyo_start, @au_buy, @au_buy_diff, @pt_buy, @pt_buy_diff, @au_sell, @au_sell_diff, @pt_sell, @pt_sell_diff)
END
