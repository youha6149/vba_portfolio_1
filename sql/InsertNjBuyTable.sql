SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE InsertIntoNjBuyTable
    @date DATE ,
    @au DECIMAL(10,2),
    @ag DECIMAL(10,2),
    @pt DECIMAL(10,2),
    @pd DECIMAL(10,2)
AS
BEGIN
    -- date�J���������j�[�N�L�[�Ƃ��Ďg�p���āA�f�[�^�����݂��邩�m�F
    IF NOT EXISTS (SELECT 1 FROM dbo.nj_buy_table WHERE date = @date)
    BEGIN
        INSERT INTO dbo.nj_buy_table
        (date, au, ag, pt, pd)
        VALUES 
        (@date, @au, @ag, @pt, @pd)
    END
END
