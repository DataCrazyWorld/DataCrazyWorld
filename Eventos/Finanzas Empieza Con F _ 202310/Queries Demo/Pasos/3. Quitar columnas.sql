USE AdventureWorksDW2022
GO
SET STATISTICS TIME, IO ON

select [ProductKey],
    [EnglishProductName],
    [StandardCost],
    [ProductLine],
    [DealerPrice],
    [Class],
    [ModelName],
    [EnglishDescription]
from [dbo].[DimProduct] as [$Table]