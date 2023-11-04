USE AdventureWorksDW2022
GO
SET STATISTICS TIME, IO ON


select [$Outer].[ProductKey] as [ProductKey],
    [$Outer].[EnglishProductName] as [EnglishProductName],
    [$Outer].[StandardCost] as [StandardCost],
    [$Outer].[ProductLine] as [ProductLine],
    [$Outer].[DealerPrice] as [DealerPrice],
    [$Outer].[Class] as [Class],
    [$Outer].[ModelName] as [ModelName],
    [$Outer].[EnglishDescription] as [EnglishDescription],
    [$Inner].[SpanishProductSubcategoryName] as [DimProductSubcategory.SpanishProductSubcategoryName]
from 
(
    select [ProductKey] as [ProductKey],
        [ProductSubcategoryKey] as [ProductSubcategoryKey2],
        [EnglishProductName] as [EnglishProductName],
        [StandardCost] as [StandardCost],
        [ProductLine] as [ProductLine],
        [DealerPrice] as [DealerPrice],
        [Class] as [Class],
        [ModelName] as [ModelName],
        [EnglishDescription] as [EnglishDescription]
    from [dbo].[DimProduct] as [$Table]
) as [$Outer]
left outer join [dbo].[DimProductSubcategory] as [$Inner] on ([$Outer].[ProductSubcategoryKey2] = [$Inner].[ProductSubcategoryKey])