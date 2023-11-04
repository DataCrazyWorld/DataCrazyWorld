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
    [$Outer].[SpanishProductSubcategoryName] as [DimProductSubcategory.SpanishProductSubcategoryName],
    [$Inner].[SpanishProductCategoryName] as [DimProductSubcategory.DimProductCategory.SpanishProductCategoryName]
from 
(
    select [$Outer].[ProductKey] as [ProductKey],
        [$Outer].[ProductSubcategoryKey2] as [ProductSubcategoryKey2],
        [$Outer].[EnglishProductName] as [EnglishProductName],
        [$Outer].[StandardCost] as [StandardCost],
        [$Outer].[ProductLine] as [ProductLine],
        [$Outer].[DealerPrice] as [DealerPrice],
        [$Outer].[Class] as [Class],
        [$Outer].[ModelName] as [ModelName],
        [$Outer].[EnglishDescription] as [EnglishDescription],
        [$Inner].[SpanishProductSubcategoryName] as [SpanishProductSubcategoryName],
        [$Inner].[ProductCategoryKey] as [ProductCategoryKey2]
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
) as [$Outer]
left outer join [dbo].[DimProductCategory] as [$Inner] on ([$Outer].[ProductCategoryKey2] = [$Inner].[ProductCategoryKey])