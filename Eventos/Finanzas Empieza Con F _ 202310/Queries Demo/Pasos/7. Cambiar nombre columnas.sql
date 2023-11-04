USE AdventureWorksDW2022
GO
SET STATISTICS TIME, IO ON

select [$Outer].[ProductKey] as [IdProducto],
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
        select [_].[ProductKey] as [ProductKey],
            [_].[ProductSubcategoryKey] as [ProductSubcategoryKey2],
            [_].[EnglishProductName] as [EnglishProductName],
            [_].[StandardCost] as [StandardCost],
            [_].[ProductLine] as [ProductLine],
            [_].[DealerPrice] as [DealerPrice],
            [_].[Class] as [Class],
            [_].[ModelName] as [ModelName],
            [_].[EnglishDescription] as [EnglishDescription]
        from 
        (
            select [ProductKey],
                [ProductSubcategoryKey],
                [EnglishProductName],
                [StandardCost],
                [ProductLine],
                [DealerPrice],
                [Class],
                [ModelName],
                [EnglishDescription]
            from [dbo].[DimProduct] as [$Table]
        ) as [_]
        where ([_].[Class] <> 'L' or [_].[Class] is null) or [_].[Class] is null
    ) as [$Outer]
    left outer join [dbo].[DimProductSubcategory] as [$Inner] on ([$Outer].[ProductSubcategoryKey2] = [$Inner].[ProductSubcategoryKey])
) as [$Outer]
left outer join [dbo].[DimProductCategory] as [$Inner] on ([$Outer].[ProductCategoryKey2] = [$Inner].[ProductCategoryKey])