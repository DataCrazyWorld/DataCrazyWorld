USE AdventureWorksDW2022
GO
SET STATISTICS TIME, IO ON

SELECT
	p.[ProductKey] as [IdProducto],
    p.[EnglishProductName],
    ceiling(p.[StandardCost]) as [StandardCost],
    p.[ProductLine],
    ceiling(p.[DealerPrice]) as [DealerPrice],
    p.[Class],
    p.[ModelName],
    p.[EnglishDescription],
	s.[SpanishProductSubcategoryName],
	left(upper(c.[SpanishProductCategoryName]),3) as [SpanishProductCategoryName]
FROM [dbo].[DimProduct] p
	left outer join [dbo].[DimProductSubcategory] s on (p.[ProductSubcategoryKey] = s.[ProductSubcategoryKey])
	left outer join [dbo].[DimProductCategory] as c on (s.[ProductCategoryKey] = c.[ProductCategoryKey])
where (p.[Class] <> 'L' or p.[Class] is null) 