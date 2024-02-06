USE AdventureWorksDW2022
GO

CREATE PROCEDURE dbo.DimProduct_SP
/*
SET STATISTICS TIME, IO ON

EXEC dbo.DimProduct_SP
*/
as
BEGIN

	SELECT
		p.[ProductKey] as [IdProducto],
		p.[EnglishProductName],
		ceiling(p.[StandardCost]) as [StandardCost],
		p.[ProductLine],
		ceiling(p.[DealerPrice]) as [DealerPrice],
		p.[Class],
		SUBSTRING(p.[ModelName],0,CHARINDEX('-',p.[ModelName])) as [ModelName],
		p.[EnglishDescription],
		s.[SpanishProductSubcategoryName],
		left(upper(c.[SpanishProductCategoryName]),3) as [SpanishProductCategoryName]
	FROM [dbo].[DimProduct] p
		left outer join [dbo].[DimProductSubcategory] s on (p.[ProductSubcategoryKey] = s.[ProductSubcategoryKey])
		left outer join [dbo].[DimProductCategory] as c on (s.[ProductCategoryKey] = c.[ProductCategoryKey])
	where (p.[Class] <> 'L' or p.[Class] is null) 
END