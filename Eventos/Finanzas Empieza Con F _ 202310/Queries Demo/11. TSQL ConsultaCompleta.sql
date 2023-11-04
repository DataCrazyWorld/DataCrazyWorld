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
    SUBSTRING(p.[ModelName],0,CHARINDEX('-',p.[ModelName])) as [ModelName],
    p.[EnglishDescription],
	s.[SpanishProductSubcategoryName],
	left(upper(c.[SpanishProductCategoryName]),3) as [SpanishProductCategoryName]
FROM [dbo].[DimProduct] p
	left outer join [dbo].[DimProductSubcategory] s on (p.[ProductSubcategoryKey] = s.[ProductSubcategoryKey])
	left outer join [dbo].[DimProductCategory] as c on (s.[ProductCategoryKey] = c.[ProductCategoryKey])
where (p.[Class] <> 'L' or p.[Class] is null) 

/*
(477 rows affected)
Tabla "Workfile". Número de examen 0, lecturas lógicas 0, lecturas físicas 0, lecturas de servidor de páginas 0, lecturas anticipadas 0, lecturas anticipadas de servidor de páginas 0, lecturas lógicas de línea de negocio 0, lecturas físicas de línea de negocio 0, lecturas de servidor de páginas de línea de negocio 0, lecturas anticipadas de línea de negocio 0, lecturas anticipadas de servidor de páginas de línea de negocio 0.
Tabla "Worktable". Número de examen 0, lecturas lógicas 0, lecturas físicas 0, lecturas de servidor de páginas 0, lecturas anticipadas 0, lecturas anticipadas de servidor de páginas 0, lecturas lógicas de línea de negocio 0, lecturas físicas de línea de negocio 0, lecturas de servidor de páginas de línea de negocio 0, lecturas anticipadas de línea de negocio 0, lecturas anticipadas de servidor de páginas de línea de negocio 0.
Tabla "DimProduct". Número de examen 1, lecturas lógicas 253, lecturas físicas 1, lecturas de servidor de páginas 0, lecturas anticipadas 251, lecturas anticipadas de servidor de páginas 0, lecturas lógicas de línea de negocio 0, lecturas físicas de línea de negocio 0, lecturas de servidor de páginas de línea de negocio 0, lecturas anticipadas de línea de negocio 0, lecturas anticipadas de servidor de páginas de línea de negocio 0.
Tabla "DimProductSubcategory". Número de examen 1, lecturas lógicas 2, lecturas físicas 1, lecturas de servidor de páginas 0, lecturas anticipadas 0, lecturas anticipadas de servidor de páginas 0, lecturas lógicas de línea de negocio 0, lecturas físicas de línea de negocio 0, lecturas de servidor de páginas de línea de negocio 0, lecturas anticipadas de línea de negocio 0, lecturas anticipadas de servidor de páginas de línea de negocio 0.
Tabla "DimProductCategory". Número de examen 1, lecturas lógicas 2, lecturas físicas 1, lecturas de servidor de páginas 0, lecturas anticipadas 0, lecturas anticipadas de servidor de páginas 0, lecturas lógicas de línea de negocio 0, lecturas físicas de línea de negocio 0, lecturas de servidor de páginas de línea de negocio 0, lecturas anticipadas de línea de negocio 0, lecturas anticipadas de servidor de páginas de línea de negocio 0.

 Tiempos de ejecución de SQL Server:
   Tiempo de CPU = 0 ms, tiempo transcurrido = 52 ms.

Completion time: 2023-09-22T16:00:54.3410653+02:00
*/