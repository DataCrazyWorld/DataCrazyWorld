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
Tabla "Workfile". N�mero de examen 0, lecturas l�gicas 0, lecturas f�sicas 0, lecturas de servidor de p�ginas 0, lecturas anticipadas 0, lecturas anticipadas de servidor de p�ginas 0, lecturas l�gicas de l�nea de negocio 0, lecturas f�sicas de l�nea de negocio 0, lecturas de servidor de p�ginas de l�nea de negocio 0, lecturas anticipadas de l�nea de negocio 0, lecturas anticipadas de servidor de p�ginas de l�nea de negocio 0.
Tabla "Worktable". N�mero de examen 0, lecturas l�gicas 0, lecturas f�sicas 0, lecturas de servidor de p�ginas 0, lecturas anticipadas 0, lecturas anticipadas de servidor de p�ginas 0, lecturas l�gicas de l�nea de negocio 0, lecturas f�sicas de l�nea de negocio 0, lecturas de servidor de p�ginas de l�nea de negocio 0, lecturas anticipadas de l�nea de negocio 0, lecturas anticipadas de servidor de p�ginas de l�nea de negocio 0.
Tabla "DimProduct". N�mero de examen 1, lecturas l�gicas 253, lecturas f�sicas 1, lecturas de servidor de p�ginas 0, lecturas anticipadas 251, lecturas anticipadas de servidor de p�ginas 0, lecturas l�gicas de l�nea de negocio 0, lecturas f�sicas de l�nea de negocio 0, lecturas de servidor de p�ginas de l�nea de negocio 0, lecturas anticipadas de l�nea de negocio 0, lecturas anticipadas de servidor de p�ginas de l�nea de negocio 0.
Tabla "DimProductSubcategory". N�mero de examen 1, lecturas l�gicas 2, lecturas f�sicas 1, lecturas de servidor de p�ginas 0, lecturas anticipadas 0, lecturas anticipadas de servidor de p�ginas 0, lecturas l�gicas de l�nea de negocio 0, lecturas f�sicas de l�nea de negocio 0, lecturas de servidor de p�ginas de l�nea de negocio 0, lecturas anticipadas de l�nea de negocio 0, lecturas anticipadas de servidor de p�ginas de l�nea de negocio 0.
Tabla "DimProductCategory". N�mero de examen 1, lecturas l�gicas 2, lecturas f�sicas 1, lecturas de servidor de p�ginas 0, lecturas anticipadas 0, lecturas anticipadas de servidor de p�ginas 0, lecturas l�gicas de l�nea de negocio 0, lecturas f�sicas de l�nea de negocio 0, lecturas de servidor de p�ginas de l�nea de negocio 0, lecturas anticipadas de l�nea de negocio 0, lecturas anticipadas de servidor de p�ginas de l�nea de negocio 0.

 Tiempos de ejecuci�n de SQL Server:
   Tiempo de CPU = 0 ms, tiempo transcurrido = 52 ms.

Completion time: 2023-09-22T16:00:54.3410653+02:00
*/