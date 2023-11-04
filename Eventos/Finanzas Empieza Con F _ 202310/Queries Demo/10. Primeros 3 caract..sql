USE AdventureWorksDW2022
GO

SET STATISTICS TIME, IO ON

select [_].[ProductKey] as [IdProducto],
    [_].[EnglishProductName] as [EnglishProductName],
    [_].[t0_0] as [StandardCost],
    [_].[ProductLine] as [ProductLine],
    [_].[t1_0] as [DealerPrice],
    [_].[Class] as [Class],
    [_].[ModelName] as [ModelName],
    [_].[EnglishDescription] as [EnglishDescription],
    [_].[SpanishProductSubcategoryName] as [DimProductSubcategory.SpanishProductSubcategoryName],
    left([_].[t0_02], 3) as [DimProductSubcategory.DimProductCategory.SpanishProductCategoryName]
from 
(
    select [_].[ProductKey] as [ProductKey],
        [_].[EnglishProductName] as [EnglishProductName],
        [_].[StandardCost] as [StandardCost],
        [_].[ProductLine] as [ProductLine],
        [_].[DealerPrice] as [DealerPrice],
        [_].[Class] as [Class],
        [_].[ModelName] as [ModelName],
        [_].[EnglishDescription] as [EnglishDescription],
        [_].[SpanishProductSubcategoryName] as [SpanishProductSubcategoryName],
        [_].[SpanishProductCategoryName] as [SpanishProductCategoryName],
        ceiling([_].[StandardCost]) as [t0_0],
        ceiling([_].[DealerPrice]) as [t1_0],
        upper([_].[SpanishProductCategoryName]) as [t0_02]
    from 
    (
        select [$Outer].[ProductKey],
            [$Outer].[EnglishProductName],
            [$Outer].[StandardCost],
            [$Outer].[ProductLine],
            [$Outer].[DealerPrice],
            [$Outer].[Class],
            [$Outer].[ModelName],
            [$Outer].[EnglishDescription],
            [$Outer].[SpanishProductSubcategoryName],
            [$Inner].[SpanishProductCategoryName]
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
    ) as [_]
) as [_]

/*

(477 rows affected)
Tabla "Workfile". Número de examen 0, lecturas lógicas 0, lecturas físicas 0, lecturas de servidor de páginas 0, lecturas anticipadas 0, lecturas anticipadas de servidor de páginas 0, lecturas lógicas de línea de negocio 0, lecturas físicas de línea de negocio 0, lecturas de servidor de páginas de línea de negocio 0, lecturas anticipadas de línea de negocio 0, lecturas anticipadas de servidor de páginas de línea de negocio 0.
Tabla "Worktable". Número de examen 0, lecturas lógicas 0, lecturas físicas 0, lecturas de servidor de páginas 0, lecturas anticipadas 0, lecturas anticipadas de servidor de páginas 0, lecturas lógicas de línea de negocio 0, lecturas físicas de línea de negocio 0, lecturas de servidor de páginas de línea de negocio 0, lecturas anticipadas de línea de negocio 0, lecturas anticipadas de servidor de páginas de línea de negocio 0.
Tabla "DimProduct". Número de examen 1, lecturas lógicas 253, lecturas físicas 1, lecturas de servidor de páginas 0, lecturas anticipadas 251, lecturas anticipadas de servidor de páginas 0, lecturas lógicas de línea de negocio 0, lecturas físicas de línea de negocio 0, lecturas de servidor de páginas de línea de negocio 0, lecturas anticipadas de línea de negocio 0, lecturas anticipadas de servidor de páginas de línea de negocio 0.
Tabla "DimProductSubcategory". Número de examen 1, lecturas lógicas 2, lecturas físicas 1, lecturas de servidor de páginas 0, lecturas anticipadas 0, lecturas anticipadas de servidor de páginas 0, lecturas lógicas de línea de negocio 0, lecturas físicas de línea de negocio 0, lecturas de servidor de páginas de línea de negocio 0, lecturas anticipadas de línea de negocio 0, lecturas anticipadas de servidor de páginas de línea de negocio 0.
Tabla "DimProductCategory". Número de examen 1, lecturas lógicas 2, lecturas físicas 1, lecturas de servidor de páginas 0, lecturas anticipadas 0, lecturas anticipadas de servidor de páginas 0, lecturas lógicas de línea de negocio 0, lecturas físicas de línea de negocio 0, lecturas de servidor de páginas de línea de negocio 0, lecturas anticipadas de línea de negocio 0, lecturas anticipadas de servidor de páginas de línea de negocio 0.

 Tiempos de ejecución de SQL Server:
   Tiempo de CPU = 0 ms, tiempo transcurrido = 86 ms.

Completion time: 2023-09-22T16:00:15.5573905+02:00
*/