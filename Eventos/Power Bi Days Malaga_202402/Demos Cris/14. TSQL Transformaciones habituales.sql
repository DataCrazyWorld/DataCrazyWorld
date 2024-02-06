USE AdventureWorksDW2022
GO

/*Quitar espacios por delante y por detrás*/

SELECT EnglishPromotionName , RTRIM(LTRIM(EnglishPromotionName))  FROM [dbo].[DimPromotion]

/*Reemplazar*/
SELECT EnglishPromotionName , REPLACE(EnglishPromotionName, 'Volume','Vol')  FROM [dbo].[DimPromotion]

/*Quitar duplicados*/
SELECT DepartmentName from dbo.DimEmployee
SELECT DISTINCT DepartmentName from dbo.DimEmployee

/*
Redondear
*/

SELECT [StandardCost], CEILING(StandardCost) as AlAza from dbo.DimProduct where StandardCost is not null

/* 
Pivotar columnas
*/

USE AdventureWorksDW2022
go

-- Origen
select Title, DepartmentName, COUNT(EmployeeKey) as #Employee 
FROM DBO.DimEmployee Group by Title, DepartmentName

-- Pivontando..
DECLARE @sql nvarchar(MAX);
 
 SET @sql = N'
 
 SELECT
   * 
  FROM
  (  
    SELECT Title
         , DepartmentName
         , EmployeeKey
 FROM [dbo].[DimEmployee]
  ) AS T
  PIVOT   
  (
  COUNT(EmployeeKey)
  FOR DepartmentName IN (' + (SELECT STUFF(
 (
 SELECT
   ',' + QUOTENAME(LTRIM(DepartmentName))
 FROM
   (SELECT DISTINCT DepartmentName
    FROM [dbo].[DimEmployee]
   ) AS T
 ORDER BY
 DepartmentName
 FOR XML PATH('')
 ), 1, 1, '')) + N')
  ) AS P;'; 
 
--En la variable @sql tenemos la consulta completa 
 
 EXEC sp_executesql @sql;

/* 
DesPivotar columnas
*/
SELECT ProductKey, [EnglishProductName],[SpanishProductName],[FrenchProductName]
FROM [dbo].[DimProduct]

SELECT ProductKey, [Language], ProductName 
FROM   
   (SELECT ProductKey, [EnglishProductName] as English,[SpanishProductName]as Spanish,[FrenchProductName] as French
   FROM [dbo].[DimProduct]) p  
UNPIVOT  
   (ProductName FOR [Language] IN   
      (English,Spanish,French)  
)AS unpvt;  
GO 

/*
FIRST_VALUE y LAST_VALUE
*/
SELECT * from  dbo.Ejemplo
 
-- Queremos sacar de cada ID su primer y ultimo importe y fecha
-- El instinto te puede decir que FIRST_VALUE y LAST_VALUE... pero ¡ojo! ¡NO! Mirad lo que hace...
Select DISTINCT
    id
    ,FIRST_VALUE(importe) OVER(PARTITION BY ID ORDER BY Fecha) as PrimerImporte
    ,FIRST_VALUE(fecha) OVER(PARTITION BY ID ORDER BY Fecha) as FechaPrimerImporte
    ,LAST_VALUE(importe) OVER(PARTITION BY ID ORDER BY Fecha) as UltimoImporte
    ,LAST_VALUE(fecha) OVER(PARTITION BY ID ORDER BY Fecha) as FechaUltimoImporte
from dbo.Ejemplo
 
-- Es mejor usar FIRST_VALUE y ordenar a la inversa para coger el último
Select DISTINCT
    id
    ,FIRST_VALUE(importe) OVER(PARTITION BY ID ORDER BY Fecha) as PrimerImporte
    ,FIRST_VALUE(fecha) OVER(PARTITION BY ID ORDER BY Fecha) as FechaPrimerImporte
    ,FIRST_VALUE(importe) OVER(PARTITION BY ID ORDER BY Fecha DESC) as UltimoImporte
    ,FIRST_VALUE(fecha) OVER(PARTITION BY ID ORDER BY Fecha DESC) as FechaUltimoImporte
from dbo.Ejemplo


