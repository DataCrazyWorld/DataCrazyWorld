
SET STATISTICS TIME, IO ON

SELECT * FROM [Contoso 100K].Data.Orders
/*
 Tiempos de ejecución de SQL Server:
   Tiempo de CPU = 0 ms, tiempo transcurrido = 576 ms.
*/
SELECT * FROM [Contoso 100M].Data.Orders
/*
 Tiempos de ejecución de SQL Server:
   Tiempo de CPU = 7515 ms, tiempo transcurrido = 525269 ms.
*/

SELECT Customerkey FROM [Contoso 100K].Data.Orders

/*
 Tiempos de ejecución de SQL Server:
   Tiempo de CPU = 0 ms, tiempo transcurrido = 244 ms.
*/

SELECT Customerkey FROM [Contoso 100M].Data.Orders

/*
 Tiempos de ejecución de SQL Server:
   Tiempo de CPU = 2812 ms, tiempo transcurrido = 285978 ms.
*/

--------------------------------------------------------------------------------------------------------------------------
SELECT CustomerKey from [Contoso 100K].Data.Orders where [Currency Code] = 'CAD'
/*
 Tiempos de ejecución de SQL Server:
   Tiempo de CPU = 0 ms, tiempo transcurrido = 85 ms.
*/

SELECT Customerkey FROM [Contoso 100M].Data.Orders where [Currency Code] = 'CAD'
/*
 Tiempos de ejecución de SQL Server:
   Tiempo de CPU = 296 ms, tiempo transcurrido = 24530 ms.
*/
