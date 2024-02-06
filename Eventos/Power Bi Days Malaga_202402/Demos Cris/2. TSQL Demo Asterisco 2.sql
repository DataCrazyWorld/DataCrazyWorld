SET STATISTICS TIME, IO ON

SELECT * FROM [AdventureWorks2022].[Sales].[SalesOrderDetail]

/*
 Tiempos de ejecución de SQL Server:
   Tiempo de CPU = 15 ms, tiempo transcurrido = 787 ms.

   Scan del cluster index
*/

SELECT SalesOrderDetailID FROM [AdventureWorks2022].[Sales].[SalesOrderDetail]

/*
 Tiempos de ejecución de SQL Server:
   Tiempo de CPU = 0 ms, tiempo transcurrido = 342 ms.
*/

SELECT * FROM [AdventureWorks2022].[Sales].[SalesOrderDetail] where ProductID = 777

/*
 Tiempos de ejecución de SQL Server:
   Tiempo de CPU = 0 ms, tiempo transcurrido = 115 ms.

   Index seek + Lookup
*/

SELECT SalesOrderDetailID FROM [AdventureWorks2022].[Sales].[SalesOrderDetail] where ProductID = 777

/*
 Tiempos de ejecución de SQL Server:
   Tiempo de CPU = 0 ms, tiempo transcurrido = 2 ms.

   Index seek.
*/