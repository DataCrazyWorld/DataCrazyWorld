#Pon aquí tu url
$path= "C:\Charlas\20241026_PowerBIDaysSantiago2024_DocumentaTuBIYTriunfa\1. Demo SSIS\DimCards.dtsx"


# read .dtsx into XML variable
[xml] $myxml = Get-Content $path

#Para leer con prefijos DTS y SQLTask tengo que indicar los namespaces
$XmlNamespace = @{ DTS = 'www.microsoft.com/SqlServer/Dts';  SQLTask = 'www.microsoft.com/sqlserver/dts/tasks/sqltask'};

#Recogemos el Lookup de la línea 300
 $myxml | Select-XML -XPath "//component[@componentClassID='Microsoft.Lookup']" | ForEach-Object {
    Write-Host $_.Node.Name        
      
    #Recojo los hijos
    [xml] $childs = $_.Node.OuterXml 
          
    $childs| Select-XML -XPath "//connection[@connectionManagerRefId]"  | ForEach-Object {
        # Registro el nombre de la connexión
        Write-Host 'Conexión: '$_.Node.connectionManagerRefId              
    }
      
    #Miro si tiene SQL Command de Variable, porque si tiene, no tengo escribir el SQLCommand. 
    #Inicializo la variable de comprobación
    $UseSQLCmdVble = $false  
    $childs| Select-XML -XPath "//property[@name='SqlCommandVariable']"  | ForEach-Object {
        if ($_.Node.InnerText) {
            # Registro el comando SQL
            Write-Host 'Comando SQL en la Variable: ' + $_.Node.InnerText
            $UseSQLCmdVble = $true     
        }                        
    }
      
    if (-not $UseSQLCmdVble){
      
        $childs| Select-XML -XPath "//property[@name='SqlCommand']"  | ForEach-Object {
            # Registro el comando SQL
            Write-Host 'Comando SQL: '$_.Node.InnerText        
        }
    }     
          
}