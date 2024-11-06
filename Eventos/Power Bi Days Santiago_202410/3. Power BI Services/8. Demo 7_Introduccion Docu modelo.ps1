#Dll Microsoft.AnalysisServices.Tabular.dll
Add-Type -Path 'C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.AnalysisServices.Tabular\v4.0_15.0.0.0__89845dcd8080cc91\Microsoft.AnalysisServices.Tabular.dll'

#Datos modelo tabular
$serverName = "powerbi://api.powerbi.com/v1.0/myorg/********" #Nombre del workspace
$User = $Env:POWERBI_USER
$Pass = $Env:POWERBI_PASS

$connectionString="Data source='$serverName';User ID='$User';Password='$Pass'"
$server = New-Object Microsoft.AnalysisServices.Tabular.Server
$server.Connect($connectionString)

#Los Databases son los Modelos Semánticos
ForEach ($db in $server.Databases)
{
    Write-Output $db.Name
}

$db = $server.Databases.Item(2) #Me quedo con el Full

#Cojo el modelo
$model = $db.Model

ForEach ($Table in $model.Tables) 
{
    #Miramos si la tabla la ve o no el usuario
    if ($Table.isHidden -eq $true) {
        $Hidden = $true
    } else {$Hidden = $false}


    If ($Hidden -eq $false)
    {        
        #Contamos las medidas
        $Measures= $Table.Measures.Count 
    
        #Contamos las columnas
        $Columns = $Table.Columns.Count

        write-Host $Table.Name'--> Num Medidas:' $Measures '; Num Columnas: ' $Columns
    }
}