###############################################################################################
# Datos a rellenar
# Path de la dll
# $SQLServer --> Ip y puerto del servidor o nombre del servidor
# $BDUser --> Login para conectarse a la BBDD
# $BDUserPass --> Contraseña del login para conectarse a la BBDD
# $docfile --> Url donde se dejara el documento
###############################################################################################
#Dll que hay que instalar y buscar el path correspondiente
Add-Type -Path 'C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.SqlServer.Smo\v4.0_16.0.0.0__89845dcd8080cc91\Microsoft.SqlServer.Smo.dll'

#Configuración del servidor de BBDD
$SQLServer = **********
$BDUser = ***********
$BDUserPass = ****************
$BBDD = ***************

#Preparo la conexión
$SmoServer = New-Object ('Microsoft.SqlServer.Management.Smo.Server') -argumentlist $SQLServer
$SmoServer.ConnectionContext.LoginSecure = $false
$SmoServer.ConnectionContext.set_Login($BDUser)
$SmoServer.ConnectionContext.set_Password($BDUserPass)

#Recojo toda la info de la BD
$db = $SmoServer.Databases[$BBDD]

Write-Host $db.Name

#Recogemos las tablas
foreach ($ta in $db.Tables) {  
    $schema = $ta.schema
    $tabla = $ta.Name
   
   #Muestro el nombre de la tabla
    Write-Host  "[$schema].[$tabla]"

    #Preparo la query que recoge las descripciones de las tablas y la lanzo
    $cmd ="SELECT
	            p.name
	            ,p.value
            FROM sys.tables t
		            left join sys.extended_properties p
			            on p.major_id = t.object_id	and p.minor_id = 0 /*tabla*/
            WHERE t.name = '$tabla' and t.schema_id = SCHEMA_ID('$schema')"
    $ds = $db.ExecuteWithResults($cmd)

    #Recojo los resultados
    if ($ds.Tables[0].Rows.Count -gt 0) {
        for($j=0; $j -lt $ds.Tables[0].Rows.Count; $j++){
            $name = $ds.Tables[0].Rows[$j].name
            $value = $ds.Tables[0].Rows[$j].value

            #Escribo el valor
            Write-Host "$name --> $value"
        }
    }
}

