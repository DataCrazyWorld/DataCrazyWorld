#Preparamos los credenciales
$password = $Env:POWERBI_PASS | ConvertTo-SecureString -asPlainText -Force
$username = $Env:POWERBI_USER
$credential = New-Object System.Management.Automation.PSCredential($username, $password)

#Conectamos al servicio
Connect-PowerBIServiceAccount -Credential $credential

#Mostramos los Workshops Cmdslet
Get-PowerBIWorkspace

#Mostramos los Workshops Cmdslet
Write-Host "Cmdlets"
Get-PowerBIWorkspace -Name "BI [Dev]"
$workspace = Get-PowerBIWorkspace -Name "BI [Dev]"

#No sale la Descripción, ni los reports...

#API Rest
Write-Host "API Rest"
$ApiUrl = "groups/" + $workspace.Id 
(Invoke-PowerBIRestMethod -Url $ApiUrl -Method Get) | ConvertFrom-Json

#Scan
Write-Host "SCAN"

#Body para el API

$body =
    '{
        "workspaces":["'+ $workspace.id + '"]
    }'

#Se lanza la solicitud de SCAN - Lanzamos uno completo que luego usaremos en todos los ejemplos
#Llamamos para generar el Scan y tener la ID
$ApiUrl = "admin/workspaces/getInfo?lineage=True&datasourceDetails=True&datasetSchema=True&datasetExpressions=True&getArtifactUsers=True"
$Scan = (Invoke-PowerBIRestMethod -Url $ApiUrl -Method Post -Body $body) | ConvertFrom-Json

#Vamos a obtener el scanStatus, porque si está Running no podemos continuar
$ApiUrl = "admin/workspaces/scanstatus/" + $Scan.id
$Scanstatus = (Invoke-PowerBIRestMethod -Url $ApiUrl -Method Get) | ConvertFrom-Json
    
While ($Scanstatus.status -eq "Running")
{
    Start-Sleep -Seconds 1
    $Scanstatus = (Invoke-PowerBIRestMethod -Url $ApiUrl -Method Get) | ConvertFrom-Json
    
}

#Vamos a obtener el scanResult
$ApiUrl = "admin/workspaces/scanResult/" + $Scan.id
$scanresult = (Invoke-PowerBIRestMethod -Url $ApiUrl -Method Get) | ConvertFrom-Json

#Vemos que información más detallada tiene sobre los workspaces
$workspacescan = $scanresult.workspaces | Where-Object { $_.id -eq $workspace.id }

$workspacescan

$workspacescan.description