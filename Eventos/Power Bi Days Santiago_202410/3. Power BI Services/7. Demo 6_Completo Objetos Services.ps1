###############################################################################################
# Datos a rellenar
#
# $workspacesToIgnore --> Lista de nombres de workspaces que se quieren ignorar
#
# $pathdocfile --> Url donde se dejara el documento
#
# $DocWorkspace --> Para indicar si documentamos los workspaces
# $DocApps --> Para indicar si documentamos las Apps
# $DocPipelines --> Para indicar si documentamos las Pipelines
# $DocDatasets --> Para indicar si documentamos los datasets
###############################################################################################

#Defino el tipo DataSet
class DataSetType
{
    #Attributes
    [string]$Id
    [string]$Name
    [string]$ConfiguredBy
	[string]$WorkspaceName
    [System.Collections.Generic.SortedList[string,string]]$ReportList #Key: ID, Value: Workspace name | Report name

    # Constructor
    DataSetType($Id,$Name, $ConfiguredBy,$WorkspaceName ) {
        $this.Id = $Id
        $this.Name = $Name
        $this.ConfiguredBy = $ConfiguredBy
		$this.WorkspaceName = $WorkspaceName
        $this.ReportList = New-Object 'System.Collections.Generic.SortedList[string,string]'
    }
}

$password = $Env:POWERBI_PASS | ConvertTo-SecureString -asPlainText -Force
$username = $Env:POWERBI_USER
$credential = New-Object System.Management.Automation.PSCredential($username, $password)

#Creamos una lista de DataSets para recopilar la info y luego mostrarla
$DatasetsList = New-Object 'System.Collections.Generic.SortedList[string,DataSetType]'

#Creamos una lista de Workspaces a ignorar
$workspacesToIgnore = @('Nombre a ignorar')

$DocWorkspace = $true
$DocApps = $true
$DocPipelines = $true
$DocDatasets = $true
#########################################################################################################################################################################
####################                                     CREACIÓN DOC                                                                              ######################
#########################################################################################################################################################################

#Configuración del documento Word
#url
$pathdocfile = "C:\Charlas\20241026_PowerBIDaysSantiago2024_DocumentaTuBIYTriunfa\3. Power BI Services"
$today = Get-Date -Format "yyyyMMdd_hhmmss"
$modelname = $db.Name
$modelserver = $db.Server    

$docfile = "$pathdocfile\PowerBIServices Components $today.docx"

#Creamos instancia del objeto word
$MSWord = New-Object -ComObject Word.Application
#Lo hacemos visible
$MSWord.Visible = $True
#Creamos nuevo documento
$mydoc = $MSWord.Documents.Add()

$comObject = $mydoc.BuiltInDocumentProperties("Title")
$binding = "System.Reflection.BindingFlags" -as [type]
[System.__ComObject].invokemember("Value",$binding::SetProperty,$null,$comObject,"PowerBIServices Components")

$comObject = $mydoc.BuiltInDocumentProperties("Company")
$binding = "System.Reflection.BindingFlags" -as [type]
[System.__ComObject].invokemember("Value",$binding::SetProperty,$null,$comObject,"Power BI Days Santiago 2024")

$comObject = $mydoc.BuiltInDocumentProperties("Author")
$binding = "System.Reflection.BindingFlags" -as [type]
[System.__ComObject].invokemember("Value",$binding::SetProperty,$null,$comObject,"Cristina Tarabini-Castellani")

$CoverPage = 'Movimiento' #Nombre de la plantilla de la portada
$Selection = $MSWord.application.selection   

$MSWord.Templates.LoadBuildingBlocks()
$bb =$MSWord.templates | Where-Object -Property name -EQ -Value 'Built-In Building Blocks.dotx'

$part = $bb.BuildingBlockEntries.item($CoverPage)

$null = $part.Insert($MSWord.Selection.range, $true)   


#Borramos la imagen por defecto de la portada
$Selection.Document.Shapes.Item('Imagen 1').Delete()

#Creamos el nuevo Canva para recoger el logo y colocarla en el mismo lugar
$null = $Selection.Document.Shapes.AddCanvas(1, 1, 200, 107.5)
$null = $Selection.Document.Shapes.Item('Lienzo 1').CanvasItems.AddPicture("C:\Charlas\20241026_PowerBIDaysSantiago2024_DocumentaTuBIYTriunfa\ImagenPBDSantiago.png", $false, $True, $null, $null, 200, 107.5)#439.2, 295.55
$Selection.Document.Shapes.Item('Lienzo 1').Left =-999996
$Selection.Document.Shapes.Item('Lienzo 1').Top = -999995
$Selection.Document.Shapes.Item('Lienzo 1').Height = 107.5
$Selection.Document.Shapes.Item('Lienzo 1').Width = 200
$Selection.Document.Shapes.Item('Lienzo 1').Title = "Logo"
$Selection.Document.Shapes.Item('Lienzo 1').AlternativeText = "Imagen del logo"
$Selection.Document.Shapes.Item('Lienzo 1').RelativeHorizontalPosition = 1
$Selection.Document.Shapes.Item('Lienzo 1').RelativeVerticalPosition = 1
$Selection.Document.Shapes.Item('Lienzo 1').CanvasItems.PictureFormat

#Reemplazamos la etiqueta [Año]
$año = Get-Date -Format "yyyy"
$Selection.Document.Shapes.Item("Grupo 252").GroupItems[3].TextFrame.TextRange.Text= $año #Año

#Reemplazamos la etiqueta [Fecha]
$fecha = Get-Date -Format "dd/MM/yyyy"
$Texto = $Selection.Document.Shapes.Item("Grupo 252").GroupItems[4].TextFrame.TextRange.Text
$Selection.Document.Shapes.Item("Grupo 252").GroupItems[4].TextFrame.TextRange.Text = $Texto.Replace("[Fecha]",$fecha) #Bloque de abajo

#Metemos un salto de página para incluir el índice
$Selection = $MSWord.Selection
$Selection.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdSectionBreakNextPage)
$Selection = $MSWord.Selection
$range = $Selection.Range	
$toc = $mydoc.TablesOfContents.Add($range)
$Selection.TypeParagraph()

#Metemos un salto de página para empezar a documentar
$Selection.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdSectionBreakNextPage)
    
#Creamos array de stilos para las cabeceras
$styles = @(
        [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading1
    ,[Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading2
    ,[Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading3
    ,[Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading4
    ,[Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading5
    ,[Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading6
    ,[Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading7
    ,[Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading8
    ,[Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading9
    
)

#Configuramos la apariencia del documento a partir de aquí
$Section = $mydoc.Sections.Item(1)
$Header = $Section.Headers.Item(1)
$Footer = $Section.Footers.Item(1)
$Footer.PageNumbers.Add()

#########################################################################################################################################################################
####################                         EMPEZAMOS A RECOGER INFO DEL SERVICIO                                                                 ######################
#########################################################################################################################################################################

Connect-PowerBIServiceAccount -Credential $credential

# Si tenemos que documentar los Datasets, necesito recoger info de los Workspaces
if (($DocWorkspace -eq $true) -or ($DocDatasets -eq $true)) {
    ##Obtenemos los workspaces
    $workspaces = Get-PowerBIWorkspace

    #########################################################################################################################################################################
    ####################                         EMPEZAMOS HACIENDO EL SCAN DE LOS WORKSPACES                                                          ######################
    #########################################################################################################################################################################

    #Hacemos el scanner de todos los workspaces para obtener info que la API de PowerShell no da
    $numworkspace = 0
    $body =
    '{
        "workspaces":['

    #Recogemos el id de todos los workspaces
    ForEach($workspace in $workspaces){
        If ($workspace.Name -in $workspacesToIgnore)
        {
            $ignore = $true
        } else {
            $ignore = $false
        }

        if (-not $ignore){
            $numworkspace = $numworkspace + 1 
            if ($numworkspace -gt 1)
            {
                $body =  $body + ','
            }

            $body = $body + '
                "'+ $workspace.id + '"'
        }
    }

    $body = $body + '
        ]
    }'

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
    
    #########################################################################################################################################################################
    ####################                         FIN DEL SCAN DE LOS WORKSPACES                                                                        ######################
    #########################################################################################################################################################################

    if ($DocWorkspace -eq $true){
        # Crea el título 1 con el nombre del workspace
        $myText = $MSWord.Selection
        $myText.Style = $styles[0]
        $myText.TypeText("Workspaces")
        $myText.TypeParagraph()  
    }

    #Por cada workspace miramos la info de usuarios y permisos
    ForEach($workspace in $workspaces){
        If ($workspace.Name -in $workspacesToIgnore)
        {
            $ignore = $true
        } else {
            $ignore = $false
        }

        if (-not $ignore){
            $workspacescan = $scanresult.workspaces | Where-Object { $_.id -eq $workspace.id }

            if ($DocWorkspace -eq $true){                
                # Crea el título 2 con el nombre del workspace
                $myText = $MSWord.Selection
                $myText.Style = $styles[1]
                $myText.TypeText($workspace.Name)
                $myText.TypeParagraph() 
            
                #Indicamos la descripción del Workspace
                $myText = $MSWord.Selection
                $myText.TypeText($workspacescan.description)
                $myText.TypeParagraph()  

                # Crea el título 3 con Permisos
                $myText = $MSWord.Selection
                $myText.Style = $styles[2]
                $myText.TypeText("Permisos")
                $myText.TypeParagraph()  

                #Lista todos los usuarios del Workspaces y sus datos
                #API url para usuarios y grupos del workspace
                $ApiUrl = "groups/" + $workspace.Id + "/users"
                $Users = (Invoke-PowerBIRestMethod -Url $ApiUrl -Method Get) | ConvertFrom-Json

                #Creamos la tabla para recoger los permisos de este workspace
                $myText = $MSWord.Selection
                $Range = @($myText.Paragraphs)[-1].Range
                $Table = $myText.Tables.add(
                    $myText.Range #$Range
                    ,$Users.value.Count + 1 #Rows
                    ,4 #Columns
                    ,[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior
                    ,[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
                )

                ## Header
                $Table.cell(1,1).range.Bold=1
                $Table.cell(1,1).range.text = "Nombre"
                $Table.cell(1,2).range.Bold=1
                $Table.cell(1,2).range.text = "Tipo"
                $Table.cell(1,3).range.Bold=1
                $Table.cell(1,3).range.text = "Permiso"
                $Table.cell(1,4).range.Bold=1
                $Table.cell(1,4).range.text = "Identificador"

                #Empezamos a mirar los datos
                $row = 2
                ForEach ($User in $Users.value) {  
                    if ($User.groupUserAccessRight -eq "User") {
                        $UName = $User.emailAddress
                    }else{
                        $UName = $User.displayName
                    }

                    $Table.cell($row,1).range.Bold = 0
                    $Table.cell($row,1).range.text = $UName                              
                    $Table.cell($row,2).range.Bold = 0
                    $Table.cell($row,2).range.text = $User.principalType
                    $Table.cell($row,3).range.Bold = 0
                    $Table.cell($row,3).range.text = $User.groupUserAccessRight
                    $Table.cell($row,4).range.Bold = 0
                    $Table.cell($row,4).range.text = $User.identifier
        
                    $row++  
                }
        
                #Para continuar escribiendo fuera de la tabla y dejar una linea
                $myText.Start= $mydoc.Content.End
                $myText.TypeParagraph()
            } 
    
            #Lista todos los reports del Workspaces y sus datos
            #API url para usuarios y grupos del workspace
            $ApiUrl = "groups/" + $workspace.Id + "/reports"
            $Reports = (Invoke-PowerBIRestMethod -Url $ApiUrl -Method Get) | ConvertFrom-Json
                    
            If ($Reports.Value.Count -gt 0)
            {
                if ($DocWorkspace -eq $true){
                    # Crea el título 2 con Permisos
                    $myText = $MSWord.Selection
                    $myText.Style = $styles[2]
                    $myText.TypeText("Informes")
                    $myText.TypeParagraph()              
                
                    #Creamos la tabla para recoger las medidas de esta carpeta
                    $myText = $MSWord.Selection
                    $Range = @($myText.Paragraphs)[-1].Range
                    $Table = $myText.Tables.add(
                        $myText.Range #$Range
                        ,$Reports.value.Count + 1 #Rows
                        ,5 #Columns
                        ,[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior
                        ,[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
                    )
            
                    ## Header
                    $Table.cell(1,1).range.Bold=1
                    $Table.cell(1,1).range.text = "Nombre"
                    $Table.cell(1,2).range.Bold=1
                    $Table.cell(1,2).range.text = "Dataset"
                    $Table.cell(1,3).range.Bold=1
                    $Table.cell(1,3).range.text = "Modificado por"
                    $Table.cell(1,4).range.Bold=1                    
                    $Table.cell(1,4).range.text = "Fecha Modificación"
                    $Table.cell(1,5).range.Bold=1                    
                    $Table.cell(1,5).range.text = "Descripción"

                    #Empezamos a mirar los datos
                    $row = 2
                }
    
                ForEach ($report in $Reports.value) {
                    $reportscan = $workspacescan.reports | Where-Object { $_.id -eq $report.id }

                    if ($DocDatasets -eq $true){
                        #Tratamos el DataSetId
                        #Miramos si ya hemos dado de alta este dataset, para recuperarlo o no
                        if ($DatasetsList.ContainsKey($report.datasetId)){
                            $dataset = [DataSetType]$DatasetsList[$report.datasetId]
                            #Lo borramos porque lo vamos a volver a cargar
                            $r = $DatasetsList.Remove($report.datasetId)
                        }
                        else
                        {
                            #Recuperamos la info del Dataset
                            $ApiUrl = "datasets/" + $report.datasetId
                            $reportdataset = (Invoke-PowerBIRestMethod -Url $ApiUrl -Method Get) | ConvertFrom-Json

                            $dataset = [DataSetType]::new($reportdataset.id, $reportdataset.Name, $reportdataset.ConfiguredBy, $workspace.Name)   
                        }

                        #Construimos el reportname para añadirlo al ReportList
                        $reportname = $workspace.Name," → ", $report.name
                        if ( -Not $dataset.ReportList.ContainsKey($report.id)){
                            $dataset.ReportList.Add($report.id, $reportname)
                        }
        
                        $DatasetsList.Add($report.datasetId, $dataset) 
                    }
                    if ($DocWorkspace -eq $true){
                        $Table.cell($row,1).range.Bold = 0
                        $Table.cell($row,1).range.text = $report.name                              
                        $Table.cell($row,2).range.Bold = 0
                        $Table.cell($row,2).range.text = $dataset.name                             
                        $Table.cell($row,3).range.Bold = 0
                        $Table.cell($row,3).range.text = $reportscan.modifiedBy                             
                        $Table.cell($row,4).range.Bold = 0
                        $Table.cell($row,4).range.text = $reportscan.modifiedDateTime                             
                        $Table.cell($row,5).range.Bold = 0
                        $Table.cell($row,5).range.text = $reportscan.description 

                        $row++ 
                    }
                }

                if ($DocWorkspace -eq $true){
                    #Para continuar escribiendo fuera de la tabla y dejar una linea
                    $myText.Start= $mydoc.Content.End
                    $myText.TypeParagraph()
                }
            }

            if ($DocWorkspace -eq $true){

                #Lista todos los dashboards del Workspaces y sus datos
                #API url para usuarios y grupos del workspace
                $ApiUrl = "groups/" + $workspace.Id + "/dashboards"
                $dashboards = (Invoke-PowerBIRestMethod -Url $ApiUrl -Method Get) | ConvertFrom-Json

                If ($dashboards.Value.Count -gt 0)
                {
                    # Crea el título 2 con Permisos
                    $myText = $MSWord.Selection
                    $myText.Style = $styles[2]
                    $myText.TypeText("Paneles")
                    $myText.TypeParagraph()    

                    #Creamos la tabla para recoger las medidas de esta carpeta
                    $myText = $MSWord.Selection
                    $Range = @($myText.Paragraphs)[-1].Range
                    $Table = $myText.Tables.add(
                        $myText.Range #$Range
                        ,$dashboards.value.Count + 1 #Rows
                        ,1 #Columns
                        ,[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior
                        ,[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
                    )
            
                    ## Header
                    $Table.cell(1,1).range.Bold=1
                    $Table.cell(1,1).range.text = "Nombre"

                    #Empezamos a mirar los datos
                    $row = 2
        
                    ForEach ($dashboard in $dashboards.value) {
                        $Table.cell($row,1).range.Bold = 0
                        $Table.cell($row,1).range.text = $dashboard.name  
                        $row++ 
                    }

                    #Para continuar escribiendo fuera de la tabla y dejar una linea
                    $myText.Start= $mydoc.Content.End
                    $myText.TypeParagraph()
                }   
            }
        }
    }

}
#Revisamos las aplicaciones creadas
$ApiUrl = "apps"
$Apps = (Invoke-PowerBIRestMethod -Url $ApiUrl -Method Get) | ConvertFrom-Json

if (($DocApps -eq $true) -and ($Apps.value.Count -gt 0)){
    # Crea el título 1 con el nombre del workspace
    $myText = $MSWord.Selection
    $myText.Style = $styles[0]
    $myText.TypeText("Aplicaciones")
    $myText.TypeParagraph()  

    $rows = 0
    ForEach ($App in $Apps.value) {
            $rows = $rows + 1
    }

    #Creamos la tabla para recoger los permisos de este workspace
    $myText = $MSWord.Selection
    $Range = @($myText.Paragraphs)[-1].Range
    $Table = $myText.Tables.add(
        $myText.Range #$Range
        ,$rows
        ,2 #Columns
        ,[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior
        ,[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
    )
                
    ## Header
    $Table.cell(1,1).range.Bold=1
    $Table.cell(1,1).range.text = "Nombre"
    $Table.cell(1,2).range.Bold=1
    $Table.cell(1,2).range.text = "Descripción"

    #Empezamos a mirar los datos
    $row = 2
    
    ForEach ($App in $Apps.value) {
            $Table.cell($row,1).range.Bold = 0
            $Table.cell($row,1).range.text = $App.name                              
            $Table.cell($row,2).range.Bold = 0
            $Table.cell($row,2).range.text = $App.description

            $row++ 

    }

    #Para continuar escribiendo fuera de la tabla y dejar una linea
    $myText.Start= $mydoc.Content.End
    $myText.TypeParagraph()
}

#Revisamos los pipelines creados
$ApiUrl = "pipelines"
$pipelines = (Invoke-PowerBIRestMethod -Url $ApiUrl -Method Get) | ConvertFrom-Json

if (($DocPipelines -eq $true) -and ($pipelines.value.Count -gt 0 )) {
    
    # Crea el título 1 
    $myText = $MSWord.Selection
    $myText.Style = $styles[0]
    $myText.TypeText("Pipelines")
    $myText.TypeParagraph()  
    
    ForEach ($pipeline in $pipelines.value) {
     
            # Crea el título 2 con el nombre del pipeline
            $myText = $MSWord.Selection
            $myText.Style = $styles[1]
            $myText.TypeText($pipeline.displayName)
            $myText.TypeParagraph()  

            $myText = $MSWord.Selection
            $myText.TypeText($pipeline.description)
            $myText.TypeParagraph() 
        
            #Lista todos los usuarios del pipeline y sus datos
            #API url para usuarios y grupos del pipeline
            $ApiUrl = "pipelines/" + $pipeline.Id + "/users"
            $PUsers = (Invoke-PowerBIRestMethod -Url $ApiUrl -Method Get) | ConvertFrom-Json
        
            if ($PUsers.value.Count -gt 0 ) {
                # Crea el título 3 con el nombre del pipeline
                $myText = $MSWord.Selection
                $myText.Style = $styles[2]
                $myText.TypeText("Usuarios")
                $myText.TypeParagraph()  

                #Creamos la tabla para recoger los permisos de este workspace
                $myText = $MSWord.Selection
                $Range = @($myText.Paragraphs)[-1].Range
                $Table = $myText.Tables.add(
                    $myText.Range #$Range
                    ,$PUsers.value.Count + 1 #Rows
                    ,3 #Columns
                    ,[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior
                    ,[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
                )
                
                ## Header
                $Table.cell(1,1).range.Bold=1
                $Table.cell(1,1).range.text = "Identificador"
                $Table.cell(1,2).range.Bold=1
                $Table.cell(1,2).range.text = "Tipo"
                $Table.cell(1,3).range.Bold=1
                $Table.cell(1,3).range.text = "Permiso"

                #Empezamos a mirar los datos
                $row = 2

                ForEach ($PUser in $PUsers.value) {
                    $Table.cell($row,1).range.Bold = 0
                    $Table.cell($row,1).range.text = $PUser.identifier                              
                    $Table.cell($row,2).range.Bold = 0
                    $Table.cell($row,2).range.text = $PUser.principalType                          
                    $Table.cell($row,3).range.Bold = 0
                    $Table.cell($row,3).range.text = $PUser.accessRight
                
                    $row++ 
                }

                #Para continuar escribiendo fuera de la tabla y dejar una linea
                $myText.Start= $mydoc.Content.End
                $myText.TypeParagraph()
            }
    
            #Pasos del pipeline
            #API url para usuarios y grupos del pipeline
            $ApiUrl = "pipelines/" + $pipeline.Id + "/stages"
            $Stages = (Invoke-PowerBIRestMethod -Url $ApiUrl -Method Get) | ConvertFrom-Json
        
            if ($Stages.value.Count -gt 0 ) {
                # Crea el título 3 con el nombre del pipeline
                $myText = $MSWord.Selection
                $myText.Style = $styles[2]
                $myText.TypeText("Pasos")
                $myText.TypeParagraph()  

                #Creamos la tabla para recoger los permisos de este workspace
                $myText = $MSWord.Selection
                $Range = @($myText.Paragraphs)[-1].Range
                $Table = $myText.Tables.add(
                    $myText.Range #$Range
                    ,$Stages.value.Count + 1 #Rows
                    ,2 #Columns
                    ,[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior
                    ,[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
                )
                
                ## Header
                $Table.cell(1,1).range.Bold=1
                $Table.cell(1,1).range.text = "Orden"
                $Table.cell(1,2).range.Bold=1
                $Table.cell(1,2).range.text = "Nombre Workspace"

                #Empezamos a mirar los datos
                $row = 2
                ForEach ($Stage in $Stages.value) {
                    $Table.cell($row,1).range.Bold = 0
                    $Table.cell($row,1).range.text = $Stage.order.ToString()                             
                    $Table.cell($row,2).range.Bold = 0
                    $Table.cell($row,2).range.text = $Stage.workspaceName   
                
                    $row++ 
                }

                #Para continuar escribiendo fuera de la tabla y dejar una linea
                $myText.Start= $mydoc.Content.End
                $myText.TypeParagraph()
            }
        
    }
}

#Por cada DATASET que hemos usado, miramos a ver qué reports se usan
if (($DocDatasets -eq $true) -and ($DatasetsList.Keys.Count -gt 0)) {
    
    # Crea el título 1 con el nombre del workspace
    $myText = $MSWord.Selection
    $myText.Style = $styles[0]
    $myText.TypeText("Datasets")
    $myText.TypeParagraph()  

    ForEach ($key in $DatasetsList.Keys)
    {
        $d = [DataSetType]$DatasetsList[$key]
		
		$FullName = $d.WorkspaceName + " → " + $d.Name

        # Crea el título 2 con el nombre del DataSet
        $myText = $MSWord.Selection
        $myText.Style = $styles[1]
        $myText.TypeText($FullName)
        $myText.TypeParagraph() 
        
        $myText = $MSWord.Selection
        $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
        $myText.Font.Bold = 1
        $myText.TypeText("Id:")
        $myText.TypeParagraph()  

        $myText = $MSWord.Selection
        $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStylePlainText
        $myText.Font.Bold = 0
        $myText.TypeText($d.Id)
        $myText.TypeParagraph()  

        $myText = $MSWord.Selection
        $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
        $myText.Font.Bold = 1
        $myText.TypeText("Configurado por:")
        $myText.TypeParagraph()  

        $myText = $MSWord.Selection
        $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStylePlainText
        $myText.Font.Bold = 0
        $myText.TypeText($d.configuredBy)
        $myText.TypeParagraph() 

        if($d.ReportList.Count -gt 0 ) {
            
            # Crea el título 2 con el nombre del DataSet
            $myText = $MSWord.Selection
            $myText.Style = $styles[2]
            $myText.TypeText("Lista de Informes")
            $myText.TypeParagraph() 

            ForEach ($dreport in $d.ReportList.Values)
            {
                $myText = $MSWord.Selection
                $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
                $myText.TypeText($dreport)
                $myText.TypeParagraph() 
            }
        }    
        
        
        #API url para usuarios y grupos del workspace
        $ApiUrl = "datasets/" + $d.Id + "/datasources"
        $datasources = (Invoke-PowerBIRestMethod -Url $ApiUrl -Method Get) | ConvertFrom-Json        
        
        if($datasources.value.Count -gt 0 ) {
            # Crea el título 2 con el nombre del DataSet
            $myText = $MSWord.Selection
            $myText.Style = $styles[2]
            $myText.TypeText("Origenes de datos")
            $myText.TypeParagraph() 

            #Creamos la tabla para recoger las datasources de esta carpeta
            $myText = $MSWord.Selection
            $Range = @($myText.Paragraphs)[-1].Range
            $Table = $myText.Tables.add(
                $myText.Range #$Range
                ,$datasources.value.Count + 1 #Rows
                ,4 #Columns
                ,[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior
                ,[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
            )
            
            ## Header
            $Table.cell(1,1).range.Bold=1
            $Table.cell(1,1).range.text = "Id"
            $Table.cell(1,2).range.Bold=1
            $Table.cell(1,2).range.text = "Tipo"
            $Table.cell(1,3).range.Bold=1
            $Table.cell(1,3).range.text = "Servidor"
            $Table.cell(1,4).range.Bold=1
            $Table.cell(1,4).range.text = "Base de datos"
            
            #Empezamos a mirar los datos
            $row = 2
            
            ForEach ($datasource in $datasources.value) {
                $Table.cell($row,1).range.Bold = 0
                $Table.cell($row,1).range.text = $datasource.datasourceId                           
                $Table.cell($row,2).range.Bold = 0
                $Table.cell($row,2).range.text = $datasource.datasourceType                   
                $Table.cell($row,3).range.Bold = 0
                $Table.cell($row,3).range.text = $datasource.connectionDetails.server                          
                $Table.cell($row,4).range.Bold = 0
                $Table.cell($row,4).range.text = $datasource.connectionDetails.database
                
                $row++ 
            }
            
            #Para continuar escribiendo fuera de la tabla y dejar una linea
            $myText.Start= $mydoc.Content.End
            $myText.TypeParagraph()
        }
    
    }
}
Disconnect-PowerBIServiceAccount

#Actualizamos el índice	
$toc.Update()

#Salvamos el documento Word
$saveFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault
$mydoc.SaveAs([ref][system.object]$docfile, [ref]$saveFormat)
$mydoc.Close()
$MSWord.Quit()

# Clean up Com object
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSWord)
Remove-Variable MSWord