###############################################################################################
# Datos a rellenar
# $pathdocfile --> Url donde se dejara el documento
#
# $SQLServer --> Ip y puerto del servidor
#
# $IA --> indicar si no hay descripción si queremos que vaya a la IA
#
###############################################################################################

#########################################################################################################################################################################
####################                                     FUNCIÓN DE ACCESO A LA IA                                                                 ######################
#########################################################################################################################################################################
function Get-AIDescription ([string]$expresion, [string] $expertise)
{
    # Azure OpenAI metadata variables
    $openai = @{
        api_key     = $Env:AZURE_OPENAI_KEY
        api_base    = $Env:AZURE_OPENAI_ENDPOINT # your endpoint should look like the following https://YOUR_RESOURCE_NAME.openai.azure.com/
        api_version = '2023-03-15-preview' # this may change in the future
        name        = 'PowerBI-gpt4o'#This will correspond to the custom name you chose for your deployment when you deployed a model.
    }
    # Completion text
    if ($expertise -eq "DAX"){
        #$promptPrefix = "Por favor, actúa como un experto en DAX."
        $promptPrefix = "Por favor, sé un experto en DAX."
        $promptSufix = "Responde de manera concisa a la siguiente pregunta sobre el código DAX escrito arriba.
        Q:Por favor, describe brevemente en español qué cálculo se este realizando.";

        $max_tokens = 100
        $UsarIA = $True
    } elseif ($expertise -eq "M") {
        #$promptPrefix = "Por favor, actúa como un experto muy detallista en Power Query."
        #$promptSufix = "Responde de manera concisa a la siguiente pregunta sobre el código de Power Query escrito arriba.
        #Q: Describe con detalle y en español en qué consiste esta transformación que se está realizando en Power Query, incluyendo los nombres de tablas, columnas, 
        #descripción de los filtros que se hacen y las transformaciones. No traduzcas lo que significan las columnas ni pongas una descripción de ellas, pero listalas cada una en una fila. 
        #Omite el codigo M en la explicación";
        
        $promptPrefix = "Por favor, sé un experto muy detallista en Power Query."
        $promptSufix ="Responde de manera concisa a la siguiente pregunta sobre el código de Power Query escrito arriba.
        Q: Describe con detalle y en español que se está realizando en esta transformación de Power Query, incluyendo los nombres de tablas, columnas, filtros y transformaciones que se hacen. Sin traducir columnas ni añadir descripciones, sólo listar cada una en una fila. Omite el código M en la explicación."
        $max_tokens = 500
        $UsarIA = $True
    } else {
        $UsarIA = $False
    }
    #Solo si hay que usar la IA
    if ($UsarIA) {
       $usermessage = "#$expresion 
       #$promptSufix"
        # Header for authentication
        $headers = [ordered]@{
            'api-key' = $openai.api_key
        }
        $messages  = @()
        $messages += @{
          role     = 'system'
          content  = $promptPrefix
        }
        $messages += @{
          role     = 'user'
          content  = $usermessage
        }
        ## Adjust these values to fine-tune completions
        $body         = @{
          #seed       = 42
          temperature = 0.5
          max_tokens  = $max_tokens
          messages    = $messages
        } | ConvertTo-Json
        # Send a completion call to generate an answer
        $url = "$($openai.api_base)/openai/deployments/$($openai.name)/chat/completions?api-version=$($openai.api_version)"
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Body $body -Method Post -ContentType 'application/json; charset=UTF-8' 
        ##Para poner bien las tildes y caracteres "raros"
        $description =[System.Text.Encoding]::UTF8.GetString($response.choices[0].message.content.ToCharArray())
        if ($description.contains("R:")){
            $description.Substring($description.indexOf("R:")+2).TrimStart().TrimEnd()
        }elseif ($description.contains("A:")) {
            $description.Substring($description.indexOf("A:")+2).TrimStart().TrimEnd()
        } else {
            $description.TrimStart().TrimEnd()
        }
     }
}
#########################################################################################################################################################################

#Dll
Add-Type -Path 'C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.SqlServer.Smo\v4.0_16.0.0.0__89845dcd8080cc91\Microsoft.SqlServer.Smo.dll'
Add-Type -Path 'C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.AnalysisServices.Tabular\v4.0_15.0.0.0__89845dcd8080cc91\Microsoft.AnalysisServices.Tabular.dll'

#Consultamos la IA sí o no
$IA = $true

#Defino un tipo de tabla
class TableType
{
    #Attributes
    [string]$Name
    [boolean]$isHidden
    [boolean]$IsSQL
    [boolean]$IsMeasure
    [string]$Description
    [System.Collections.IDictionary]$List #En función del IsMeasure será de un tipo String|String o String|MeasureType

    # Constructor
    TableType($Name,$isHidden, $IsSQL, $IsMeasure) {
        $this.Name = $Name
        $this.isHidden = $isHidden
        $this.IsSQL = $IsSQL
        $this.IsMeasure = $IsMeasure
    }
}

#Defino un tipo medida
class MeasureType
{
    [String]$Name
    [String]$DAXExpression
    [String]$Description

    #Constructor
    MeasureType($Name, $DAXExpression, $Description) {
        $this.Name = $Name
        $this.DAXExpression = $DAXExpression
        $this.Description = $Description
    }
}

#Datos modelo tabular
$serverName = "powerbi://api.powerbi.com/v1.0/myorg/******" #Nombre del workspace
$User = $Env:POWERBI_USER
$Pass = $Env:POWERBI_PASS

$connectionString="Data source='$serverName';User ID='$User';Password='$Pass'"
$server = New-Object Microsoft.AnalysisServices.Tabular.Server
$server.Connect($connectionString)

#Configuración del servidor de BBDD de donde coger los metadatos
$SQLServer = *************
$BDUser = ********************
$BDUserPass = ***************
$BBDD = *****************

$SmoServer = New-Object ('Microsoft.SqlServer.Management.Smo.Server') -argumentlist $SQLServer
$SmoServer.ConnectionContext.LoginSecure = $false
$SmoServer.ConnectionContext.set_Login($BDUser)
$SmoServer.ConnectionContext.set_Password($BDUserPass)

$dbsql = $SmoServer.Databases[$BBDD]

ForEach ($db in $server.Databases)
{
    Write-Output $db.Name

    if($db.Name -eq 'RETABet Datamodel Full')
    {

        $model = $db.Model;

        #Para pintarlas luego en orden
        #Recoge Todas las carpetas y su tipo
        $DisplayFolders = New-Object 'System.Collections.Generic.Dictionary[string,TableType]'
    
        #Para recoger los datos de las NO Medidas
        $NoMeasureMExpressions = New-Object 'System.Collections.Generic.Dictionary[string,string]'
    
        ForEach ($Table in $model.Tables) 
        {
            #Inicializamos variables de comprobación
            $Hidden = $false
            $IsSQL = $false
            $IsMeasure = $false

            #Miramos si la tabla la ve o no el usuario
            if ($Table.isHidden -eq $true) {
                $Hidden = $true
            }

            if ($Hidden -eq $false) 
            {
           
               #Categorizamos si es una tabla que mira SQL o no, para saber si ir a la BBDD correspondiente a buscar su descripción
               if ($Table.RefreshPolicy.Count -eq 0){
                #Si no tiene politicas de refresco, no tiene particiones
                    ForEach ($Partition in $Table.Partitions) 
                    {       
                        ForEach ($source in $Partition.Source)
                        {                           
                           $NoMeasureMExpressions.Add($Table.Name, $source.Expression)      

                           if ($source.Expression -like "*Sql.Database*" )
                           {
                                $IsSQL = $true
                                break
                           }
                        } 

                        if ($isSQL -eq $true) {break}
                    }
               } else {
               #Si tiene politicas de refresco hay que ir a la politica a buscar la expresión           
                    if ($Table.RefreshPolicy.SourceExpression -like "*Sql.Database*" )
                    {
                        $IsSQL = $true
                        $NoMeasureMExpressions.Add($Table.Name, $Table.RefreshPolicy.SourceExpression)
                    }
               }

               #Categorizamos si es una tabla de medidas o no
               if ($Table.Measures.Count -eq 0)
               {
                    $IsMeasure = $false
               }else{
                    $IsMeasure = $true
               }

            }       

            $TaTy = [TableType]::new($Table.name,$Hidden,$IsSQL,$IsMeasure)     

            #########################################################################################################################################################################
            ####################                                     RECOJO LAS TABLAS PARA ORDENARLAS                                                         ######################
            #########################################################################################################################################################################
        
            #Las Medidas
            if ($TaTy.IsMeasure -eq $true -and $TaTy.isHidden -eq $false)
            {
           
                #Para recoger los datos de las Medidas
                $MeasureDisplayFolders = New-Object 'System.Collections.Generic.SortedList[string,System.Collections.Generic.SortedList[string,MeasureType]]'
                $MySubdic = New-Object 'System.Collections.Generic.SortedList[string,MeasureType]'

                ForEach ($Measure in $Table.Measures)
                {
                    #Creamos el MeasureType
                    $MeTy = [MeasureType]::new($Measure.name,$Measure.expression,$Measure.description)

                    #Recogemos la carpeta donde se verá
                    $displayfolder = $Measure.displayfolder

                    if ($MeasureDisplayFolders.ContainsKey($displayfolder)){
                        $mySubdic = $MeasureDisplayFolders["$displayfolder"]
                        $r = $MeasureDisplayFolders.Remove($displayfolder)
                    }
                    else
                    {
                        $mySubdic = New-Object 'System.Collections.Generic.SortedList[string,[MeasureType]]'
                    }
                    $mySubdic.Add($MeTy.Name,$MeTy)
                    $MeasureDisplayFolders.Add($displayfolder, $mySubdic) 
                }

                #Almaceno en la lista para dibujar después
                $TaTy.List = $MeasureDisplayFolders
                $TaTy.Description = $Table.Description #Si es medida, me cojo la descripción de la tabla si existe.
            }
            elseif ($TaTy.isHidden -eq $false)
            {   
            
                 #Para recoger los datos de las NO Medidas
                $NoMeasureDisplayFolders = New-Object 'System.Collections.Generic.SortedList[string,string]'

                #Almaceno las descripciones SQL Server
                if ($IsSQL -eq $true) {                
                
                    #Recogemos las descripciones de las tablas
                    $MExpression = $NoMeasureMExpressions[$Table.Name]
                    $schema1 = $MExpression.Substring($MExpression.IndexOf('Schema="')+'Schema="'.Length)
                    $schema = $schema1.Substring(0,$schema1.IndexOf('"'))

                    $item1 =  $MExpression.Substring( $MExpression.IndexOf('Item="')+'Item="'.Length)
                    $item = $item1.Substring(0,$item1.IndexOf('"'))
                
                    $cmd ="Select 
	                            p.Value as [MS_Description]
                            from sys.tables t
		                                left join sys.extended_properties p
		 	                                on p.major_id = t.object_id
		 	                                and p.name ='MS_Description'
                            WHERE OBJECT_SCHEMA_NAME (object_id) = '$schema'
	                            and OBJECT_NAME(object_id) ='$item'"

                    $ds = $dbsql.ExecuteWithResults($cmd)
            
                    $Descrip = $ds.Tables[0].Rows[0].MS_Description

                    if ($Descrip -is [DBNull])
                    {
                        $Descrip = ''
                    }

                    $Taty.Description = $Descrip

                    #Recorremos las columnas y miramos su descripción
                    ForEach ($Column in $Table.Columns)
                    {
                        #Solo si se ve la columna
                        if ($Column.IsHidden -eq $false)
                        {
                            $ColumnName = $Column.Name.replace("'","''")
                            $Descrip = '' #Lo reinicio

                    
                            $cmd ="Select 
	                                    p.Value as [MS_Description]
                                    from sys.columns c
		                                      left join sys.extended_properties p
		 	                                        on p.major_id = c.object_id and c.column_id = p.minor_id
		 	                                        and p.name ='MS_Description'
                                    WHERE OBJECT_SCHEMA_NAME (object_id) = '$schema'
	                                    and OBJECT_NAME(object_id) ='$item'
	                                    and c.name = '$ColumnName'"
                            $ds = $dbsql.ExecuteWithResults($cmd)
            
                            $Descrip = $ds.Tables[0].Rows[0].MS_Description
            
                            if ($Descrip -is [DBNull])
                            {
                                $Descrip = ''
                            }

                            $NoMeasureDisplayFolders.Add($ColumnName, $Descrip)
                        }
                    }

                }else {
                    #No es Medida ni SQL. Simplemente añadimos la lista de las columnas si tiene
                    $Taty.Description = $Table.Description

                    ForEach ($Column in $Table.Columns)
                    {
                        #Solo si se ve la columna
                        if ($Column.IsHidden -eq $false)
                        {
                            $ColumnName = $Column.Name.replace("'","''")
                            $Descrip = '' #Lo reinicio
                            $NoMeasureDisplayFolders.Add($ColumnName, $Descrip)
                        }
                    }
                }

                #Almaceno en la lista para dibujar después ordenadamente
                $TaTy.List = $NoMeasureDisplayFolders
            }
        
            if ($TaTy.isHidden -eq $false)
            {
                $DisplayFolders.Add($Table.Name, $TaTy) #Para indicar si son o no Medidas
            }
        }

        #########################################################################################################################################################################
        ####################                                     CREACIÓN DOC                                                                              ######################
        #########################################################################################################################################################################

        #Configuración del documento Word
        #Configuración del documento Word
        #url
        $pathdocfile = "C:\Charlas\20241026_PowerBIDaysSantiago2024_DocumentaTuBIYTriunfa\3. Power BI Services"
        $today = Get-Date -Format "yyyymmmdd_hhmmss"
        $modelname = $db.Name
        $modelserver = $db.Server    

        $docfile = "$pathdocfile\$modelserver $modelname $today.docx"

        #Creamos instancia del objeto word
        $MSWord = New-Object -ComObject Word.Application
        #Lo hacemos visible
        $MSWord.Visible = $True
        #Creamos nuevo documento
        $mydoc = $MSWord.Documents.Add()

        $comObject = $mydoc.BuiltInDocumentProperties("Title")
        $binding = "System.Reflection.BindingFlags" -as [type]
        [System.__ComObject].invokemember("Value",$binding::SetProperty,$null,$comObject,"$modelserver - $modelname")

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

        #Creamos el nuevo Canva para recoger el logo y colocarla en el mismo lugar. Hay que configurar el tamaño
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

        #Ponemos el documento en Horizontal
        $mydoc.PageSetup.Orientation = [Microsoft.Office.Interop.Word.WdOrientation]::wdOrientLandscape

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

        #Empezamos a leer el $DisplayFolders
        foreach ($key in $displayFolders.Keys) {

            # Crea el título 1 con el nombre de la carpeta
            [TableType]$folder = $displayFolders[$key]
            $myText = $MSWord.Selection
            $myText.Style = $styles[0]
            $myText.TypeText($folder.Name)
            $myText.TypeParagraph()  

            if($folder.IsMeasure -eq $true)
            {
               #Es Medida
               #Si tiene descripción la añadimos
                if ($folder.Description -ne '')
                {                
                    $myText = $MSWord.Selection
                    $myText.TypeText($folder.Description)
                    $myText.TypeParagraph()  
                }


                #Recorremos las subcarpetas
                foreach ($keylist in $folder.List.Keys) {
                    if(-not($keylist -eq ''))
                    {
                        $style = 1
                        #Por defecto todos se empiezan a pintar con el estilo 1. Tenemos que detectar cuantas "\" tienen. Si es 0: estilo 3, si es 1: estilo 4, si es 2: estilo 5,...
                        if($keylist.Contains("\")){
                            $slashindex = $keylist.LastIndexOf("\") +1
                        }
                        else{
                            $slashindex = 0
                        }
                        
                        $charCount = ($keylist.ToCharArray() | Where-Object {$_ -eq '\'} | Measure-Object).Count
                        $folder2 = $keylist.Substring($slashindex)
                
                        # Crea el título X con el nombre de la carpeta donde se guarda la medida
                        $myText = $MSWord.Selection
                        $myText.Style = $styles[$style + $charCount]
                        $myText.TypeText($folder2)
                        $myText.TypeParagraph()  
                    } 
                
                    $measuresbyfolder = $($folder.List[$keylist]) #System.Collections.Generic.SortedList[string,MeasureType]

                    #Si vamos a mirar la IA necesitaré una columna más
                    if ($IA -eq $true){
                        $numCols = 4
                    }
                    else{
                        $numCols = 3
                    }
                
                    #Creamos la tabla para recoger las medidas de esta carpeta
                    $myText = $MSWord.Selection
                    $Range = @($myText.Paragraphs)[-1].Range
                    $Table = $myText.Tables.add(
                        $myText.Range #$Range
                        ,$measuresbyfolder.Keys.Count + 1 #Rows
                        ,$numCols #Columns
                        ,[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior
                        ,[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
                    )
                
                    ## Header
                    $Table.cell(1,1).range.Bold=1
                    $Table.cell(1,1).range.text = "Medida"
                    $Table.cell(1,2).range.Bold=1
                    $Table.cell(1,2).range.text = "Descripción"
                    $Table.cell(1,3).range.Bold=1
                    $Table.cell(1,3).range.text = "Expresión"

                    if ($IA -eq $true){
                        $Table.cell(1,4).range.Bold=1
                        $Table.cell(1,4).range.text = "Desc. IA"
                    }

                    #Empezamos a mirar los datos
                    $row = 2
        
                    foreach ($measurekey in $measuresbyfolder.Keys) {
                        [MeasureType]$measureinfo = $measuresbyfolder[$measurekey]

                        $Table.cell($row,1).range.Bold = 0
                        $Table.cell($row,1).range.text = $measureinfo.Name
                    
                        if ($measureinfo.Description -eq '' -and $IA -eq $true){
                            $mes_des = Get-AIDescription $measureinfo.DAXExpression "DAX"  
                            $Table.cell($row,4).range.Bold = 0
                            $Table.cell($row,4).range.text = "Sí"                        
                        }else {
                            $mes_des = $measureinfo.Description
                        }

                        $Table.cell($row,2).range.Bold = 0
                        $Table.cell($row,2).range.text = $mes_des
                        $Table.cell($row,3).range.Bold = 0
                        $Table.cell($row,3).range.text = $measureinfo.DAXExpression
        
                        $row++      
                    }
        
                    #Para continuar escribiendo fuera de la tabla y dejar una linea
                    $myText.Start= $mydoc.Content.End
                    $myText.TypeParagraph()
                }
            }else
            {
               #No es Medida

               #Si tiene descripción la añadimos
                if ($folder.Description -ne '')
                {                
                    $myText = $MSWord.Selection
                    $myText.TypeText($folder.Description)
                    $myText.TypeParagraph()  
                }

               #Si tiene MExpresion
               if ($NoMeasureMExpressions.ContainsKey($folder.Name)){
            
                   #Creamos el título 2 para indicar que son las transformación
                    $myText = $MSWord.Selection
                    $myText.Style = $styles[1]
                    $myText.TypeText("Transformación")
                    $myText.TypeParagraph()  
                
                    #Si vamos a mirar la IA necesitaré una columna más
                    if ($IA -eq $true){
                        $numCols = 2
                    }
                    else{
                        $numCols = 1
                    }  

                    #Creamos la tabla para recoger la info de esta tabla
                    $myText = $MSWord.Selection
                    $Range = @($myText.Paragraphs)[-1].Range
                    $Table = $myText.Tables.add(
                        $myText.Range #$Range
                        ,2 #Rows
                        ,$numCols #Columns
                        ,[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior
                        ,[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
                    )

                    ## Header
                    $Table.cell(1,1).range.Bold=1
                    $Table.cell(1,1).range.text = "Transformación"
                    $Table.cell(2,1).range.Bold = 0
                    $Table.cell(2,1).range.text = $NoMeasureMExpressions[$folder.Name]
                    if ($IA -eq $true){
                        $Table.cell(1,2).range.Bold=1
                        $Table.cell(1,2).range.text= "Descripción IA"
                        $Table.cell(2,2).range.Bold=0
                        $Table.cell(2,2).range.text = Get-AIDescription $Mexpression "M" 
                    }
        
                    #Para continuar escribiendo fuera de la tabla y dejar una linea
                    $myText.Start= $mydoc.Content.End
                    $myText.TypeParagraph()
               }

               #Listamos las columnas si tiene
               if($folder.List.Count -gt 0){
                   #Creamos el título 2 para indicar que son las columnas
                   $myText = $MSWord.Selection
                   $myText.Style = $styles[1]
                   $myText.TypeText("Columnas")
                   $myText.TypeParagraph()

                   #Creamos la tabla para recoger las columnas de esta tabla
                   $myText = $MSWord.Selection
                   $Range = @($myText.Paragraphs)[-1].Range
                   $Table = $myText.Tables.add(
                       $myText.Range #$Range
                       ,$folder.List.Count + 1 #Rows
                       ,2 #Columns
                       ,[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior
                       ,[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
                   )
                   ## Header
                   $Table.cell(1,1).range.Bold=1
                   $Table.cell(1,1).range.text = "Nombre"
                   $Table.cell(1,2).range.Bold=1
                   $Table.cell(1,2).range.text = "Descripción Origen"
            
                   $row = 2

                   foreach ($columkey in $folder.List.Keys) {
                        #rellenamos la celda con el nombre de la columna
                        $Table.cell($row,1).range.Bold = 0
                        $Table.cell($row,1).range.text = $columkey
                        $Table.cell($row,2).range.Bold = 0
                        $Table.cell($row,2).range.text = $folder.List[$columkey]

                        $row++

                   }
            
                   #Para continuar escribiendo fuera de la tabla y dejar una linea
                   $myText.Start= $mydoc.Content.End
                   $myText.TypeParagraph()

               }           
            }                   
        }
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
    }
}

