###############################################################################################
# Datos a rellenar
# $SQLServer --> Ip y puerto del servidor
# $BDUser --> Login para conectarse a la BBDD
# $BDUserPass --> Contraseña del login para conectarse a la BBDD
# $docfile --> Url donde se dejara el documento
###############################################################################################
#Dll que hay que instalar y buscar el path correspondiente
Add-Type -Path 'C:\Windows\Microsoft.NET\assembly\GAC_MSIL\Microsoft.SqlServer.Smo\v4.0_16.0.0.0__89845dcd8080cc91\Microsoft.SqlServer.Smo.dll'

#Configuración del servidor de BBDD
$SQLServer = ****************
$BDUser = ****************
$BDUserPass = ***************
$BBDD = *******************

$SmoServer = New-Object ('Microsoft.SqlServer.Management.Smo.Server') -argumentlist $SQLServer
$SmoServer.ConnectionContext.LoginSecure = $false
$SmoServer.ConnectionContext.set_Login($BDUser)
$SmoServer.ConnectionContext.set_Password($BDUserPass)

$db = $SmoServer.Databases[$BBDD]

#Configuración del documento Word
#url
$docfile = "C:\Charlas\20241026_PowerBIDaysSantiago2024_DocumentaTuBIYTriunfa\2. Demo SQLServer\DocumentationSaCriDB.docx"
#Creamos instancia del objeto word
$MSWord = New-Object -ComObject Word.Application
#Lo hacemos visible
$MSWord.Visible = $True
#Creamos nuevo documento
$mydoc = $MSWord.Documents.Add()

#Rellenamos algunas propiedades del documento.
#Título
$comObject = $mydoc.BuiltInDocumentProperties("Title")
$binding = "System.Reflection.BindingFlags" -as [type]
[System.__ComObject].invokemember("Value",$binding::SetProperty,$null,$comObject,"Documentacion SaCriDB")

#Compañia - Empresa
$comObject = $mydoc.BuiltInDocumentProperties("Company")
$binding = "System.Reflection.BindingFlags" -as [type]
[System.__ComObject].invokemember("Value",$binding::SetProperty,$null,$comObject,"Power BI Days Santiago 2024")

#Autor
$comObject = $mydoc.BuiltInDocumentProperties("Author")
$binding = "System.Reflection.BindingFlags" -as [type]
[System.__ComObject].invokemember("Value",$binding::SetProperty,$null,$comObject,"Cristina Tarabini-Castellani") 

#Preparamos la portada
$CoverPage = 'Movimiento' #Nombre de la plantilla de la portada
$Selection = $MSWord.application.selection   

#Se coge la plantilla de la portada elegida
$MSWord.Templates.LoadBuildingBlocks()
$bb =$MSWord.templates | Where-Object -Property name -EQ -Value 'Built-In Building Blocks.dotx'

$part = $bb.BuildingBlockEntries.item($CoverPage)

#Preparamos la portada
$CoverPage = 'Movimiento' #Nombre de la plantilla de la portada
$Selection = $MSWord.application.selection   

#Se coge la plantilla de la portada elegida
$MSWord.Templates.LoadBuildingBlocks()
$bb =$MSWord.templates | Where-Object -Property name -EQ -Value 'Built-In Building Blocks.dotx'

$part = $bb.BuildingBlockEntries.item($CoverPage)

#Tratamos la plantilla de portada elegida. 
######################### Esto es para la portada "Movimiento"
$null = $part.Insert($MSWord.Selection.range, $true)   

#Las 3 líneas verdes verticales. Les voy a cambiar el color a Amarillo :)
$Selection.Document.Shapes.Item("Grupo 252").GroupItems[1].Fill.ForeColor.Rgb = 2551538 #24319208 ##ff00a7
#El rectángulo grande verde. También lo cambio a Rosa                                 
$Selection.Document.Shapes.Item("Grupo 252").GroupItems[2].Fill.ForeColor.Rgb = 2551538 #24319208 ##ff00a7

#Borramos la imagen por defecto de la portada. No he sido capaz de sustituirla
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

$Selection.Document.Shapes.Item("Grupo 252").GroupItems[3].TextFrame.TextRange.Font.Color = 0 #En negro las letras
$Selection.Document.Shapes.Item("Grupo 252").GroupItems[3].TextFrame.TextRange.Text= $año #Año

#Reemplazamos la etiqueta [Fecha]
$fecha = Get-Date -Format "dd/MM/yyyy"
$Texto = $Selection.Document.Shapes.Item("Grupo 252").GroupItems[4].TextFrame.TextRange.Text
$Selection.Document.Shapes.Item("Grupo 252").GroupItems[4].TextFrame.TextRange.Text = $Texto.Replace("[Fecha]",$fecha) 
$Selection.Document.Shapes.Item("Grupo 252").GroupItems[4].TextFrame.TextRange.Font.Color = 0 #En negro las letras

#########################Fin portada "Movimiento"

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

#Configuramos la apariencia del documento
$Section = $mydoc.Sections.Item(1)
$Header = $Section.Headers.Item(1)
$Footer = $Section.Footers.Item(1)
$Footer.PageNumbers.Add()

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

#Creamos una referencia al documento actual, para añadirle el texto
$myText = $MSWord.Selection
$myText.Style = $styles[0]
$myText.TypeText("Documentation End-to-End")
$myText.TypeParagraph()

$myText = $MSWord.Selection
$date = Get-date -Format "dddd dd/MM/yyyy HH:mm K"
$myText.Font.Bold = 1
$myText.TypeText("Generation date: ")
$myText.Font.Bold = 0
$myText.TypeText($date)
$myText.TypeParagraph()

#Recorremos la BBDD
#Creamos el título 1 para indicar que la BBDD
$myText = $MSWord.Selection
$myText.Style = $styles[0]
$myText.TypeText($db.Name)
$myText.TypeParagraph()    

if ($db.tables.Count -gt 0) {
    #Creamos el título 2 para indicar que van a escribirse tablas
    $myText = $MSWord.Selection
    $myText.Style = $styles[1]
    $myText.TypeText("Tablas")
    $myText.TypeParagraph()  
}

foreach ($ta in $db.Tables) {         
    #Creamos el título 3 para indicar que van a escribirse una tabla
    $myText = $MSWord.Selection
    $myText.Style = $styles[2]
    $myText.TypeText($ta)
    $myText.TypeParagraph()  

    #Recojo el nombre de la tabla
    $tablename = $ta.Name 
    $tableschema = $ta.schema

    $cmd ="SELECT
	            p.name
	            ,p.value
            FROM sys.tables t
		            left join sys.extended_properties p
			            on p.major_id = t.object_id	and p.minor_id = 0 /*tabla*/
            WHERE t.name = '$tablename' and t.schema_id = SCHEMA_ID('$tableschema')"
           
    #write-host $cmd

    $ds = $db.ExecuteWithResults($cmd)

    #Creamos la tabla para recoger las propiedades de esta tabla si existen
    if ($ds.Tables[0].Rows.Count -gt 0) {
        $myText = $MSWord.Selection
        $Range = @($myText.Paragraphs)[-1].Range
        $Table = $myText.Tables.add(
            $myText.Range #$Range
            ,$ds.Tables[0].Rows.Count + 1 #Rows
            ,2 #Columns
            ,[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior
            ,[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
        )
        ## Header tabla propiedades tablas
        $Table.cell(1,1).range.Bold=1
        $Table.cell(1,1).range.text = "Propiedad"
        $Table.cell(1,2).range.Bold=1
        $Table.cell(1,2).range.text = "Valor"

        $row = 2

        for($j=0; $j -lt $ds.Tables[0].Rows.Count; $j++){
            $name = $ds.Tables[0].Rows[$j].name

            if ($name -is [DBNull])
            {
                $name = ''
            }
            
            $value = $ds.Tables[0].Rows[$j].value
            
            if ($value -is [DBNull])
            {
                $value = ''
            }
            
            $Table.cell($row,1).range.Bold = 0
            $Table.cell($row,1).range.text = "$name"
            $Table.cell($row,2).range.Bold = 0
            $Table.cell($row,2).range.text = "$value"
        
            $row++     
        }
        
        #Para continuar escribiendo fuera de la tabla y dejar una linea
        $myText.Start= $mydoc.Content.End
        $myText.TypeParagraph()
    }
               
    #Creamos el título 3 para indicar que van a escribirse las columnas
    $myText = $MSWord.Selection
    $myText.Style = $styles[3]
    $myText.TypeText("Columnas")
    $myText.TypeParagraph()  

    #write-host "Columnas"

    #Recogemos la info de las columnas
    $cmd="DECLARE @sql nvarchar(MAX);
 
            SET @sql = N'
 
            SELECT
            * 
            FROM
            ( 
			Select 
		        c.column_id
		        ,''COLUMN'' as [Propiedad]
		        ,c.name as [Valor]
	        from sys.columns c
			WHERE OBJECT_SCHEMA_NAME (object_id) = ''$tableschema'' and OBJECT_NAME (object_id) = ''$tablename''
				
			UNION ALL
 
	        Select 
		        c.column_id
		        ,p.name as [Propiedad]
		        ,p.Value as [Valor]
	        from sys.columns c
				        left join sys.extended_properties p
					        on p.major_id = c.object_id and c.column_id = p.minor_id
	        WHERE OBJECT_SCHEMA_NAME (object_id) = ''$tableschema'' and OBJECT_NAME (object_id) = ''$tablename''
            ) AS T
            PIVOT   
            (
	        min([Valor])
            FOR [Propiedad] IN (' + (SELECT STUFF(
            (
            SELECT
            ',' + QUOTENAME(LTRIM(Propiedad))
            FROM
            (
			SELECT 'COLUMN' as [Propiedad]
			UNION ALL
            SELECT DISTINCT p.name as [Propiedad]
	        from sys.columns c
				        left join sys.extended_properties p
					        on p.major_id = c.object_id and c.column_id = p.minor_id
	        WHERE OBJECT_SCHEMA_NAME (object_id) = '$tableschema' and OBJECT_NAME (object_id) = '$tablename'
            ) AS T
            ORDER BY
            [Propiedad]
            FOR XML PATH('')
            ), 1, 1, '')) + N')
            ) AS P
            order by column_id;'; 
 
        --En la variable @sql tenemos la consulta completa 
            EXEC sp_executesql @sql;"
        
    

    $ds = $db.ExecuteWithResults($cmd)

    #Creamos la tabla para recoger las columnas de esta tabla si existen resultados
    if ($ds.Tables[0].Rows.Count -gt 0) {
        $myText = $MSWord.Selection
        $Range = @($myText.Paragraphs)[-1].Range
        $Table = $myText.Tables.add(
            $myText.Range #$Range
            ,$ds.Tables[0].Rows.Count + 1 #Rows + Header
            ,$ds.Tables[0].Columns.count -1 #Columns - Column_id
            ,[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior
            ,[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
        )
        
        $numCols = $ds.Tables[0].Columns.Count

        ## Header de las columnas. Ignoraremos el Column_id (es el 0)
        ## Header tabla propiedades tablas
        $Table.cell(1,1).range.Bold=1
        $Table.cell(1,1).range.text = "Nombre Columna"

        for($j=1; $j -lt $numCols; $j++){        
            $Table.cell(1,$j).range.Bold=1
            $propiedad = $ds.Tables[0].Columns[$j].ColumnName
            $Table.cell(1,$j).range.text = "$propiedad"
        }
        
        $row = 2

        #Recorremos ya los resultados de las columnas
        for($j=0; $j -lt $ds.Tables[0].Rows.Count; $j++){
            $columnname = $ds.Tables[0].Rows[$j].name
                        
            if ($columnname -is [DBNull])
            {
                $columnname = ''
            }
                        
            $Table.cell($row,1).range.Bold = 0
            $Table.cell($row,1).range.text = "$columnname"

            for($i=1; $i -lt $numCols; $i++ ){
                $columnvalue = $ds.Tables[0].Rows[$j].ItemArray[$i]

                if ($columnvalue -is [DBNull])
                {
                    $columnvalue = ''
                }

                $Table.cell($row,$i).range.Bold = 0
                $Table.cell($row,$i).range.text = "$columnvalue"

            }

            $row++     
        }

        #Para continuar escribiendo fuera de la tabla y dejar una linea
        $myText.Start= $mydoc.Content.End
        $myText.TypeParagraph()
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
