###############################################################################################
# Datos a rellenar
# $docfile --> Url donde se dejara el documento
###############################################################################################

#Configuración del documento Word
#url
$docfile = "C:\Test\EjemDoc.docx"
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
[System.__ComObject].invokemember("Value",$binding::SetProperty,$null,$comObject,"Ejemplo de doc con PowerShell")

#Compañia - Empresa
$comObject = $mydoc.BuiltInDocumentProperties("Company")
$binding = "System.Reflection.BindingFlags" -as [type]
[System.__ComObject].invokemember("Value",$binding::SetProperty,$null,$comObject,"Data Crazy World")

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

#Tratamos la plantilla de portada elegida. 
######################### Esto es para la portada "Movimiento"
$null = $part.Insert($MSWord.Selection.range, $true)   

#Las 3 líneas verdes verticales. Les voy a cambiar el color a Rosa :)
$Selection.Document.Shapes.Item("Grupo 78").GroupItems[1].Fill.ForeColor.Rgb = 16711847 ##ff00a7
#El rectángulo grande verde. También lo cambio a Rosa                                 
$Selection.Document.Shapes.Item("Grupo 78").GroupItems[2].Fill.ForeColor.Rgb = 16711847 ##ff00a7

#Borramos la imagen por defecto de la portada. No he sido capaz de sustituirla
$Selection.Document.Shapes.Item('Imagen 1').Delete()

#Creamos el nuevo Canva para recoger el logo y colocarla en el mismo lugar. Hay que configurar el tamaño
$null = $Selection.Document.Shapes.AddCanvas(1, 1, 200, 107.5)
$null = $Selection.Document.Shapes.Item('Lienzo 1').CanvasItems.AddPicture("C:\DataCrazyWorld\Logos DataCrazyWorld\Elementos\Logos\Color_Small.png", $false, $True, $null, $null, 200, 107.5)#439.2, 295.55
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
$Selection.Document.Shapes.Item("Grupo 78").GroupItems[3].TextFrame.TextRange.Text= $año #Año

#Reemplazamos la etiqueta [Fecha]
$fecha = Get-Date -Format "dd/MM/yyyy"
$Texto = $Selection.Document.Shapes.Item("Grupo 78").GroupItems[4].TextFrame.TextRange.Text
$Selection.Document.Shapes.Item("Grupo 78").GroupItems[4].TextFrame.TextRange.Text = $Texto.Replace("[Fecha]",$fecha) 

#########################Fin portada "Movimiento"

#Metemos un salto de página para incluir la tabla de contenidos
$Selection = $MSWord.Selection
$Selection.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdSectionBreakNextPage)
$Selection = $MSWord.Selection
$range = $Selection.Range	
$toc = $mydoc.TablesOfContents.Add($range)
$Selection.TypeParagraph()

#Metemos un salto de página para empezar a documentar
$Selection.InsertBreak([Microsoft.Office.Interop.Word.WdBreakType]::wdSectionBreakNextPage)

#Creamos array de stilos para las cabeceras. Esto es opcional... me sirve para acortar luego los comandos
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

#Creamos una referencia al documento actual, para añadirle el texto
$myText = $MSWord.Selection

#Ponemos el Titulo1
$myText.Style = $styles[0] #Titulo 1
$myText.TypeText("Ejemplo de documento generado por PowerShell")
$myText.TypeParagraph()

#Seguimos rellenando con texto normal
$myText.TypeText("Rellenas el contenido que quieras")
$myText.TypeParagraph()

#Ponemos el Titulo2
$myText.Style = $styles[1] #Titulo 2
$myText.TypeText("Ejemplo título 2")
$myText.TypeParagraph()

#Hacemos una lista
$myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
$myText.Font.Bold = 1
$myText.TypeText('Tit. Lista: ')
$myText.TypeParagraph()             
$myText.Font.Bold = 0
$myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStylePlainText
$myText.TypeText('Detalle lista')
$myText.TypeParagraph()  

#Dejamos libre una línea
$myText.TypeParagraph()

#Creamos una tabla
$Range = @($myText.Paragraphs)[-1].Range
$Table = $myText.Tables.add(
    $myText.Range #$Range
    ,2 #Rows
    ,3 #Columns
    ,[Microsoft.Office.Interop.Word.WdDefaultTableBehavior]::wdWord9TableBehavior
    ,[Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitContent
)
## Cabecera de la Tabla
$Table.cell(1,1).range.Bold=1
$Table.cell(1,1).range.text = "Columna 1"
$Table.cell(1,2).range.Bold=1
$Table.cell(1,2).range.text = "Columna 2"
$Table.cell(1,3).range.Bold=1
$Table.cell(1,3).range.text = "Columna 3"

#rellenamos la celda con la descripción de la columna cogida de BBDD
$Table.cell(2,1).range.Bold = 0
$Table.cell(2,1).range.text = 'Cont. Col1'
$Table.cell(2,2).range.Bold = 0
$Table.cell(2,2).range.text = 'Cont. Col2'
$Table.cell(2,3).range.Bold = 0
$Table.cell(2,3).range.text = 'Cont. Col3'

#Para continuar escribiendo fuera de la tabla y dejar una linea
$myText.Start= $mydoc.Content.End
$myText.TypeParagraph()

#Actualizamos el índice	
$toc.Update()

#Salvamos el documento Word
$saveFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault
$mydoc.SaveAs([ref][system.object]$docfile, [ref]$saveFormat)
$mydoc.Close()
$MSWord.Quit()

# Limpiar el objeto Com
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSWord)
Remove-Variable MSWord