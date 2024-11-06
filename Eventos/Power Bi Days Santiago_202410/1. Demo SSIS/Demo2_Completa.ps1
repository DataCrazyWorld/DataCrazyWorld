#Preparando la info sobre los .dtsx
#Pon aquí tu url
$path = "C:\Charlas\20241026_PowerBIDaysSantiago2024_DocumentaTuBIYTriunfa\1. Demo SSIS\"
$extension = ".dtsx"

#Preparando la generación del documento Word
$docfile = "C:\Charlas\20241026_PowerBIDaysSantiago2024_DocumentaTuBIYTriunfa\1. Demo SSIS\DocumentacionDTSX.docx"


# Create a new instance/object of MS Word
$MSWord = New-Object -ComObject Word.Application

# Make MS Word visible
$MSWord.Visible = $True

# Add a new document
$mydoc = $MSWord.Documents.Add()


$Section = $mydoc.Sections.Item(1)
$Header = $Section.Headers.Item(1)
$Footer = $Section.Footers.Item(1)
$Footer.PageNumbers.Add()

# Create a reference to the current document so we can begin adding text
$myText = $MSWord.Selection
$myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading1
$myText.TypeText("Documentation Generation Info")
$myText.TypeParagraph()

$myText = $MSWord.Selection
$date = Get-date -Format "dddd dd/MM/yyyy HH:mm K"
$myText.Font.Bold = 1
$myText.TypeText("Generation date: ")
$myText.Font.Bold = 0
$myText.TypeText($date)
$myText.TypeParagraph()

#Me voy a recorrer todos los dtsx para generar la documentación
Get-ChildItem $path |
Foreach-Object {
   If ($_.Extension -eq $extension)
   {
         # read .dtsx into XML variable
        [xml] $myxml = Get-Content $_.FullName

        # Crea el título 1
        $myText = $MSWord.Selection
        $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading1
        $myText.TypeText($_.Name)
        $myText.TypeParagraph()
        
        #Para leer con prefijos DTS y SQLTask tengo que indicar los namespaces
        $XmlNamespace = @{ DTS = 'www.microsoft.com/SqlServer/Dts';  SQLTask = 'www.microsoft.com/sqlserver/dts/tasks/sqltask'};

        #Variables del paquete        
        $HasAlreadyPutHeading2 = $false

        $myxml | Select-XML -XPath "//DTS:Variable" -Namespace $XmlNamespace| ForEach-Object {
            # Crea el título 2 si no existe
            if ($HasAlreadyPutHeading2 -eq $false)
            {
                $myText = $MSWord.Selection
                $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading2
                $myText.TypeText("Variables")
                $myText.TypeParagraph()     
                $HasAlreadyPutHeading2 = $true    
            }

            # Crea el título 3 con el nombre de la variable
            $myText = $MSWord.Selection
            $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading4
            $myText.TypeText($_.Node.ObjectName)
            $myText.TypeParagraph()    
            
            # Si existe expresión
            if ($_.Node.Expression) {
                $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
                $myText.Font.Bold = 1
                $myText.TypeText('Expresión: ')
                $myText.TypeParagraph()             
                $myText.Font.Bold = 0
                $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStylePlainText
                $myText.TypeText($_.Node.Expression)
                $myText.TypeParagraph()  
                $myText.TypeParagraph()   
            }   
              
            #Recojo los hijos
            [xml] $childs = $_.Node.OuterXml 
          
            $childs| Select-XML -XPath "//DTS:VariableValue"  -Namespace $XmlNamespace | ForEach-Object {
                if ($_.Node.InnerText) {
                    # Registro el comando SQL
                    $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
                    $myText.Font.Bold = 1
                    $myText.TypeText('Valor de la Variable: ')
                    $myText.TypeParagraph()            
                    $myText.Font.Bold = 0
                    $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStylePlainText
                    $myText.TypeText($_.Node.InnerText)
                    $myText.TypeParagraph()  
                    $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStylePlainText  
                }                 
            }     

        }
       
       #Objets SQLTaskData    
       $HasAlreadyPutHeading2 = $false 
       $myxml | Select-XML -XPath "//DTS:Executable[@DTS:ExecutableType='Microsoft.ExecuteSQLTask']" -Namespace $XmlNamespace| ForEach-Object {
            # Crea el título 2 si no existe
            if ($HasAlreadyPutHeading2 -eq $false)
            {
                $myText = $MSWord.Selection
                $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading2
                $myText.TypeText("SqlTaskData")
                $myText.TypeParagraph()     
                $HasAlreadyPutHeading2 = $true    
            }  

            # Crea el título 3 con el nombre de la TaskData
            $myText = $MSWord.Selection
            $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading3
            $myText.TypeText($_.Node.ObjectName)
            $myText.TypeParagraph()   
            
            #Recojo los hijos, que son cada uno de los nodos Property
            [xml] $childs = $_.Node.OuterXml  
            
            $myText = $MSWord.Selection            
        
            $childs| Select-XML -XPath "//SQLTask:SqlTaskData[@SQLTask:SqlStatementSource]" -Namespace $XmlNamespace  | ForEach-Object {
                # Registro el statement
                $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStylePlainText
                $myText.Font.Bold = 1
                $myText.TypeText('SqlStatementSource: ')
                $myText.TypeParagraph()                     
                $myText.Font.Bold = 0
                $myText.TypeText($_.Node.SqlStatementSource)
                $myText.TypeParagraph()   
                $myText.TypeParagraph()                     
            }   
              
        }
       
      # Objeto OLEDB Source
      $HasAlreadyPutHeading2 = $false
      
      $myxml | Select-XML -XPath "//component[@componentClassID='Microsoft.OLEDBSource']" | ForEach-Object {
          # Crea el título 2 si no existe
          if ($HasAlreadyPutHeading2 -eq $false)
          {
              $myText = $MSWord.Selection
              $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading2
              $myText.TypeText("Microsoft.OLEDBSource")
              $myText.TypeParagraph()     
              $HasAlreadyPutHeading2 = $true    
          }
      
          # Crea el título 3 si no existe
          $myText = $MSWord.Selection
          $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading3
          $myText.TypeText($_.Node.Name)
          $myText.TypeParagraph()            
      
          #Recojo los hijos, que son cada uno de los nodos Property
          [xml] $childs = $_.Node.OuterXml        
          
          $myText = $MSWord.Selection            
      
          $childs| Select-XML -XPath "//connection[@connectionManagerRefId]"  | ForEach-Object {
              # Registro el nombre de la connexión
              $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
              $myText.Font.Bold = 1
              $myText.TypeText('Conexión: ')
              $myText.Font.Bold = 0
              $myText.TypeText($_.Node.connectionManagerRefId)
              $myText.TypeParagraph()                     
          }
               
          #Miro si tiene SQL Command de Variable, porque si tiene, no tengo escribir el SQLCommand. 
          #Inicializo la variable de comprobación
          $UseSQLCmdVble = $false  
          $childs| Select-XML -XPath "//property[@name='SqlCommandVariable']"  | ForEach-Object {
              if ($_.Node.InnerText) {
                  # Registro el comando SQL
                  $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
                  $myText.Font.Bold = 1
                  $myText.TypeText('Comando SQL en la Variable: ')
                  $myText.Font.Bold = 0
                  $myText.TypeText($_.Node.InnerText)
                  $myText.TypeParagraph()    
                  $UseSQLCmdVble = $true     
                 }                        
          }
          
          $childs| Select-XML -XPath "//property[@name='OpenRowsetVariable']"  | ForEach-Object {
              if ($_.Node.InnerText) {
                  # Registro el comando SQL
                  $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
                  $myText.Font.Bold = 1
                  $myText.TypeText('Comando SQL en la Variable: ')
                  $myText.Font.Bold = 0
                  $myText.TypeText($_.Node.InnerText)
                  $myText.TypeParagraph()    
                  $UseSQLCmdVble = $true     
                 }                        
          }
      
          if (-not $UseSQLCmdVble){
      
              $childs| Select-XML -XPath "//property[@name='SqlCommand']"  | ForEach-Object {
                  # Registro el comando SQL
                  if ($_.Node.InnerText)
                  {
                      $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
                      $myText.Font.Bold = 1
                      $myText.TypeText('Comando SQL: ')
                      $myText.TypeParagraph()             
                      $myText.Font.Bold = 0
                      $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStylePlainText
                      $myText.TypeText($_.Node.InnerText)
                      $myText.TypeParagraph()    
                   }                             
              }                          
      
              $childs| Select-XML -XPath "//property[@name='OpenRowset']"  | ForEach-Object {
                  # Registro el comando SQL
                  if ($_.Node.InnerText)
                  {
                      $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
                      $myText.Font.Bold = 1
                      $myText.TypeText('Tabla origen: ')
                      $myText.Font.Bold = 0
                      $myText.TypeText($_.Node.InnerText)
                      $myText.TypeParagraph()      
                   }                                                        
              }
          }
      
      }
      
      # Objeto Lookup
      $HasAlreadyPutHeading2 = $false
      
      $myxml | Select-XML -XPath "//component[@componentClassID='Microsoft.Lookup']" | ForEach-Object {
          # Crea el título 2 si no existe
          if ($HasAlreadyPutHeading2 -eq $false)
          {
              $myText = $MSWord.Selection
              $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading2
              $myText.TypeText("Microsoft.Lookup")
              $myText.TypeParagraph()     
              $HasAlreadyPutHeading2 = $true    
          }
      
          # Crea el título 3 si no existe
          $myText = $MSWord.Selection
          $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading3
          $myText.TypeText($_.Node.Name)
          $myText.TypeParagraph()            
      
          #Recojo los hijos
          [xml] $childs = $_.Node.OuterXml 
          
          $childs| Select-XML -XPath "//connection[@connectionManagerRefId]"  | ForEach-Object {
              # Registro el nombre de la connexión
              $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
              $myText.Font.Bold = 1
              $myText.TypeText('Conexión: ')
              $myText.Font.Bold = 0
              $myText.TypeText($_.Node.connectionManagerRefId)
              $myText.TypeParagraph()                     
          }
      
          #Miro si tiene SQL Command de Variable, porque si tiene, no tengo escribir el SQLCommand. 
          #Inicializo la variable de comprobación
          $UseSQLCmdVble = $false  
          $childs| Select-XML -XPath "//property[@name='SqlCommandVariable']"  | ForEach-Object {
              if ($_.Node.InnerText) {
                  # Registro el comando SQL
                  $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
                  $myText.Font.Bold = 1
                  $myText.TypeText('Comando SQL en la Variable: ')
                  $myText.Font.Bold = 0
                  $myText.TypeText($_.Node.InnerText)
                  $myText.TypeParagraph()    
                  $UseSQLCmdVble = $true     
                 }                        
          }
      
          if (-not $UseSQLCmdVble){
      
              $childs| Select-XML -XPath "//property[@name='SqlCommand']"  | ForEach-Object {
                  # Registro el comando SQL
                  $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
                  $myText.Font.Bold = 1
                  $myText.TypeText('Comando SQL: ')
                  $myText.TypeParagraph()             
                  $myText.Font.Bold = 0
                  $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStylePlainText
                  $myText.TypeText($_.Node.InnerText)
                  $myText.TypeParagraph()                                 
              }
          }
      
          
      }
      
      # Objeto Comando OLEDB
      $HasAlreadyPutHeading2 = $false
      
      $myxml | Select-XML -XPath "//component[@componentClassID='Microsoft.OLEDBCommand']" | ForEach-Object {
          # Crea el título 2 si no existe
          if ($HasAlreadyPutHeading2 -eq $false)
          {
              $myText = $MSWord.Selection
              $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading2
              $myText.TypeText("Microsoft.OLEDBCommand")
              $myText.TypeParagraph()     
              $HasAlreadyPutHeading2 = $true    
          }
      
          # Crea el título 3 si no existe
          $myText = $MSWord.Selection
          $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading3
          $myText.TypeText($_.Node.Name)
          $myText.TypeParagraph()            
      
          #Recojo los hijos
          [xml] $childs = $_.Node.OuterXml 
          
          $childs| Select-XML -XPath "//connection[@connectionManagerRefId]"  | ForEach-Object {
              # Registro el nombre de la connexión
              $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
              $myText.Font.Bold = 1
              $myText.TypeText('Conexión: ')
              $myText.Font.Bold = 0
              $myText.TypeText($_.Node.connectionManagerRefId)
              $myText.TypeParagraph()                     
          }
      
          #Miro si tiene SQL Command de Variable, porque si tiene, no tengo escribir el SQLCommand. 
          #Inicializo la variable de comprobación
          $UseSQLCmdVble = $false  
          $childs| Select-XML -XPath "//property[@name='SqlCommandVariable']"  | ForEach-Object {
              if ($_.Node.InnerText) {
                  # Registro el comando SQL
                  $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
                  $myText.Font.Bold = 1
                  $myText.TypeText('Comando SQL en la Variable: ')
                  $myText.Font.Bold = 0
                  $myText.TypeText($_.Node.InnerText)
                  $myText.TypeParagraph()    
                  $UseSQLCmdVble = $true     
                 }                        
          }
      
          if (-not $UseSQLCmdVble){
      
              $childs| Select-XML -XPath "//property[@name='SqlCommand']"  | ForEach-Object {
                  # Registro el comando SQL
                  $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
                  $myText.Font.Bold = 1
                  $myText.TypeText('Comando SQL: ')
                  $myText.TypeParagraph()             
                  $myText.Font.Bold = 0
                  $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStylePlainText
                  $myText.TypeText($_.Node.InnerText)
                  $myText.TypeParagraph()                                 
              }
          }    
          
      }

      # Objeto OLEDB Destination
      $HasAlreadyPutHeading2 = $false
      
      $myxml | Select-XML -XPath "//component[@componentClassID='Microsoft.OLEDBDestination']" | ForEach-Object {
          # Crea el título 2 si no existe
          if ($HasAlreadyPutHeading2 -eq $false)
          {
              $myText = $MSWord.Selection
              $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading2
              $myText.TypeText("Microsoft.OLEDBDestination")
              $myText.TypeParagraph()     
              $HasAlreadyPutHeading2 = $true    
          }
      
          # Crea el título 3 si no existe
          $myText = $MSWord.Selection
          $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleHeading3
          $myText.TypeText($_.Node.Name)
          $myText.TypeParagraph()            
      
          #Recojo los hijos, que son cada uno de los nodos Property
          [xml] $childs = $_.Node.OuterXml        
          
          $myText = $MSWord.Selection            
      
          $childs| Select-XML -XPath "//connection[@connectionManagerRefId]"  | ForEach-Object {
              # Registro el nombre de la connexión
              $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
              $myText.Font.Bold = 1
              $myText.TypeText('Conexión: ')
              $myText.Font.Bold = 0
              $myText.TypeText($_.Node.connectionManagerRefId)
              $myText.TypeParagraph()                     
          }
          
               
          #Miro si tiene SQL Command de Variable, porque si tiene, no tengo escribir el SQLCommand. 
          #Inicializo la variable de comprobación
          $UseSQLCmdVble = $false  
          $childs| Select-XML -XPath "//property[@name='SqlCommandVariable']"  | ForEach-Object {
              if ($_.Node.InnerText) {
                  # Registro el comando SQL
                  $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
                  $myText.Font.Bold = 1
                  $myText.TypeText('Comando SQL en la Variable: ')
                  $myText.Font.Bold = 0
                  $myText.TypeText($_.Node.InnerText)
                  $myText.TypeParagraph()    
                  $UseSQLCmdVble = $true     
                 }                        
          }
          
          $childs| Select-XML -XPath "//property[@name='OpenRowsetVariable']"  | ForEach-Object {
              if ($_.Node.InnerText) {
                  # Registro el comando SQL
                  $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
                  $myText.Font.Bold = 1
                  $myText.TypeText('Comando SQL en la Variable: ')
                  $myText.Font.Bold = 0
                  $myText.TypeText($_.Node.InnerText)
                  $myText.TypeParagraph()    
                  $UseSQLCmdVble = $true     
                 }                        
          }
      
          if (-not $UseSQLCmdVble){
      
              $childs| Select-XML -XPath "//property[@name='SqlCommand']"  | ForEach-Object {
                  # Registro el comando SQL
                  if ($_.Node.InnerText)
                  {
                      $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
                      $myText.Font.Bold = 1
                      $myText.TypeText('Comando SQL: ')
                      $myText.TypeParagraph()             
                      $myText.Font.Bold = 0
                      $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStylePlainText
                      $myText.TypeText($_.Node.InnerText)
                      $myText.TypeParagraph()    
                   }                             
              }                          
      
              $childs| Select-XML -XPath "//property[@name='OpenRowset']"  | ForEach-Object {
                  # Registro la tabla
                  if ($_.Node.InnerText)
                  {
                      $myText.Style = [Microsoft.Office.Interop.Word.WdBuiltinStyle]::wdStyleListBullet
                      $myText.Font.Bold = 1
                      $myText.TypeText('Tabla destino: ')
                      $myText.Font.Bold = 0
                      $myText.TypeText($_.Node.InnerText)
                      $myText.TypeParagraph()   
                  }                              
              }
          }
      
      }
   }
}

$saveFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault
$mydoc.SaveAs([ref][system.object]$docfile, [ref]$saveFormat)
$mydoc.Close()
$MSWord.Quit()

# Clean up Com object
$null =
[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSWord)
Remove-Variable MSWord