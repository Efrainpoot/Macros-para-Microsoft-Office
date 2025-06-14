Sub ImprimirPdfSeparadosv4()
    ' Declaración de Variables
    Dim paginasDocumento As Integer
    Dim totalPaginas As Integer
    Dim pagActual As Integer
    Dim carpeta As String
    Dim nombreDocs As String
    Dim miRango As Range
    Dim nombreArchivo As String
    Dim numeroDoc As Integer
    
    On Error GoTo ErrorHandler
    
    ' Verificar que hay un documento activo
    If ActiveDocument Is Nothing Then
        MsgBox "No hay ningún documento abierto.", vbExclamation
        Exit Sub
    End If
    
    ' Asignación de valores a variables
    totalPaginas = ActiveDocument.Range.Information(wdNumberOfPagesInDocument)
    
    ' Validar entrada del usuario para páginas por documento
    Do
        paginasDocumento = Val(InputBox("¿Cuántas páginas tiene cada documento?", "Número de páginas", "1"))
        If paginasDocumento <= 0 Then
            MsgBox "Por favor ingrese un número válido mayor a 0.", vbExclamation
        End If
    Loop Until paginasDocumento > 0
    
    ' Seleccionar carpeta destino con explorador
    carpeta = SeleccionarCarpeta()
    If carpeta = "" Then
        MsgBox "No se seleccionó ninguna carpeta. Operación cancelada.", vbInformation
        Exit Sub
    End If
    
    ' Asegurar que la carpeta termine con backslash
    If Right(carpeta, 1) <> "\" Then carpeta = carpeta & "\"
    
    ' Calcular cuántos documentos se generarán
    Dim totalDocumentos As Integer
    totalDocumentos = Int((totalPaginas - 1) / paginasDocumento) + 1
    
    ' Obtener lista de nombres personalizados
    Dim listaNombres() As String
    Dim usarNombresPersonalizados As Boolean
    usarNombresPersonalizados = ObtenerListaNombres(totalDocumentos, listaNombres)
    
    ' Nombre base por defecto si no se usan nombres personalizados
    If Not usarNombresPersonalizados Then
        nombreDocs = InputBox("¿Qué nombre base tendrán los documentos?", "Nombre documentos", "misdocs")
        If nombreDocs = "" Then
            If MsgBox("¿Desea cancelar la operación?", vbYesNo + vbQuestion, "Confirmar cancelación") = vbYes Then
                Exit Sub
            Else
                nombreDocs = "documento" ' Valor por defecto
            End If
        End If
    End If

    pagActual = 1
    numeroDoc = 1
    
    ' Proceso principal
    Do While pagActual <= totalPaginas
        ' Buscar patrón para nombre personalizado (opcional)
        Set miRango = ActiveDocument.Content
        miRango.SetRange Start:=ActiveDocument.Range(ActiveDocument.GoTo(wdGoToPage, wdGoToAbsolute, pagActual).Start, _
                                                     ActiveDocument.GoTo(wdGoToPage, wdGoToAbsolute, pagActual + paginasDocumento - 1).End).Start, _
                      End:=ActiveDocument.Range(ActiveDocument.GoTo(wdGoToPage, wdGoToAbsolute, pagActual).Start, _
                                               ActiveDocument.GoTo(wdGoToPage, wdGoToAbsolute, pagActual + paginasDocumento - 1).End).End
        
        ' Definir nombre del archivo
        If usarNombresPersonalizados Then
            nombreArchivo = LimpiarNombreArchivo(listaNombres(numeroDoc - 1))
        Else
            nombreArchivo = nombreDocs & "_" & Format(numeroDoc, "000")
        End If
        
        ' Buscar patrón personalizado (código original mejorado) - solo si no se usan nombres personalizados
        If Not usarNombresPersonalizados Then
            miRango.Find.ClearFormatting
            With miRango.Find
                .Text = "(_)*(-)"
                .MatchWildcards = True
                .Forward = True
                .Wrap = wdFindStop
            End With
            
            If miRango.Find.Execute Then
                Dim textoEncontrado As String
                textoEncontrado = Trim(miRango.Text)
                textoEncontrado = Replace(textoEncontrado, "_", "")
                textoEncontrado = Replace(textoEncontrado, "-", "")
                textoEncontrado = LimpiarNombreArchivo(textoEncontrado)
                
                If Len(textoEncontrado) > 0 Then
                    nombreArchivo = textoEncontrado & "_" & Format(numeroDoc, "000")
                End If
                
                ' Hacer el texto bold pero no eliminarlo
                miRango.Font.Bold = True
            End If
        End If
        
        ' Calcular página final (no exceder el total)
        Dim paginaFinal As Integer
        paginaFinal = pagActual + paginasDocumento - 1
        If paginaFinal > totalPaginas Then paginaFinal = totalPaginas
        
        ' Exportar a PDF
        Dim rutaCompleta As String
        rutaCompleta = carpeta & nombreArchivo & ".pdf"
        
        ActiveDocument.ExportAsFixedFormat _
            OutputFileName:=rutaCompleta, _
            ExportFormat:=wdExportFormatPDF, _
            OpenAfterExport:=False, _
            OptimizeFor:=wdExportOptimizeForPrint, _
            Range:=wdExportFromTo, _
            From:=pagActual, _
            To:=paginaFinal, _
            Item:=wdExportDocumentContent, _
            IncludeDocProps:=True, _
            KeepIRM:=True, _
            CreateBookmarks:=wdExportCreateNoBookmarks, _
            DocStructureTags:=True, _
            BitmapMissingFonts:=True, _
            UseISO19005_1:=False
        
        ' Avanzar a la siguiente sección
        pagActual = pagActual + paginasDocumento
        numeroDoc = numeroDoc + 1
        
        ' Mostrar progreso
        Application.StatusBar = "Procesando... " & Int((pagActual - 1) / totalPaginas * 100) & "% completado"
    Loop
    
    Application.StatusBar = False
    MsgBox "Generación terminada. Se crearon " & (numeroDoc - 1) & " archivos PDF en:" & vbCrLf & carpeta, vbInformation
    
    ' Abrir la carpeta destino en el explorador
    Shell "explorer.exe """ & carpeta & """", vbNormalFocus
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error: " & Err.Description & vbCrLf & "Número de error: " & Err.Number, vbCritical
End Sub

' Función para obtener y validar lista de nombres personalizados
Function ObtenerListaNombres(totalDocumentos As Integer, ByRef listaNombres() As String) As Boolean
    Dim respuesta As VbMsgBoxResult
    Dim textInput As String
    Dim nombres() As String
    Dim nombresLimpios() As String
    Dim i As Integer, j As Integer
    Dim contador As Integer
    Dim nombreLimpio As String
    Dim duplicados As Boolean
    Dim hayNombres As Boolean
    
    ' Preguntar si el usuario quiere usar nombres personalizados
    respuesta = MsgBox("¿Desea proporcionar nombres personalizados para los " & totalDocumentos & " archivos PDF?" & vbCrLf & vbCrLf & _
                      "Si selecciona 'Sí', podrá elegir entre:" & vbCrLf & _
                      "• Ingresar los nombres manualmente" & vbCrLf & _
                      "• Cargar los nombres desde un archivo .txt" & vbCrLf & vbCrLf & _
                      "Si selecciona 'No', se usará un nombre base con numeración automática.", _
                      vbYesNo + vbQuestion, "Nombres personalizados")
    
    If respuesta = vbNo Then
        ObtenerListaNombres = False
        Exit Function
    End If
    
    ' Preguntar el método para obtener los nombres
    Dim metodoSeleccionado As VbMsgBoxResult
    metodoSeleccionado = MsgBox("Seleccione el método para proporcionar los nombres:" & vbCrLf & vbCrLf & _
                               "• SÍ: Cargar desde archivo .txt" & vbCrLf & _
                               "• NO: Ingresar manualmente" & vbCrLf & vbCrLf & _
                               "Recomendado: Use archivo .txt para listas largas", _
                               vbYesNo + vbQuestion, "Método de entrada")
    
    If metodoSeleccionado = vbYes Then
        ' Cargar desde archivo
        Dim archivoTexto As String
        archivoTexto = CargarNombresDesdeArchivo(totalDocumentos)
        
        If archivoTexto = "" Then
            ObtenerListaNombres = False
            Exit Function
        End If
        
        textInput = archivoTexto
    Else
        ' Método manual (código existente)
        Do
            duplicados = False
            textInput = ObtenerTextoMultilinea("Ingrese los " & totalDocumentos & " nombres, uno por línea:", _
                                  "Ejemplo:" & vbCrLf & _
                                  "José Márquez" & vbCrLf & _
                                  "Martín Hernández" & vbCrLf & _
                                  "Sofía Gutiérrez" & vbCrLf & vbCrLf & _
                                  "IMPORTANTE: Debe proporcionar exactamente " & totalDocumentos & " nombres únicos.", _
                                  totalDocumentos)

            ' Si el usuario cancela
            If textInput = "" Then
                If MsgBox("¿Desea cancelar la operación?", vbYesNo + vbQuestion, "Confirmar cancelación") = vbYes Then
                    ObtenerListaNombres = False
                    Exit Function
                End If
            End If

            ' Procesar entrada manual
            Call ProcesarEntradaManual(textInput, totalDocumentos, nombres, duplicados)
        Loop Until UBound(nombres) + 1 = totalDocumentos And Not duplicados And nombres(0) <> ""
        
        ' Redimensionar el array de salida y asignar los valores
        ReDim listaNombres(UBound(nombres))
        For i = 0 To UBound(nombres)
            listaNombres(i) = nombres(i)
        Next i
        
        ObtenerListaNombres = True
        Exit Function
    End If
    
    ' Procesar los nombres (tanto desde archivo como manual)
    ' Dividir por saltos de línea y limpiar espacios
    ' Manejar diferentes tipos de salto de línea (Windows: vbCrLf, Unix: vbLf, Mac: vbCr)
    textInput = Replace(textInput, vbCrLf, vbLf)
    textInput = Replace(textInput, vbCr, vbLf)
    nombres = Split(textInput, vbLf)
    
    ' Verificar cantidad correcta
    If UBound(nombres) + 1 <> totalDocumentos Then
        MsgBox "Error: Se requieren exactamente " & totalDocumentos & " nombres." & vbCrLf & _
               "Se encontraron " & (UBound(nombres) + 1) & " nombres." & vbCrLf & vbCrLf & _
               "Por favor, verifique el archivo o inténte nuevamente.", vbExclamation
        ObtenerListaNombres = False
        Exit Function
    End If
    
    ' Filtrar nombres vacíos (líneas en blanco)
    contador = 0
    hayNombres = False
    
    For i = 0 To UBound(nombres)
        nombres(i) = Trim(nombres(i))
        If nombres(i) <> "" Then
            If Not hayNombres Then
                ReDim nombresLimpios(0)
                hayNombres = True
            Else
                ReDim Preserve nombresLimpios(contador)
            End If
            nombresLimpios(contador) = nombres(i)
            contador = contador + 1
        End If
    Next i
    
    ' Verificar que se encontraron nombres válidos
    If Not hayNombres Then
        MsgBox "Error: No se encontraron nombres válidos en la lista proporcionada.", vbExclamation
        ObtenerListaNombres = False
        Exit Function
    End If
    
    ' Actualizar el array principal con los nombres limpios
    nombres = nombresLimpios
    
    ' Verificar cantidad correcta después de limpiar
    If UBound(nombres) + 1 <> totalDocumentos Then
        MsgBox "Error: Después de eliminar nombres vacíos, se requieren exactamente " & totalDocumentos & " nombres." & vbCrLf & _
               "Se encontraron " & (UBound(nombres) + 1) & " nombres válidos." & vbCrLf & vbCrLf & _
               "Por favor, verifique el archivo o inténte nuevamente.", vbExclamation
        ObtenerListaNombres = False
        Exit Function
    End If
    
    ' Verificar duplicados
    For i = 0 To UBound(nombres)
        For j = i + 1 To UBound(nombres)
            If LCase(Trim(nombres(i))) = LCase(Trim(nombres(j))) Then
                MsgBox "Error: Se encontró un nombre duplicado: '" & nombres(i) & "'" & vbCrLf & _
                       "Posiciones: " & (i + 1) & " y " & (j + 1) & vbCrLf & vbCrLf & _
                       "Todos los nombres deben ser únicos.", vbExclamation
                ObtenerListaNombres = False
                Exit Function
            End If
        Next j
    Next i
    
    ' Redimensionar el array de salida y asignar los valores
    ReDim listaNombres(UBound(nombres))
    For i = 0 To UBound(nombres)
        listaNombres(i) = nombres(i)
    Next i
    
    ObtenerListaNombres = True
End Function

' NUEVA SUBRUTINA: ProcesarEntradaManual - Esta era la que faltaba
Sub ProcesarEntradaManual(textInput As String, totalDocumentos As Integer, ByRef nombres() As String, ByRef duplicados As Boolean)
    Dim i As Integer, j As Integer
    Dim contador As Integer
    Dim hayNombres As Boolean
    Dim nombresLimpios() As String
    
    ' Inicializar variables
    duplicados = False
    hayNombres = False
    contador = 0
    
    ' Procesar diferentes tipos de salto de línea
    textInput = Replace(textInput, vbCrLf, vbLf)
    textInput = Replace(textInput, vbCr, vbLf)
    nombres = Split(textInput, vbLf)
    
    ' Filtrar nombres vacíos y limpiar espacios
    For i = 0 To UBound(nombres)
        nombres(i) = Trim(nombres(i))
        If nombres(i) <> "" Then
            If Not hayNombres Then
                ReDim nombresLimpios(0)
                hayNombres = True
            Else
                ReDim Preserve nombresLimpios(contador)
            End If
            nombresLimpios(contador) = nombres(i)
            contador = contador + 1
        End If
    Next i
    
    ' Si se encontraron nombres válidos, actualizar el array
    If hayNombres Then
        nombres = nombresLimpios
    Else
        ' Si no hay nombres válidos, crear un array con un elemento vacío
        ReDim nombres(0)
        nombres(0) = ""
        Exit Sub
    End If
    
    ' Verificar cantidad correcta
    If UBound(nombres) + 1 <> totalDocumentos Then
        MsgBox "Error: Se requieren exactamente " & totalDocumentos & " nombres." & vbCrLf & _
               "Se ingresaron " & (UBound(nombres) + 1) & " nombres válidos." & vbCrLf & vbCrLf & _
               "Por favor, inténte nuevamente.", vbExclamation
        ReDim nombres(0)
        nombres(0) = ""
        Exit Sub
    End If
    
    ' Verificar duplicados
    For i = 0 To UBound(nombres)
        For j = i + 1 To UBound(nombres)
            If LCase(Trim(nombres(i))) = LCase(Trim(nombres(j))) Then
                MsgBox "Error: Se encontró un nombre duplicado: '" & nombres(i) & "'" & vbCrLf & _
                       "Posiciones: " & (i + 1) & " y " & (j + 1) & vbCrLf & vbCrLf & _
                       "Todos los nombres deben ser únicos. Por favor, corrija los duplicados.", vbExclamation
                duplicados = True
                Exit Sub
            End If
        Next j
    Next i
End Sub

' Función para seleccionar carpeta usando el explorador de archivos
Function SeleccionarCarpeta() As String
    Dim folderDialog As fileDialog
    Dim carpetaSeleccionada As String
    
    ' Crear el diálogo de selección de carpeta
    Set folderDialog = Application.fileDialog(msoFileDialogFolderPicker)
    
    With folderDialog
        .Title = "Seleccione la carpeta destino para los archivos PDF"
        .AllowMultiSelect = False
        ' Establecer carpeta inicial (opcional)
        .InitialFileName = Environ("USERPROFILE") & "\Desktop\"
    End With
    
    ' Mostrar el diálogo y capturar selección
    If folderDialog.Show = -1 Then
        carpetaSeleccionada = folderDialog.SelectedItems(1)
    Else
        carpetaSeleccionada = "" ' Usuario canceló
    End If
    
    Set folderDialog = Nothing
    SeleccionarCarpeta = carpetaSeleccionada
End Function

' Función auxiliar para limpiar nombres de archivo
Function LimpiarNombreArchivo(nombre As String) As String
    Dim caracteresInvalidos As String
    Dim i As Integer
    
    caracteresInvalidos = "\/:*?""<>|"
    LimpiarNombreArchivo = nombre
    
    For i = 1 To Len(caracteresInvalidos)
        LimpiarNombreArchivo = Replace(LimpiarNombreArchivo, Mid(caracteresInvalidos, i, 1), "_")
    Next i
    
    ' Eliminar espacios extras y recortar
    LimpiarNombreArchivo = Trim(LimpiarNombreArchivo)
    
    ' Asegurar que no esté vacío
    If LimpiarNombreArchivo = "" Then LimpiarNombreArchivo = "documento"
End Function

' Función para obtener texto multilínea usando un formulario personalizado
Function ObtenerTextoMultilinea(prompt As String, ejemplo As String, totalDocs As Integer) As String
    Dim textoIngresado As String
    Dim resultado As String
    
    ' Crear un formulario temporal usando Application.InputBox con un enfoque alternativo
    ' Como no podemos crear formularios complejos fácilmente, usaremos una solución más simple
    
    resultado = InputBox(prompt & vbCrLf & vbCrLf & ejemplo, "Lista de nombres - Separe cada nombre con ENTER", "")
    If resultado = "" Then
        ObtenerTextoMultilinea = ""
        Exit Function
    End If
    
    ' Si el usuario no puede ingresar todo el texto, ofrecemos una alternativa
    If resultado <> "" And InStr(resultado, vbCrLf) = 0 And InStr(resultado, vbLf) = 0 Then
        ' Si no hay saltos de línea, probablemente el usuario no pudo ingresar todo
        Dim respuesta As VbMsgBoxResult
        respuesta = MsgBox("Se detectó que ingresó un solo nombre o no pudo ingresar todos los nombres." & vbCrLf & vbCrLf & _
                          "¿Desea intentar un método alternativo?" & vbCrLf & vbCrLf & _
                          "Método alternativo: Ingresar nombres uno por uno.", vbYesNo + vbQuestion, "Método alternativo")
        
        If respuesta = vbYes Then
            resultado = ObtenerNombresUnoAUno(totalDocs)
        End If
    End If
    
    ObtenerTextoMultilinea = resultado
End Function

' Función alternativa para ingresar nombres uno por uno
Function ObtenerNombresUnoAUno(totalDocumentos As Integer) As String
    Dim nombres As String
    Dim nombreActual As String
    Dim i As Integer
    
    nombres = ""
    
    For i = 1 To totalDocumentos
        nombreActual = InputBox("Ingrese el nombre #" & i & " de " & totalDocumentos & ":", "Nombre " & i, "")
        If nombreActual = "" Then
            ' Confirmar cancelación
            If MsgBox("¿Desea cancelar la operación?", vbYesNo + vbQuestion, "Confirmar cancelación") = vbYes Then
                ObtenerNombresUnoAUno = ""
                Exit Function
            Else
                ' Permitir continuar con el mismo índice
                i = i - 1
                GoTo ContinuarLoop
            End If
        End If
        
        nombres = nombres & nombreActual
        If i < totalDocumentos Then nombres = nombres & vbCrLf

ContinuarLoop:
        Next i
    
    ObtenerNombresUnoAUno = nombres
End Function

' Función para cargar nombres desde un archivo de texto con soporte UTF-8
Function CargarNombresDesdeArchivo(totalDocumentos As Integer) As String
    Dim fileDialog As fileDialog
    Dim archivoSeleccionado As String
    Dim contenidoArchivo As String
    Dim linea As String
    Dim stream As Object
    
    ' Crear el diálogo de selección de archivo
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    
    With fileDialog
        .Title = "Seleccione el archivo de texto con los nombres"
        .AllowMultiSelect = False
        .InitialFileName = Environ("USERPROFILE") & "\Desktop\"
        
        ' Filtros de archivo
        .Filters.Clear
        .Filters.Add "Archivos de texto", "*.txt"
        .Filters.Add "Todos los archivos", "*.*"
    End With
    
    ' Mostrar el diálogo y verificar selección
    If fileDialog.Show = -1 Then
        archivoSeleccionado = fileDialog.SelectedItems(1)
    Else
        ' Usuario canceló
        MsgBox "No se seleccionó ningún archivo. Operación cancelada.", vbInformation
        CargarNombresDesdeArchivo = ""
        Exit Function
    End If
    
    ' Verificar que el archivo existe
    If Dir(archivoSeleccionado) = "" Then
        MsgBox "Error: El archivo seleccionado no existe o no se puede acceder a él.", vbExclamation
        CargarNombresDesdeArchivo = ""
        Exit Function
    End If
    
    ' Leer el contenido del archivo usando ADODB.Stream para manejar UTF-8
    On Error GoTo ErrorLectura
    
    ' Crear objeto Stream para leer UTF-8
    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .Type = 2 ' Especifica que es texto
        .Charset = "UTF-8" ' Especifica la codificación UTF-8
        .Open
        .LoadFromFile archivoSeleccionado
        contenidoArchivo = .ReadText
        .Close
    End With
    
    Set stream = Nothing
    
    ' Remover BOM (Byte Order Mark) si existe al inicio del archivo UTF-8
    If Len(contenidoArchivo) > 0 Then
        If Asc(Left(contenidoArchivo, 1)) = 65279 Then ' BOM UTF-8
            contenidoArchivo = Mid(contenidoArchivo, 2)
        End If
    End If
    
    ' Normalizar saltos de línea
    contenidoArchivo = Replace(contenidoArchivo, vbCrLf, vbLf)
    contenidoArchivo = Replace(contenidoArchivo, vbCr, vbLf)
    
    ' Remover el último salto de línea si existe
    If Right(contenidoArchivo, 1) = vbLf Then
        contenidoArchivo = Left(contenidoArchivo, Len(contenidoArchivo) - 1)
    End If
    
    ' Verificar que el archivo no esté vacío
    If Trim(contenidoArchivo) = "" Then
        MsgBox "Error: El archivo seleccionado está vacío.", vbExclamation
        CargarNombresDesdeArchivo = ""
        Exit Function
    End If
    
    ' Mostrar información del archivo cargado
    Dim lineasEncontradas As Integer
    Dim nombresArray() As String
    nombresArray = Split(contenidoArchivo, vbLf)
    lineasEncontradas = UBound(nombresArray) + 1
    
    ' Crear vista previa (primeras 3 líneas)
    Dim vistaPrevia As String
    Dim i As Integer
    vistaPrevia = ""
    For i = 0 To UBound(nombresArray)
        If i >= 3 Then Exit For
        vistaPrevia = vistaPrevia & nombresArray(i)
        If i < UBound(nombresArray) And i < 2 Then vistaPrevia = vistaPrevia & vbCrLf
    Next i
    
    MsgBox "Archivo cargado exitosamente:" & vbCrLf & vbCrLf & _
           "Archivo: " & Right(archivoSeleccionado, Len(archivoSeleccionado) - InStrRev(archivoSeleccionado, "\")) & vbCrLf & _
           "Líneas encontradas: " & lineasEncontradas & vbCrLf & _
           "Líneas requeridas: " & totalDocumentos & vbCrLf & vbCrLf & _
           "Vista previa (primeras 3 líneas):" & vbCrLf & _
           vistaPrevia, _
           vbInformation, "Archivo cargado"
    
    CargarNombresDesdeArchivo = contenidoArchivo
    Exit Function
    
ErrorLectura:
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
    MsgBox "Error al leer el archivo: " & Err.Description & vbCrLf & _
           "Verifique que el archivo no esté siendo usado por otro programa.", vbExclamation
    CargarNombresDesdeArchivo = ""
End Function