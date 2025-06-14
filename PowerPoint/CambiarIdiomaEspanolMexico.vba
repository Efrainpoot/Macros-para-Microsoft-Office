Option Explicit

' Constante para Español México
Const SPANISH_MEXICO As Long = &H80A&  ' 2058 decimal

' Macro principal para cambiar idioma a Español México
Public Sub CambiarIdiomaEspanolMexico()
    Dim ppt As Presentation
    Dim totalElementos As Long
    Dim mensaje As String
    
    ' Verificar que PowerPoint esté activo y tenga una presentación
    If Application.Presentations.count = 0 Then
        MsgBox "No hay presentaciones abiertas en PowerPoint.", vbExclamation, "Sin presentaciones"
        Exit Sub
    End If
    
    ' Usar la presentación activa
    Set ppt = Application.ActivePresentation
    
    ' Cambiar idioma en toda la presentación
    totalElementos = CambiarIdiomaCompleto(ppt)
    
    ' Mostrar resultado
    mensaje = "Proceso completado exitosamente!" & vbCrLf & vbCrLf
    mensaje = mensaje & "Presentación: " & ppt.Name & vbCrLf
    mensaje = mensaje & "Elementos procesados: " & totalElementos & vbCrLf
    mensaje = mensaje & "Idioma establecido: Español (México)"
    
    MsgBox mensaje, vbInformation, "Cambio de idioma completado"
End Sub

' Función principal que cambia el idioma en toda la presentación
Private Function CambiarIdiomaCompleto(pres As Presentation) As Long
    Dim totalElementos As Long
    Dim slide As slide
    Dim shape As shape
    
    totalElementos = 0
    
    ' 1. Establecer idioma por defecto de la presentación
    On Error Resume Next
    pres.DefaultLanguageID = SPANISH_MEXICO
    On Error GoTo 0
    
    ' 2. Cambiar idioma en el Slide Master
    totalElementos = totalElementos + CambiarIdiomaEnMaster(pres.SlideMaster)
    
    ' 3. Cambiar idioma en todas las diapositivas
    For Each slide In pres.Slides
        For Each shape In slide.shapes
            If CambiarIdiomaEnShape(shape) Then
                totalElementos = totalElementos + 1
            End If
        Next shape
    Next slide
    
    ' 4. Cambiar idioma en Notes Master si existe
    On Error Resume Next
    If Not pres.NotesMaster Is Nothing Then
        totalElementos = totalElementos + CambiarIdiomaEnMaster(pres.NotesMaster)
    End If
    On Error GoTo 0
    
    ' 5. Cambiar idioma en Handout Master si existe
    On Error Resume Next
    If Not pres.HandoutMaster Is Nothing Then
        totalElementos = totalElementos + CambiarIdiomaEnMaster(pres.HandoutMaster)
    End If
    On Error GoTo 0
    
    CambiarIdiomaCompleto = totalElementos
End Function

' Cambiar idioma en un Master (plantilla)
Private Function CambiarIdiomaEnMaster(masterObj As Object) As Long
    Dim shape As shape
    Dim contador As Long
    Dim layout As CustomLayout
    
    contador = 0
    
    On Error Resume Next
    
    ' Cambiar en el master principal
    For Each shape In masterObj.shapes
        If CambiarIdiomaEnShape(shape) Then
            contador = contador + 1
        End If
    Next shape
    
    ' Cambiar en todos los layouts si es SlideMaster
    If TypeName(masterObj) = "Master" Then
        For Each layout In masterObj.CustomLayouts
            For Each shape In layout.shapes
                If CambiarIdiomaEnShape(shape) Then
                    contador = contador + 1
                End If
            Next shape
        Next layout
    End If
    
    On Error GoTo 0
    CambiarIdiomaEnMaster = contador
End Function

' Cambiar idioma en una forma específica
Private Function CambiarIdiomaEnShape(shp As shape) As Boolean
    Dim textRange As textRange
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' Verificar si la forma tiene texto
    If Not shp.HasTextFrame Then
        CambiarIdiomaEnShape = False
        Exit Function
    End If
    
    ' Obtener el rango de texto
    Set textRange = shp.TextFrame.textRange
    
    ' Cambiar idioma del texto completo
    textRange.LanguageID = SPANISH_MEXICO
    
    ' Cambiar idioma párrafo por párrafo para asegurar
    For i = 1 To textRange.Paragraphs.count
        textRange.Paragraphs(i).LanguageID = SPANISH_MEXICO
    Next i
    
    ' Cambiar idioma palabra por palabra si es necesario
    For i = 1 To textRange.Words.count
        textRange.Words(i).LanguageID = SPANISH_MEXICO
    Next i
    
    CambiarIdiomaEnShape = True
    Exit Function
    
ErrorHandler:
    CambiarIdiomaEnShape = False
End Function

' Macro para verificar idiomas actuales con opción de corrección automática
Public Sub VerificarIdiomasActuales()
    Dim ppt As Presentation
    Dim slide As slide
    Dim shape As shape
    Dim reporte As String
    Dim idiomasEncontrados As String
    Dim necesitaCorreccion As Boolean
    Dim respuesta As VbMsgBoxResult
    Dim totalElementos As Long
    
    If Application.Presentations.count = 0 Then
        MsgBox "No hay presentaciones abiertas.", vbExclamation, "Sin presentaciones"
        Exit Sub
    End If
    
    Set ppt = Application.ActivePresentation
    necesitaCorreccion = False
    
    ' Construir reporte detallado
    reporte = "REPORTE DE IDIOMAS - " & ppt.Name & vbCrLf
    reporte = reporte & String(38, "=") & vbCrLf & vbCrLf
    
    ' Verificar idioma por defecto
    If ppt.DefaultLanguageID <> SPANISH_MEXICO Then
        reporte = reporte & "[X] Idioma por defecto: " & ppt.DefaultLanguageID & " (" & NombreIdioma(ppt.DefaultLanguageID) & ")" & vbCrLf
        necesitaCorreccion = True
    Else
        reporte = reporte & "[OK] Idioma por defecto: " & ppt.DefaultLanguageID & " (" & NombreIdioma(ppt.DefaultLanguageID) & ")" & vbCrLf
    End If
    
    reporte = reporte & vbCrLf & "IDIOMAS EN DIAPOSITIVAS:" & vbCrLf
    idiomasEncontrados = ""
    
    ' Revisar todas las diapositivas
    For Each slide In ppt.Slides
        For Each shape In slide.shapes
            On Error Resume Next
            If shape.HasTextFrame Then
                If shape.TextFrame.HasText Then
                    Dim langID As Long
                    langID = shape.TextFrame.textRange.LanguageID
                    
                    ' Solo reportar la primera vez que encuentra cada idioma
                    If InStr(idiomasEncontrados, CStr(langID) & ",") = 0 Then
                        idiomasEncontrados = idiomasEncontrados & langID & ","
                        
                        If langID <> SPANISH_MEXICO Then
                            reporte = reporte & "[X] Diapositiva " & slide.SlideIndex & ": " & langID & " (" & NombreIdioma(langID) & ")" & vbCrLf
                            necesitaCorreccion = True
                        Else
                            reporte = reporte & "[OK] Diapositiva " & slide.SlideIndex & ": " & langID & " (" & NombreIdioma(langID) & ")" & vbCrLf
                        End If
                    End If
                End If
            End If
            On Error GoTo 0
        Next shape
    Next slide
    
    ' Mostrar resultado y opciones según el estado
    If necesitaCorreccion Then
        reporte = reporte & vbCrLf & "*** SE ENCONTRARON IDIOMAS INCORRECTOS ***" & vbCrLf
        reporte = reporte & "Deseas cambiar todo a Espanol (Mexico) automaticamente?"
        
        respuesta = MsgBox(reporte, vbYesNo + vbQuestion, "Corrección de idiomas requerida")
        
        If respuesta = vbYes Then
            ' Ejecutar corrección automática
            totalElementos = CambiarIdiomaCompleto(ppt)
            
            ' Mostrar resultado de la corrección
            Dim mensajeExito As String
            mensajeExito = "*** CORRECCION COMPLETADA EXITOSAMENTE ***" & vbCrLf & vbCrLf
            mensajeExito = mensajeExito & "RESUMEN:" & vbCrLf
            mensajeExito = mensajeExito & "- Presentacion: " & ppt.Name & vbCrLf
            mensajeExito = mensajeExito & "- Elementos procesados: " & totalElementos & vbCrLf
            mensajeExito = mensajeExito & "- Idioma establecido: Espanol (Mexico)" & vbCrLf & vbCrLf
            mensajeExito = mensajeExito & "Todos los textos han sido configurados correctamente."
            
            MsgBox mensajeExito, vbInformation, "Cambio completado"
        End If
    Else
        reporte = reporte & vbCrLf & "*** PERFECTO - TODO ESTA CORRECTO ***" & vbCrLf
        reporte = reporte & "Todos los elementos ya estan en Espanol (Mexico)."
        
        MsgBox reporte, vbInformation, "Idiomas correctos"
    End If
End Sub

' Función auxiliar para obtener nombre del idioma
Private Function NombreIdioma(langID As Long) As String
    Select Case langID
        Case 1033: NombreIdioma = "Inglés (Estados Unidos)"
        Case 2058: NombreIdioma = "Español (México)"
        Case &H80A&: NombreIdioma = "Español (México)"
        Case 1034: NombreIdioma = "Español (España)"
        Case 3082: NombreIdioma = "Español (España - Internacional)"
        Case Else: NombreIdioma = "Idioma desconocido"
    End Select
End Function

' Macro para cambiar idioma en presentación específica (por nombre)
Public Sub CambiarIdiomaPresentacionEspecifica(nombrePresentacion As String)
    Dim ppt As Presentation
    Dim totalElementos As Long
    
    On Error GoTo ErrorHandler
    
    Set ppt = Application.Presentations(nombrePresentacion)
    totalElementos = CambiarIdiomaCompleto(ppt)
    
    MsgBox "Idioma cambiado en " & nombrePresentacion & vbCrLf & "Elementos procesados: " & totalElementos, vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "No se pudo encontrar la presentación: " & nombrePresentacion, vbExclamation
End Sub
