Sub RenombrarObjetosParaTransformacion()
    '====================================================================
    ' Macro: RenombrarObjetosParaTransformacion
    ' Descripción: Renombra todos los objetos con prefijo "!! " para
    '              facilitar el uso de la transición "Transformación"
    '              Creado para uso universal en PowerPoint
    ' Nota: La transición "Transformación" debe aplicarse manualmente
    '====================================================================
    
    Dim pres As Presentation
    Dim sld As slide
    Dim shp As shape
    Dim i As Integer
    Dim j As Integer
    Dim originalName As String
    Dim newName As String
    Dim totalSlides As Integer
    Dim totalObjects As Integer
    Dim processedObjects As Integer
    Dim skippedObjects As Integer
    
    ' Obtener la presentación activa
    Set pres = Application.ActivePresentation
    
    ' Inicializar contadores
    totalSlides = pres.Slides.count
    totalObjects = 0
    processedObjects = 0
    skippedObjects = 0
    
    ' Contar todos los objetos
    For i = 1 To totalSlides
        Set sld = pres.Slides(i)
        totalObjects = totalObjects + sld.shapes.count
    Next i
    
    ' Mostrar mensaje de inicio
    If MsgBox("Esta macro renombrará " & totalObjects & " objetos en " & totalSlides & " diapositivas." & vbCrLf & vbCrLf & _
              "Todos los objetos tendrán el prefijo '!! ' para optimizar" & vbCrLf & _
              "el uso de la transición 'Transformación'." & vbCrLf & vbCrLf & _
              "¿Deseas continuar?", vbYesNo + vbQuestion, "Confirmar Renombrado") = vbNo Then
        Exit Sub
    End If
    
    ' Renombrar todos los objetos en todas las diapositivas
    For i = 1 To totalSlides
        Set sld = pres.Slides(i)
        
        ' Recorrer todas las formas en la diapositiva
        For j = 1 To sld.shapes.count
            Set shp = sld.shapes(j)
            originalName = shp.Name
            
            ' Verificar si el nombre ya tiene el prefijo "!! "
            If Left(originalName, 3) <> "!! " Then
                ' Crear el nuevo nombre con prefijo
                newName = "!! " & originalName
                
                ' Verificar que el nuevo nombre no exceda el límite de caracteres
                If Len(newName) <= 255 Then
                    ' Intentar renombrar el objeto
                    On Error Resume Next
                    shp.Name = newName
                    If Err.Number = 0 Then
                        processedObjects = processedObjects + 1
                    Else
                        skippedObjects = skippedObjects + 1
                    End If
                    On Error GoTo 0
                Else
                    ' Si el nombre es muy largo, truncar el original
                    newName = "!! " & Left(originalName, 252)
                    On Error Resume Next
                    shp.Name = newName
                    If Err.Number = 0 Then
                        processedObjects = processedObjects + 1
                    Else
                        skippedObjects = skippedObjects + 1
                    End If
                    On Error GoTo 0
                End If
            Else
                ' El objeto ya tiene el prefijo, contarlo como ya procesado
                skippedObjects = skippedObjects + 1
            End If
        Next j
    Next i
    
    ' Mostrar mensaje de finalización
    Dim mensaje As String
    mensaje = "? Renombrado completado exitosamente!" & vbCrLf & vbCrLf
    mensaje = mensaje & "RESUMEN:" & vbCrLf
    mensaje = mensaje & "• Diapositivas procesadas: " & totalSlides & vbCrLf
    mensaje = mensaje & "• Objetos renombrados: " & processedObjects & vbCrLf
    mensaje = mensaje & "• Objetos omitidos: " & skippedObjects & vbCrLf
    mensaje = mensaje & "• Total de objetos: " & totalObjects & vbCrLf & vbCrLf
    mensaje = mensaje & "SIGUIENTE PASO:" & vbCrLf
    mensaje = mensaje & "Ahora puedes aplicar manualmente la transición" & vbCrLf
    mensaje = mensaje & "'Transformación' desde la pestaña Transiciones."
    
    MsgBox mensaje, vbInformation, "Renombrado Completado"
End Sub

'====================================================================
' FUNCIONES AUXILIARES
'====================================================================

Sub VerificarObjetosRenombrados()
    '====================================================================
    ' Función auxiliar para verificar objetos con prefijo "!! "
    '====================================================================
    
    Dim pres As Presentation
    Dim sld As slide
    Dim shp As shape
    Dim i As Integer, j As Integer
    Dim objetosConPrefijo As Integer
    Dim objetosSinPrefijo As Integer
    Dim totalObjetos As Integer
    Dim ejemplos As String
    Dim contadorEjemplos As Integer
    
    Set pres = Application.ActivePresentation
    objetosConPrefijo = 0
    objetosSinPrefijo = 0
    totalObjetos = 0
    ejemplos = ""
    contadorEjemplos = 0
    
    For i = 1 To pres.Slides.count
        Set sld = pres.Slides(i)
        For j = 1 To sld.shapes.count
            Set shp = sld.shapes(j)
            totalObjetos = totalObjetos + 1
            
            If Left(shp.Name, 3) = "!! " Then
                objetosConPrefijo = objetosConPrefijo + 1
                ' Agregar ejemplos (máximo 5)
                If contadorEjemplos < 5 Then
                    ejemplos = ejemplos & "• " & shp.Name & " (Diap. " & i & ")" & vbCrLf
                    contadorEjemplos = contadorEjemplos + 1
                End If
            Else
                objetosSinPrefijo = objetosSinPrefijo + 1
            End If
        Next j
    Next i
    
    Dim resultado As String
    resultado = "VERIFICACIÓN DE RENOMBRADO:" & vbCrLf & vbCrLf
    resultado = resultado & "Objetos con prefijo '!! ': " & objetosConPrefijo & vbCrLf
    resultado = resultado & "Objetos sin prefijo: " & objetosSinPrefijo & vbCrLf
    resultado = resultado & "Total de objetos: " & totalObjetos & vbCrLf & vbCrLf
    
    If objetosConPrefijo > 0 Then
        resultado = resultado & "EJEMPLOS DE OBJETOS RENOMBRADOS:" & vbCrLf
        resultado = resultado & ejemplos
        If objetosConPrefijo > 5 Then
            resultado = resultado & "... y " & (objetosConPrefijo - 5) & " más." & vbCrLf
        End If
    End If
    
    If objetosSinPrefijo > 0 Then
        resultado = resultado & vbCrLf & "Hay " & objetosSinPrefijo & " objetos sin renombrar."
        resultado = resultado & vbCrLf & "Ejecuta 'RenombrarObjetosParaTransformacion' para completar el proceso."
    Else
        resultado = resultado & vbCrLf & "¡Todos los objetos están correctamente renombrados!"
    End If
    
    MsgBox resultado, vbInformation, "Verificación de Renombrado"
End Sub

Sub RemoverPrefijosObjetos()
    '====================================================================
    ' Función para remover el prefijo "!! " de todos los objetos
    ' Útil si necesitas revertir el proceso
    '====================================================================
    
    Dim pres As Presentation
    Dim sld As slide
    Dim shp As shape
    Dim i As Integer, j As Integer
    Dim originalName As String
    Dim newName As String
    Dim objetosModificados As Integer
    
    Set pres = Application.ActivePresentation
    objetosModificados = 0
    
    If MsgBox("Esta función REMOVERÁ el prefijo '!! ' de todos los objetos." & vbCrLf & vbCrLf & _
              "¿Estás seguro de que deseas continuar?", _
              vbYesNo + vbExclamation, "Confirmar Remoción de Prefijos") = vbNo Then
        Exit Sub
    End If
    
    For i = 1 To pres.Slides.count
        Set sld = pres.Slides(i)
        For j = 1 To sld.shapes.count
            Set shp = sld.shapes(j)
            originalName = shp.Name
            
            ' Si tiene el prefijo, removerlo
            If Left(originalName, 3) = "!! " Then
                newName = Mid(originalName, 4) ' Remover los primeros 3 caracteres
                On Error Resume Next
                shp.Name = newName
                If Err.Number = 0 Then
                    objetosModificados = objetosModificados + 1
                End If
                On Error GoTo 0
            End If
        Next j
    Next i
    
    MsgBox "Prefijos removidos exitosamente." & vbCrLf & vbCrLf & _
           "Objetos modificados: " & objetosModificados, _
           vbInformation, "Remoción Completada"
End Sub

Sub InstruccionesTransformacion()
    '====================================================================
    ' Muestra instrucciones para aplicar manualmente la transición Transformación
    '====================================================================
    
    Dim instrucciones As String
    instrucciones = "CÓMO APLICAR LA TRANSICIÓN 'TRANSFORMACIÓN':" & vbCrLf & vbCrLf
    instrucciones = instrucciones & "1. Selecciona TODAS las diapositivas:" & vbCrLf
    instrucciones = instrucciones & "   • Ctrl + A en el panel de diapositivas" & vbCrLf & vbCrLf
    instrucciones = instrucciones & "2. Ve a la pestaña 'TRANSICIONES'" & vbCrLf & vbCrLf
    instrucciones = instrucciones & "3. Busca y selecciona 'TRANSFORMACIÓN'" & vbCrLf
    instrucciones = instrucciones & "   • También puede aparecer como 'MORPH'" & vbCrLf & vbCrLf
    instrucciones = instrucciones & "4. Ajusta la duración si es necesario" & vbCrLf & vbCrLf
    instrucciones = instrucciones & "¡Listo! La transición se aplicará a todas las diapositivas" & vbCrLf & vbCrLf
    instrucciones = instrucciones & "NOTA: Los objetos ya están renombrados con '!! '" & vbCrLf
    instrucciones = instrucciones & "    para optimizar el funcionamiento de la transición."
    
    MsgBox instrucciones, vbInformation, "Instrucciones - Transición Transformación"
End Sub

