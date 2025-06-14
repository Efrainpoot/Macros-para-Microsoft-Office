Sub AplicarAnimacionBarrido()
    Dim slide As slide
    Dim shape As shape
    Dim animEffect As Effect
    Dim existingEffect As Effect
    Dim hasAnimation As Boolean
    Dim shapesArray() As shape
    Dim i As Integer, j As Integer
    Dim shapeCount As Integer
    Dim tempShape As shape
    
    ' Verificar que hay una diapositiva activa
    If ActiveWindow.ViewType <> ppViewSlide And ActiveWindow.ViewType <> ppViewNormal Then
        MsgBox "Por favor, selecciona una diapositiva en vista Normal o de Diapositiva"
        Exit Sub
    End If
    
    ' Obtener la diapositiva activa
    Set slide = ActiveWindow.View.slide
    
    ' Contar formas válidas y crear array
    shapeCount = 0
    For Each shape In slide.shapes
        ' Incluir todas las formas excepto algunos placeholders específicos
        If EsFormaValida(shape) Then
            shapeCount = shapeCount + 1
        End If
    Next shape
    
    ' Redimensionar array
    ReDim shapesArray(1 To shapeCount)
    
    ' Llenar array con las formas
    i = 0
    For Each shape In slide.shapes
        If EsFormaValida(shape) Then
            i = i + 1
            Set shapesArray(i) = shape
        End If
    Next shape
    
    ' Ordenar formas por posición vertical (de arriba hacia abajo)
    ' Usando algoritmo de burbuja simple
    For i = 1 To shapeCount - 1
        For j = i + 1 To shapeCount
            ' Si la forma actual está más abajo que la siguiente, intercambiar
            If shapesArray(i).Top > shapesArray(j).Top Then
                Set tempShape = shapesArray(i)
                Set shapesArray(i) = shapesArray(j)
                Set shapesArray(j) = tempShape
            End If
        Next j
    Next i
    
    ' Aplicar animaciones en el orden correcto
    For i = 1 To shapeCount
        Set shape = shapesArray(i)
        On Error Resume Next ' Ignorar errores individuales
        
        ' Verificar si la forma ya tiene animaciones
        hasAnimation = False
        For Each existingEffect In slide.TimeLine.MainSequence
            If existingEffect.shape Is shape Then
                hasAnimation = True
                Exit For
            End If
        Next existingEffect
        
        ' Solo aplicar animación si no tiene animaciones previas
        If Not hasAnimation Then
            ' Determinar el nivel de animación apropiado
            If shape.Type = msoGroup Then
                ' Para grupos: animar como un solo objeto - FORZAR ANIMACIÓN DE GRUPO
                Set animEffect = slide.TimeLine.MainSequence.AddEffect( _
                    shape:=shape, _
                    effectId:=msoAnimEffectWipe, _
                    Level:=msoAnimateShapePicture, _
                    trigger:=msoAnimTriggerAfterPrevious)
                
                ' Configuración adicional para grupos
                If Not animEffect Is Nothing Then
                    On Error Resume Next
                    animEffect.EffectParameters.Direction = msoAnimDirectionLeft
                    animEffect.EffectParameters.Direction = 4
                    animEffect.Exit = msoFalse
                    animEffect.Timing.TriggerType = msoAnimTriggerAfterPrevious
                    On Error GoTo 0
                End If
            ElseIf shape.HasTextFrame And shape.TextFrame.HasText Then
                ' Para formas con texto: usar animación por párrafo
                Set animEffect = slide.TimeLine.MainSequence.AddEffect( _
                    shape:=shape, _
                    effectId:=msoAnimEffectWipe, _
                    Level:=msoAnimateTextByFirstLevel, _
                    trigger:=msoAnimTriggerAfterPrevious)
            Else
                ' Para formas sin texto: animación normal
                Set animEffect = slide.TimeLine.MainSequence.AddEffect( _
                    shape:=shape, _
                    effectId:=msoAnimEffectWipe, _
                    Level:=msoAnimateTextByAllLevels, _
                    trigger:=msoAnimTriggerAfterPrevious)
            End If
            
            ' Configurar dirección desde la izquierda
            If Not animEffect Is Nothing Then
                On Error Resume Next
                ' Múltiples intentos para asegurar la dirección
                animEffect.EffectParameters.Direction = msoAnimDirectionLeft
                animEffect.EffectParameters.Direction = 4 ' Valor numérico para izquierda
                animEffect.Exit = msoFalse
                animEffect.Timing.TriggerType = msoAnimTriggerAfterPrevious
                On Error GoTo 0
            End If
        End If
        
        On Error GoTo 0 ' Restaurar manejo normal de errores
    Next i
    
    ' Asegurar que todas las animaciones tengan "Después de la anterior"
    For Each existingEffect In slide.TimeLine.MainSequence
        On Error Resume Next
        existingEffect.Timing.TriggerType = msoAnimTriggerAfterPrevious
        On Error GoTo 0
    Next existingEffect
    
    ' Aplicar dirección "Desde la izquierda" a todas las animaciones
    For Each existingEffect In slide.TimeLine.MainSequence
        On Error Resume Next
        existingEffect.EffectParameters.Direction = msoAnimDirectionLeft
        existingEffect.EffectParameters.Direction = 4
        On Error GoTo 0
    Next existingEffect
    

End Sub

' Función auxiliar para determinar si una forma es válida para animación
Function EsFormaValida(shape As shape) As Boolean
    On Error Resume Next
    
    ' Por defecto, incluir la forma
    EsFormaValida = True
    
    ' INCLUIR EXPLÍCITAMENTE LOS GRUPOS
    If shape.Type = msoGroup Then
        EsFormaValida = True
        On Error GoTo 0
        Exit Function
    End If
    
    ' Excluir solo placeholders específicos que no queremos animar
    If shape.Type = msoPlaceholder Then
        ' Solo verificar PlaceholderFormat si es realmente un placeholder
        Select Case shape.PlaceholderFormat.Type
            Case ppPlaceholderTitle, ppPlaceholderSubtitle, ppPlaceholderSlideNumber, _
                 ppPlaceholderDate, ppPlaceholderFooter, ppPlaceholderHeader
                EsFormaValida = False
            Case Else
                EsFormaValida = True ' Incluir otros tipos de placeholders
        End Select
    End If
    
    ' Excluir formas muy pequeñas (probablemente decorativas), pero NO grupos
    If shape.Width < 10 Or shape.Height < 10 Then
        EsFormaValida = False
    End If
    
    On Error GoTo 0
End Function

