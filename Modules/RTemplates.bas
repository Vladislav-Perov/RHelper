' Первый вариант
Sub CreateTaskTemplate()
    Dim taskNum As String
    Dim imgCountInput As String
    Dim imgCount As Integer
    Dim i As Integer
    Dim currentParagraph As Paragraph
    
    ' Запрос номера задания
    taskNum = InputBox("Введите номер задания:", "Номер задания", "1")
    If taskNum = "" Then Exit Sub ' Если пользователь нажал Отмена
    
    ' Запрос количества рисунков
    imgCountInput = InputBox("Введите количество рисунков для задания:", "Количество рисунков", "1")
    If imgCountInput = "" Then Exit Sub
    
    ' Проверка на корректность ввода числа
    If Not IsNumeric(imgCountInput) Then
        MsgBox "Количество рисунков должно быть числом!", vbCritical, "Ошибка"
        Exit Sub
    End If
    imgCount = CInt(imgCountInput)
    
    ' --- Строка 1: Задание N. ---
    Set currentParagraph = Selection.Paragraphs.Add
    currentParagraph.Range.Text = "Задание " & taskNum & "." & vbCrLf
    With currentParagraph
        .Alignment = wdAlignParagraphJustify
        .LeftIndent = CentimetersToPoints(1.25)
        .FirstLineIndent = CentimetersToPoints(0)
    End With
    Selection.MoveDown Unit:=wdLine, Count:=1
    
    ' --- Строка 2: Результат выполнения задания N представлен в листинге N. ---
    Set currentParagraph = Selection.Paragraphs.Add
    currentParagraph.Range.Text = "Результат выполнения задания " & taskNum & " представлен в листинге " & taskNum & "." & vbCrLf
    With currentParagraph
        .Alignment = wdAlignParagraphJustify
        .LeftIndent = CentimetersToPoints(1.25)
        .FirstLineIndent = CentimetersToPoints(0)
    End With
    Selection.MoveDown Unit:=wdLine, Count:=1
    
    ' --- Строка 3: Листинг N – Код задания N ---
    Set currentParagraph = Selection.Paragraphs.Add
    currentParagraph.Range.Text = "Листинг " & taskNum & " – Код задания " & taskNum & vbCrLf
    With currentParagraph
        .Alignment = wdAlignParagraphJustify
        .LeftIndent = CentimetersToPoints(0)
        .FirstLineIndent = CentimetersToPoints(0)
    End With
    Selection.MoveDown Unit:=wdLine, Count:=1
    
    ' --- Строка 4: Результат выполнения задания N представлен на рисунке/рисунках ---
    Set currentParagraph = Selection.Paragraphs.Add
    If imgCount <= 1 Then
        currentParagraph.Range.Text = "Результат выполнения задания " & taskNum & " представлен на рисунке 1." & vbCrLf
    Else
        currentParagraph.Range.Text = "Результат выполнения задания " & taskNum & " представлен на рисунках 1 - " & imgCount & "." & vbCrLf
    End If
    With currentParagraph
        .Alignment = wdAlignParagraphJustify
        .LeftIndent = CentimetersToPoints(1.25)
        .FirstLineIndent = CentimetersToPoints(0)
    End With
    Selection.MoveDown Unit:=wdLine, Count:=1
    
    ' --- Строка 5 и далее: Оформление подписей к рисункам ---
    For i = 1 To imgCount
        Set currentParagraph = Selection.Paragraphs.Add
        If i = imgCount Then
            currentParagraph.Range.Text = "Рисунок " & i & " – Результат выполнения задания " & taskNum
        Else
            currentParagraph.Range.Text = "Рисунок " & i & " – Результат выполнения задания " & taskNum & vbCrLf
        End If
        With currentParagraph
            .Alignment = wdAlignParagraphCenter
            .LeftIndent = CentimetersToPoints(0)
            .FirstLineIndent = CentimetersToPoints(0)
        End With
        Selection.MoveDown Unit:=wdLine, Count:=1
    Next i
End Sub

