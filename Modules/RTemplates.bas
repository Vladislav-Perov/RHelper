Option Explicit

'===================================================================
'  LabTemplates — вставка шапок и блоков заданий лабораторных работ
'
'  Запускаемый макрос:  ЛабШаблоны
'  (назначь на кнопку/горячую клавишу — Word > Файл > Параметры >
'   Настройка ленты/сочетаний клавиш > Категория "Макросы")
'
'  Пункты меню:
'    1 - Вставка полной шапки (поля шапки + все задания + вывод + разрыв страницы)
'    2 - Вставка задания (только блоки "Задание N", без шапки)
'    3 - Выход
'
'  Отмена (кнопка "Отмена" в любом окне ввода) прерывает выполнение
'  макроса целиком, на любом шаге.
'===================================================================

Private redLineIndent As Single   ' отступ красной строки (в пунктах) — выбирается один раз за запуск

'-------------------------------------------------------------------
' ГЛАВНОЕ МЕНЮ
'-------------------------------------------------------------------
Sub ЛабШаблоны()
    Dim choice As Variant
    choice = GetNumericInput( _
        "Выберите действие:" & vbCrLf & _
        "1 - Вставка полной шапки" & vbCrLf & _
        "2 - Вставка задания" & vbCrLf & _
        "3 - Выход", _
        "Меню шаблонов лабораторных работ", "1", 1, 3)

    If VarType(choice) = vbBoolean Then Exit Sub ' отмена

    Select Case CLng(choice)
        Case 1
            InsertFullHeader
        Case 2
            InsertTasksOnly
        Case 3
            ' выход — ничего не делаем
    End Select
End Sub

'-------------------------------------------------------------------
' ПУНКТ 1 — ПОЛНАЯ ШАПКА
'-------------------------------------------------------------------
Sub InsertFullHeader()
    Dim indentChoice As Variant
    Dim formatChoice As Variant
    Dim taskCount As Variant
    Dim i As Long

    indentChoice = GetIndentChoice()
    If VarType(indentChoice) = vbBoolean Then Exit Sub
    redLineIndent = indentChoice

    formatChoice = GetHeaderFormat()
    If VarType(formatChoice) = vbBoolean Then Exit Sub

    taskCount = GetNumericInput("Сколько заданий в лабораторной работе?", "Количество заданий", "1", 1, 50)
    If VarType(taskCount) = vbBoolean Then Exit Sub

    InsertTitle
    If CLng(formatChoice) = 1 Then
        InsertHeaderFieldsTRPO
    Else
        InsertHeaderFieldsStandard
    End If
    InsertResultsSubheading

    For i = 1 To CLng(taskCount)
        If Not InsertOneTask(i) Then Exit Sub ' отмена где-то внутри — прерываем всё
    Next i

    If CLng(formatChoice) = 1 Then
        InsertParagraph "Вывод по лабораторной работе №:", True, wdAlignParagraphLeft, redLineIndent, 0, 0
    Else
        InsertParagraph "Вывод:", True, wdAlignParagraphLeft, redLineIndent, 0, 0
    End If
    Selection.InsertBreak wdPageBreak

    MsgBox "Шапка и " & CLng(taskCount) & " задание(-й) вставлены.", vbInformation, "Готово"
End Sub

'-------------------------------------------------------------------
' ПУНКТ 2 — ТОЛЬКО ЗАДАНИЯ (без шапки, без "Вывод", без разрыва страницы)
'-------------------------------------------------------------------
Sub InsertTasksOnly()
    Dim indentChoice As Variant
    Dim taskCount As Variant
    Dim i As Long

    indentChoice = GetIndentChoice()
    If VarType(indentChoice) = vbBoolean Then Exit Sub
    redLineIndent = indentChoice

    taskCount = GetNumericInput("Сколько заданий нужно вставить?", "Количество заданий", "1", 1, 50)
    If VarType(taskCount) = vbBoolean Then Exit Sub

    For i = 1 To CLng(taskCount)
        If Not InsertOneTask(i) Then Exit Sub
    Next i

    MsgBox CLng(taskCount) & " задание(-й) вставлено.", vbInformation, "Готово"
End Sub

'-------------------------------------------------------------------
' Заголовок "Лабораторная работа №" — общий для обоих форматов
'-------------------------------------------------------------------
Sub InsertTitle()
    InsertParagraph "Лабораторная работа №", True, wdAlignParagraphLeft, redLineIndent, 0, 0
End Sub

'-------------------------------------------------------------------
' Подзаголовок "Результаты выполнения лабораторной работы №" — общий
'-------------------------------------------------------------------
Sub InsertResultsSubheading()
    InsertParagraph "Результаты выполнения лабораторной работы №", True, wdAlignParagraphCenter, 0, 6, 6
End Sub

'-------------------------------------------------------------------
' Шапка — поля ТРПО
'-------------------------------------------------------------------
Sub InsertHeaderFieldsTRPO()
    InsertParagraph "Дата выполнения лабораторной работы №:", True, wdAlignParagraphLeft, redLineIndent, 0, 0
    InsertParagraph "Тема лабораторной работы №:", True, wdAlignParagraphLeft, redLineIndent, 0, 0
    InsertParagraph "Цель лабораторной работы №:", True, wdAlignParagraphLeft, redLineIndent, 0, 0
    InsertParagraph "Задание для выполнения лабораторной работы №:", True, wdAlignParagraphLeft, redLineIndent, 0, 0
    InsertParagraph "Оснащение лабораторной работы №:", True, wdAlignParagraphLeft, redLineIndent, 0, 0
End Sub

'-------------------------------------------------------------------
' Шапка — поля "Обычная"
'-------------------------------------------------------------------
Sub InsertHeaderFieldsStandard()
    InsertParagraph "Дата выполнения:", True, wdAlignParagraphLeft, redLineIndent, 0, 0
    InsertParagraph "Тема:", True, wdAlignParagraphLeft, redLineIndent, 0, 0
    InsertParagraph "Цель:", True, wdAlignParagraphLeft, redLineIndent, 0, 0
    InsertParagraph "Задание для выполнения:", True, wdAlignParagraphLeft, redLineIndent, 0, 0
    InsertParagraph "Оснащение:", True, wdAlignParagraphLeft, redLineIndent, 0, 0
End Sub

'-------------------------------------------------------------------
' Один блок "Задание N" целиком: заголовок + (листинги и/или рисунки)
' Возвращает False, если пользователь нажал "Отмена" — тогда вызывающий
' код должен прервать выполнение всего макроса.
'-------------------------------------------------------------------
Function InsertOneTask(taskNum As Long) As Boolean
    Dim listingsCount As Variant
    Dim figuresCount As Variant

    listingsCount = GetNumericInput("Задание " & taskNum & ". Сколько листингов?", _
                                     "Листинги — задание " & taskNum, "0", 0, 50)
    If VarType(listingsCount) = vbBoolean Then InsertOneTask = False: Exit Function

    figuresCount = GetNumericInput("Задание " & taskNum & ". Сколько рисунков?", _
                                    "Рисунки — задание " & taskNum, "1", 0, 50)
    If VarType(figuresCount) = vbBoolean Then InsertOneTask = False: Exit Function

    InsertParagraph "Задание " & taskNum & ".", True, wdAlignParagraphLeft, redLineIndent, 6, 6

    If CLng(listingsCount) = 0 Then
        InsertResultAndFigures taskNum, CLng(figuresCount)
    Else
        InsertListingsAndFigures taskNum, CLng(listingsCount), CLng(figuresCount)
    End If

    InsertOneTask = True
End Function

'-------------------------------------------------------------------
' Задание без листингов: результат + N рисунков (как в ТРПО)
'-------------------------------------------------------------------
Sub InsertResultAndFigures(taskNum As Long, figuresCount As Long)
    Dim i As Long
    If figuresCount <= 0 Then Exit Sub

    InsertResultLine taskNum, figuresCount
    For i = 1 To figuresCount
        InsertFigureBlock taskNum
    Next i
End Sub

'-------------------------------------------------------------------
' Задание с листингами.
' Каждому листингу достаётся минимум 1 рисунок по порядку, пока
' рисунки не закончатся; весь "лишний" остаток (если рисунков больше,
' чем листингов) уходит последнему листингу. Если рисунков меньше,
' чем листингов — у листингов "сверху" нет ни результата, ни рисунка.
'-------------------------------------------------------------------
Sub InsertListingsAndFigures(taskNum As Long, listingsCount As Long, figuresCount As Long)
    Dim i As Long, f As Long
    Dim assigned() As Long
    Dim remainder As Long

    ReDim assigned(1 To listingsCount)
    For i = 1 To listingsCount
        If i <= figuresCount Then
            assigned(i) = 1
        Else
            assigned(i) = 0
        End If
    Next i
    remainder = figuresCount - listingsCount
    If remainder > 0 Then assigned(listingsCount) = assigned(listingsCount) + remainder

    If listingsCount = 1 Then
        InsertParagraph "Код задания " & taskNum & " представлен в листинге ?", _
                         False, wdAlignParagraphLeft, redLineIndent, 6, 6
    Else
        InsertParagraph "Код задания " & taskNum & " представлен в листингах ? - ?", _
                         False, wdAlignParagraphLeft, redLineIndent, 6, 6
    End If

    For i = 1 To listingsCount
        InsertListingBlock taskNum
        If assigned(i) > 0 Then
            InsertResultLine taskNum, assigned(i)
            For f = 1 To assigned(i)
                InsertFigureBlock taskNum
            Next f
        End If
    Next i
End Sub

'-------------------------------------------------------------------
' Подпись листинга (без рамки) + рамка под код (Consolas 12, без отступа)
'-------------------------------------------------------------------
Sub InsertListingBlock(taskNum As Long)
    InsertParagraph "Листинг ? – Код задания " & taskNum, False, wdAlignParagraphLeft, 0, 0, 0
    InsertParagraph "", False, wdAlignParagraphLeft, 0, 0, 0, "Consolas", 12, True
End Sub

'-------------------------------------------------------------------
' "Результат выполнения задания N представлен на рисунке(-ах) ..."
'-------------------------------------------------------------------
Sub InsertResultLine(taskNum As Long, figuresInGroup As Long)
    If figuresInGroup = 1 Then
        InsertParagraph "Результат выполнения задания " & taskNum & " представлен на рисунке ?", _
                         False, wdAlignParagraphLeft, redLineIndent, 6, 6
    Else
        InsertParagraph "Результат выполнения задания " & taskNum & " представлен на рисунках ? - ?", _
                         False, wdAlignParagraphLeft, redLineIndent, 6, 6
    End If
End Sub

'-------------------------------------------------------------------
' Место под рисунок (пустой абзац по центру) + подпись под ним
'-------------------------------------------------------------------
Sub InsertFigureBlock(taskNum As Long)
    InsertParagraph "", False, wdAlignParagraphCenter, 0, 6, 6
    InsertParagraph "Рисунок ? – Результат выполнения задания " & taskNum, _
                     False, wdAlignParagraphCenter, 0, 6, 6
End Sub

'-------------------------------------------------------------------
' Универсальная вставка абзаца с нужным форматированием.
' Печатает текст в текущей позиции курсора и переходит на новую строку.
'-------------------------------------------------------------------
Sub InsertParagraph(text As String, bold As Boolean, alignment As WdParagraphAlignment, _
                     firstLineIndent As Single, spaceBefore As Single, spaceAfter As Single, _
                     Optional fontName As String = "Times New Roman", _
                     Optional fontSize As Single = 14, _
                     Optional withBorder As Boolean = False)
    With Selection
        .Font.Name = fontName
        .Font.Size = fontSize
        .Font.bold = bold
        .ParagraphFormat.alignment = alignment
        .ParagraphFormat.LeftIndent = 0
        .ParagraphFormat.firstLineIndent = firstLineIndent
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        .ParagraphFormat.spaceBefore = spaceBefore
        .ParagraphFormat.spaceAfter = spaceAfter
        SetBorder withBorder
        .TypeText text
        .TypeParagraph
    End With
End Sub

'-------------------------------------------------------------------
' Рамка (0.5pt, сплошная) вокруг текущего абзаца — вкл/выкл.
' Вызывается из InsertParagraph на каждом абзаце, поэтому рамка
' никогда не "утекает" в соседние абзацы.
'-------------------------------------------------------------------
Sub SetBorder(enable As Boolean)
    Dim sides As Variant
    Dim s As Variant
    sides = Array(wdBorderTop, wdBorderLeft, wdBorderBottom, wdBorderRight)
    For Each s In sides
        With Selection.ParagraphFormat.Borders(s)
            If enable Then
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
            Else
                .LineStyle = wdLineStyleNone
            End If
        End With
    Next s
End Sub

'-------------------------------------------------------------------
' Выбор отступа красной строки: 1.25 или 1.5 см
'-------------------------------------------------------------------
Function GetIndentChoice() As Variant
    Dim choice As String
    Do
        choice = InputBox("Какой отступ красной строки используется в этой работе?" & vbCrLf & _
                           "1 - 1.25 см" & vbCrLf & "2 - 1.5 см", "Отступ красной строки", "1")
        If StrPtr(choice) = 0 Then
            GetIndentChoice = False
            Exit Function
        End If
        Select Case choice
            Case "1"
                GetIndentChoice = CentimetersToPoints(1.25)
                Exit Function
            Case "2"
                GetIndentChoice = CentimetersToPoints(1.5)
                Exit Function
            Case Else
                MsgBox "Введите 1 или 2.", vbExclamation, "Ошибка ввода"
        End Select
    Loop
End Function

'-------------------------------------------------------------------
' Выбор формата шапки: ТРПО / Обычная
'-------------------------------------------------------------------
Function GetHeaderFormat() As Variant
    Dim choice As String
    Do
        choice = InputBox("Выберите формат шапки:" & vbCrLf & "1 - ТРПО" & vbCrLf & "2 - Обычная", _
                           "Формат шапки", "1")
        If StrPtr(choice) = 0 Then
            GetHeaderFormat = False
            Exit Function
        End If
        Select Case choice
            Case "1", "2"
                GetHeaderFormat = CLng(choice)
                Exit Function
            Case Else
                MsgBox "Введите 1 или 2.", vbExclamation, "Ошибка ввода"
        End Select
    Loop
End Function

'-------------------------------------------------------------------
' Универсальный запрос целого числа в диапазоне [minVal; maxVal].
' Возвращает Long при успехе или False (Boolean), если нажата "Отмена".
' Отмена определяется через StrPtr(raw) = 0 — это отличает нажатие
' "Отмена" (нулевой указатель на строку) от ввода реально пустой
' строки при нажатой "ОК" (валидный указатель на строку нулевой длины).
'-------------------------------------------------------------------
Function GetNumericInput(prompt As String, title As String, defaultVal As String, _
                          minVal As Long, maxVal As Long) As Variant
    Dim raw As String
    Do
        raw = InputBox(prompt, title, defaultVal)
        If StrPtr(raw) = 0 Then
            GetNumericInput = False
            Exit Function
        End If

        If InStr(raw, ".") > 0 Or InStr(raw, ",") > 0 Then
            MsgBox "Введите целое число, без дробной части.", vbExclamation, "Ошибка ввода"
        ElseIf Not IsNumeric(raw) Then
            MsgBox "Нужно ввести число.", vbExclamation, "Ошибка ввода"
        ElseIf CLng(raw) < minVal Or CLng(raw) > maxVal Then
            MsgBox "Число должно быть от " & minVal & " до " & maxVal & ".", vbExclamation, "Ошибка ввода"
        Else
            GetNumericInput = CLng(raw)
            Exit Function
        End If
    Loop
End Function
