Attribute VB_Name = "MICSOutModule"
Sub BWMicsTables()
    Call CreateMicsTables(False)
    MsgBox "Создание таблиц MICS-1 и MICS-2 завершено."
    
End Sub

Sub ColorMicsTables()
    Call CreateMicsTables(True)
    MsgBox "Создание таблиц MICS-1 и MICS-2" & vbCrLf & "выполнено, хозяин!"

End Sub

Sub CallToWordTemplate()
    docpath = ActiveWorkbook.Names("BaseFolderPath").RefersToRange.Value & "\MICS template.dotm"
    CreateObject("Shell.Application").Open (docpath)
End Sub

Sub CreateMicsTables(usecolor)
Attribute CreateMicsTables.VB_ProcData.VB_Invoke_Func = " \n14"

    If usecolor Then
    ' Цветная схема таблиц
        clHeader = RGB(0, 90, 173)
        clSubheader = RGB(79, 129, 189)
        clInline1 = RGB(208, 216, 232)
        clInline2 = RGB(233, 237, 244)
        clHilight1 = RGB(255, 255, 153)
        clHilight2 = RGB(255, 255, 202)
    Else
    ' Ч/б схема таблиц
        clHeader = RGB(0, 0, 0)
        clSubheader = RGB(89, 89, 89)
        clInline1 = RGB(255, 255, 255)
        clInline2 = RGB(255, 255, 255)
        clHilight1 = RGB(191, 191, 191)
        clHilight2 = RGB(191, 191, 191)
    End If
    

' Сортируем данные
    ThisWorkbook.Worksheets("MICS").Unprotect Password:="111"
    srcLastRow = ThisWorkbook.Worksheets("MICS").Cells.SpecialCells(xlLastCell).Row
    ThisWorkbook.Worksheets("MICS").Protect Password:="111", AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True
    
    ThisWorkbook.Worksheets("MICS").Sort.SortFields.Clear
    ThisWorkbook.Worksheets("MICS").Sort.SortFields.Add Key:=Range("A2:A" & srcLastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ThisWorkbook.Worksheets("MICS").Sort.SortFields.Add Key:=Range("G2:G" & srcLastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ThisWorkbook.Worksheets("MICS").Sort.SortFields.Add Key:=Range("C2:C" & srcLastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ThisWorkbook.Worksheets("MICS").Sort.SortFields.Add Key:=Range("R2:R" & srcLastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ThisWorkbook.Worksheets("MICS").Sort.SortFields.Add Key:=Range("P2:P" & srcLastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ThisWorkbook.Worksheets("MICS").Sort.SortFields.Add Key:=Range("Q2:Q" & srcLastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ThisWorkbook.Worksheets("MICS").Sort
        .SetRange Range("A1:S" & srcLastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Application.ScreenUpdating = False
    On Error GoTo restoreupdate1
    
    ' чистка таблицы MICS-1
    ThisWorkbook.Worksheets("MICS-1").Range("A:A").Cells.Clear
    ' по очереди перебираем и пишем в MICS-1
    actLD = "": actGr = "": actLN = ""
    Row = 1
    For idx = 2 To srcLastRow
      If ThisWorkbook.Worksheets("MICS").Cells(idx, 3).Value <> "" Then
      If ThisWorkbook.Worksheets("MICS").Cells(idx, 1).Value <> actLD Then
        If actLD <> "" Then
            ThisWorkbook.Worksheets("MICS-1").Cells(Row, 1).Font.Size = 10
            Row = Row + 1
        End If
        actGr = "": actLN = ""
        actLD = ThisWorkbook.Worksheets("MICS").Cells(idx, 1).Value
        sTxt = ThisWorkbook.Worksheets("MICS").Cells(idx, 2).Value
        If sTxt <> "" Then sTxt = actLD & " (" & sTxt & ")" Else sTxt = actLD
        With ThisWorkbook.Worksheets("MICS-1").Cells(Row, 1)
            .Value = "Логическое устройство " & sTxt
            .HorizontalAlignment = xlHAlignCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = clHeader
        End With
        ' форматирование заголовка LD
        Row = Row + 1
      End If
      If Left(ThisWorkbook.Worksheets("MICS").Cells(idx, 7).Value, 1) <> actGr Then
        actLN = ""
        actGr = Left(ThisWorkbook.Worksheets("MICS").Cells(idx, 7).Value, 1)
        If actGr = "0" Then sactGr = "L" Else sactGr = actGr
        sTxt = sactGr & ": " & ThisWorkbook.Worksheets("MICS").Cells(idx, 6).Value
        With ThisWorkbook.Worksheets("MICS-1").Cells(Row, 1)
            .Value = sTxt
            .HorizontalAlignment = xlHAlignLeft
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = clSubheader
        End With
        ' форматирование заголовка группы
        Row = Row + 1
      End If
      If ThisWorkbook.Worksheets("MICS").Cells(idx, 3).Value <> actLN Then
        actLN = ThisWorkbook.Worksheets("MICS").Cells(idx, 3).Value
        sTxt = ThisWorkbook.Worksheets("MICS").Cells(idx, 5).Value
        If sTxt <> "" Then sTxt = actLN & " (" & sTxt & ")" Else sTxt = actLN
        With ThisWorkbook.Worksheets("MICS-1").Cells(Row, 1)
            .Value = sTxt
            .HorizontalAlignment = xlHAlignLeft
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Size = 10
            .Font.Bold = False
            .Font.Color = vbBlack
            If (Row Mod 2) = 0 Then .Interior.Color = clInline1 Else .Interior.Color = clInline2
        End With
        ' форматирование строки узла
        Row = Row + 1
      End If
      
      End If
    Next idx
    
restoreupdate1:
    Application.ScreenUpdating = True
    
    ' follow next
    
    Application.ScreenUpdating = False
    On Error GoTo restoreupdate2
    
    ' чистка таблицы MICS-2
    ThisWorkbook.Worksheets("MICS-2").Range("A:E").Cells.Clear
    ' по очереди перебираем и пишем в MICS-2
    actLD = "": actLN = "": actGr = "":
    Row = 1
    
    'Заголовок таблицы LD
    For idx = 2 To srcLastRow
      If ThisWorkbook.Worksheets("MICS").Cells(idx, 3).Value <> "" Then
      If ThisWorkbook.Worksheets("MICS").Cells(idx, 1).Value <> actLD Then
        If actLD <> "" Then
            ThisWorkbook.Worksheets("MICS-2").Cells(Row, 1).Resize(, 5).Merge
            ThisWorkbook.Worksheets("MICS-2").Cells(Row, 1).Font.Size = 12
            Row = Row + 1
        End If
        actLN = "": actGr = ""
        actLD = ThisWorkbook.Worksheets("MICS").Cells(idx, 1).Value
        sTxt = ThisWorkbook.Worksheets("MICS").Cells(idx, 2).Value
        If sTxt <> "" Then sTxt = actLD & " (" & sTxt & ")" Else sTxt = actLD
        With ThisWorkbook.Worksheets("MICS-2").Cells(Row, 1)
            .Resize(, 5).Merge
            .Value = "Логические узлы в составе " & sTxt
            .HorizontalAlignment = xlHAlignLeft
            .Font.Size = 12
            .Font.Bold = True
        End With
        Row = Row + 1
      End If
      
      'Заголовок таблицы LN
      If ThisWorkbook.Worksheets("MICS").Cells(idx, 3).Value <> actLN Then
        If actLN <> "" Then
            ThisWorkbook.Worksheets("MICS-2").Cells(Row, 1).Resize(, 5).Merge
            ThisWorkbook.Worksheets("MICS-2").Cells(Row, 1).Font.Size = 12
            Row = Row + 1
        End If
        actGr = ""
        actLN = ThisWorkbook.Worksheets("MICS").Cells(idx, 3).Value
        sTxt = ThisWorkbook.Worksheets("MICS").Cells(idx, 5).Value
        If sTxt <> "" Then sTxt = " (" & sTxt & ")"
        With ThisWorkbook.Worksheets("MICS-2").Cells(Row, 1)
            .Resize(, 5).Merge
            .Value = "Логический узел """ & actLN & """ " & sTxt
            .HorizontalAlignment = xlHAlignLeft
            .Font.Size = 12
            .Font.Bold = True
        End With
        Row = Row + 1
        
        With ThisWorkbook.Worksheets("MICS-2").Range _
            (ThisWorkbook.Worksheets("MICS-2").Cells(Row, 1), _
            ThisWorkbook.Worksheets("MICS-2").Cells(Row, 5))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Merge
        End With
        With ThisWorkbook.Worksheets("MICS-2").Cells(Row, 1)
            .Value = ThisWorkbook.Worksheets("MICS").Cells(idx, 4).Value & " class"
            .HorizontalAlignment = xlHAlignCenter
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = clHeader
        End With
        Row = Row + 1
        With ThisWorkbook.Worksheets("MICS-2")
            .Cells(Row, 1).Value = "Data object name"
            .Cells(Row, 2).Value = "Common data class"
            .Cells(Row, 3).Value = "Explanation"
            .Cells(Row, 4).Value = "M/O/C/E"
            .Cells(Row, 5).Value = "Remarks"
        End With
        With ThisWorkbook.Worksheets("MICS-2").Range _
            (ThisWorkbook.Worksheets("MICS-2").Cells(Row, 1), _
            ThisWorkbook.Worksheets("MICS-2").Cells(Row, 5))
            .HorizontalAlignment = xlHAlignLeft
            .VerticalAlignment = xlVAlignCenter
            .WrapText = True
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = clSubheader
        End With
        Row = Row + 1
        With ThisWorkbook.Worksheets("MICS-2")
            .Cells(Row, 1).Value = "LNName"
            .Cells(Row, 3).Value = "LN-Prefix, class name and LN-Instance-ID"
            .Cells(Row, 4).Value = "M"
            .Cells(Row, 5).Value = actLN
        End With
        With ThisWorkbook.Worksheets("MICS-2").Range _
            (ThisWorkbook.Worksheets("MICS-2").Cells(Row, 1), _
            ThisWorkbook.Worksheets("MICS-2").Cells(Row, 5))
            .HorizontalAlignment = xlHAlignLeft
            .VerticalAlignment = xlVAlignCenter
            .WrapText = True
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Size = 10
            .Font.Bold = False
            .Font.Color = vbBlack
            .Interior.Color = clInline2
        End With
        Row = Row + 1
        
        With ThisWorkbook.Worksheets("MICS-2").Range _
            (ThisWorkbook.Worksheets("MICS-2").Cells(Row, 1), _
            ThisWorkbook.Worksheets("MICS-2").Cells(Row, 5))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Merge
        End With
        With ThisWorkbook.Worksheets("MICS-2").Cells(Row, 1)
            .Value = "Data objects"
            .HorizontalAlignment = xlHAlignLeft
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = clSubheader
        End With
        Row = Row + 1
      End If
        
        ' заголовок группы объектов
      If ThisWorkbook.Worksheets("MICS").Cells(idx, 8).Value <> actGr Then
        actGr = ThisWorkbook.Worksheets("MICS").Cells(idx, 8).Value
        With ThisWorkbook.Worksheets("MICS-2").Range _
            (ThisWorkbook.Worksheets("MICS-2").Cells(Row, 1), _
            ThisWorkbook.Worksheets("MICS-2").Cells(Row, 5))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Merge
        End With
        With ThisWorkbook.Worksheets("MICS-2").Cells(Row, 1)
            .Value = actGr
            .HorizontalAlignment = xlHAlignLeft
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = vbBlack
            If (Row Mod 2) = 0 Then .Interior.Color = clInline1 Else .Interior.Color = clInline2
        End With
        Row = Row + 1
      End If
        
      ' описание объекта
      With ThisWorkbook.Worksheets("MICS-2")
            .Cells(Row, 1).Value = ThisWorkbook.Worksheets("MICS").Cells(idx, 9).Value
            .Cells(Row, 2).Value = ThisWorkbook.Worksheets("MICS").Cells(idx, 12).Value
            .Cells(Row, 3).Value = ThisWorkbook.Worksheets("MICS").Cells(idx, 13).Value
            .Cells(Row, 4).Value = ThisWorkbook.Worksheets("MICS").Cells(idx, 14).Value
            If ThisWorkbook.Worksheets("MICS").Cells(idx, 19).Value = "" Then
                .Cells(Row, 5).Value = ThisWorkbook.Worksheets("MICS").Cells(idx, 15).Value
                Else
                If ThisWorkbook.Worksheets("MICS").Cells(idx, 15).Value = "" Then
                    .Cells(Row, 5).Value = ThisWorkbook.Worksheets("MICS").Cells(idx, 19).Value
                    Else
                    .Cells(Row, 5).Value = ThisWorkbook.Worksheets("MICS").Cells(idx, 15).Value & ", " & ThisWorkbook.Worksheets("MICS").Cells(idx, 19).Value
                End If
            End If
      End With
      With ThisWorkbook.Worksheets("MICS-2").Range _
            (ThisWorkbook.Worksheets("MICS-2").Cells(Row, 1), _
            ThisWorkbook.Worksheets("MICS-2").Cells(Row, 5))
            .HorizontalAlignment = xlHAlignLeft
            .VerticalAlignment = xlVAlignCenter
            .WrapText = True
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Font.Size = 10
            .Font.Bold = False
            .Font.Color = vbBlack
            If ThisWorkbook.Worksheets("MICS").Cells(idx, 14).Value = "E" Then
              If (Row Mod 2) = 0 Then .Interior.Color = clHilight1 Else .Interior.Color = clHilight2
            Else
              If (Row Mod 2) = 0 Then .Interior.Color = clInline1 Else .Interior.Color = clInline2
            End If
      End With
      Row = Row + 1
        
      
      End If ' of common row presence check
    Next idx
restoreupdate2:
    Application.ScreenUpdating = True
    
    ThisWorkbook.Save
    
End Sub
