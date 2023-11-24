Attribute VB_Name = "MicsInModule"
Dim doc, NSDdoc As Object ' парсеры двух xml объявлены на весь модуль
Dim doc1, doc2 As Object
Dim sdocs As String

Dim waserror As Boolean
Function CmpIcdObjects(objname) As Long
  For Each icdobj In doc1.getElementsByTagName(objname)
  ' для каждого LNtype в составе doc1 ищем такой же тип в составе doc2
     sicdobjId = icdobj.getAttribute("id")
     Set icdobj2 = doc2.SelectSingleNode("//" & objname & "[@id='" & sicdobjId & "']")
     If Not (icdobj2 Is Nothing) Then
        'проверяем соответствие по количеству childs
        m = icdobj.ChildNodes.Length
        If Not (m = icdobj2.ChildNodes.Length) Then
                MsgBox sicdobjId & " in " & sdocs
        End If
     End If
  Next icdobj
End Function

Sub CompareICDs()

Dim icdfiles() As Variant
icdfiles = Array("D:\ICD_LT.icd", "D:\ICD_TR.icd", "D:\ICD_CRN.icd", "D:\ICD_BMCS.icd", "D:\ICD_TD.icd", "D:\ICD_100.icd", "D:\ICD_BC.icd")

For i = LBound(icdfiles) To UBound(icdfiles)

For j = i To UBound(icdfiles)

If Not (i = j) Then
doc1name = icdfiles(i)

doc2name = icdfiles(j)

  Set doc1 = CreateObject("MSXML.DOMDocument")
  doc1.async = False
  doc1.Load doc1name
  If doc1.parseError.ErrorCode Then
    ShowError NSDdoc.parseError
  End If
        
  Set doc2 = CreateObject("MSXML.DOMDocument")
  doc2.async = False
  doc2.Load doc2name
  If doc2.parseError.ErrorCode Then
    ShowError NSDdoc.parseError
  End If
        
  sdocs = doc1name & " vs " & doc2name
        
  ' обработка LNTypes
  CmpIcdObjects ("LNodeType")
  CmpIcdObjects ("DOType")
  CmpIcdObjects ("DAType")
  CmpIcdObjects ("EnumType")
Set doc1 = Nothing
Set doc2 = Nothing

End If
Next j
Next i
  MsgBox "Check complete."
End Sub


Sub ReadTemplates()
  Dim RootPath, ProjectPath, SourceFileName, ICDFileName, NSDFileName As String

  RootPath = ActiveWorkbook.Names("BaseFolderPath").RefersToRange.Value & "\"
  ProjectPath = RootPath & ActiveWorkbook.Names("ProjectName").RefersToRange.Value & "\"
  ICDFileName = ProjectPath & ActiveWorkbook.Names("ICDName").RefersToRange.Value & ".icd"
  NSDFileName = RootPath & "\IEC_61850-7-4_2007A2.nsd"
  
  Set NSDdoc = CreateObject("MSXML.DOMDocument")
  NSDdoc.async = False
  NSDdoc.Load NSDFileName
  If NSDdoc.parseError.ErrorCode Then
    ShowError NSDdoc.parseError
  Else
    Set doc = CreateObject("MSXML.DOMDocument")
    doc.async = False
    doc.Load ICDFileName
    If doc.parseError.ErrorCode Then
        ShowError doc.parseError
    Else
        ' собственно обработка
        ' чистка листа
        
        ThisWorkbook.Worksheets("MICS").Unprotect Password:="111"
        lastRow = ThisWorkbook.Worksheets("MICS").Cells.SpecialCells(xlLastCell).Row
        ThisWorkbook.Worksheets("MICS").Protect Password:="111", AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True

        ThisWorkbook.Worksheets("MICS").Range("2:" & lastRow).Cells = ""
        ' начальная строка вывода данных
        idx = 2: waserror = False
        For Each LD In doc.getElementsByTagName("LDevice")
          sLDinst = LD.getAttribute("inst"): sLDdesc = LD.getAttribute("desc")
          idx = ParseLns(doc.SelectNodes("//LDevice[@inst='" & sLDinst & "']/LN0"), sLDinst, sLDdesc, idx)
          idx = ParseLns(doc.SelectNodes("//LDevice[@inst='" & sLDinst & "']/LN"), sLDinst, sLDdesc, idx)
        Next LD
    End If
  End If
  ' чистка
  Set doc = Nothing
  Set NSDdoc = Nothing
  ThisWorkbook.Save
  If waserror Then MsgBox "Имеются ошибки обработки данных!" & vbCrLf & "Проверьте таблицу MICS!", vbExclamation
End Sub

Function ParseLns(LnsCollection, LDinst, LDdesc, idx) As Long
' idx -  это номер строки таблицы, начиная с которой надо вставлять данные
' функция пшет строки по очереди и возвращает номер следующей свободной строки
    For Each LN In LnsCollection
      sLnName = LN.getAttribute("prefix") & LN.getAttribute("lnClass") & LN.getAttribute("inst")
      sLnType = LN.getAttribute("lnType")
      sLnDesc = LN.getAttribute("desc") ' это описание узла из секции имплементации
      
      Set LNodeType = doc.SelectSingleNode("//LNodeType[@id='" & sLnType & "']")
      If LNodeType Is Nothing Then
            ThisWorkbook.Worksheets("MICS").Cells(idx, 3) = sLnName
            ThisWorkbook.Worksheets("MICS").Cells(idx, 5) = "Нет в Templates"
            waserror = True
            idx = idx + 1
        Else
          If sLnDesc = "" Then sLnDesc = LNodeType.getAttribute("desc") ' если в имплементации не было описания, то берем описание из шаблона
          sLnClass = LNodeType.getAttribute("lnClass")
          For Each DoNode In LNodeType.ChildNodes ' перебираем DO и форматируем их информацию для MICS
            ThisWorkbook.Worksheets("MICS").Cells(idx, 1) = LDinst
            ThisWorkbook.Worksheets("MICS").Cells(idx, 2) = LDdesc
            ThisWorkbook.Worksheets("MICS").Cells(idx, 3) = sLnName
            ThisWorkbook.Worksheets("MICS").Cells(idx, 4) = sLnClass
            ThisWorkbook.Worksheets("MICS").Cells(idx, 5) = sLnDesc
            ThisWorkbook.Worksheets("MICS").Cells(idx, 6) = getlngroup(sLnClass)
            ThisWorkbook.Worksheets("MICS").Cells(idx, 7) = getlnsort(sLnClass)
            
            sDoName = DoNode.getAttribute("name")
            sDoType = DoNode.getAttribute("type")
            ' надо отделить индекс из состава DOName для последующей правильной сортировки Ind1,Ind10,Ind2 итп
            ' и для последующего анализа Omulti
            sDoName1 = sDoName: sDoName2 = ""
            bHasMulti = False
            Do While ISIN(Right(sDoName1, 1), Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")) '- тупо, но работает
                  sDoName2 = Right(sDoName1, 1) & sDoName2
                  sDoName1 = Left(sDoName1, Len(sDoName1) - 1)
                  bHasMulti = True ' есть циферки
                Loop
            ' CDC объекта вычисляется из его типа, задаваемого строго по схеме [CDC_]AnySuffix или [CDCEx]AnySuffix
            pos1 = InStr(1, sDoType, "_", vbTextCompare): If pos1 = 0 Then pos1 = Len(sDoType)
            pos2 = InStr(1, sDoType, "Ex", vbTextCompare): If pos2 = 0 Then pos2 = Len(sDoType)
            sCDCsmall = Left(sDoType, WorksheetFunction.Min(pos1, pos2) - 1)
            sCdc = UCase(sCDCsmall)
            sorder = 0
            ThisWorkbook.Worksheets("MICS").Cells(idx, 8) = GetUsage(sCdc, sorder) ' определение функциональной группы по CDC
            ThisWorkbook.Worksheets("MICS").Cells(idx, 18) = sorder ' особый порядок сортировки CDC
            'Если объект контролируемый - определим его модель управления по templates
            Set DActrl = doc.SelectSingleNode("//DOType[@id='" & sDoType & "']/DA[@name='ctlModel']/Val")
            If DActrl Is Nothing Then ctrlmodel = "" Else ctrlmodel = DActrl.Text
            If ctrlmodel <> "status-only" Then ctrlmodel = ""
            ThisWorkbook.Worksheets("MICS").Cells(idx, 19) = ctrlmodel
            
            ThisWorkbook.Worksheets("MICS").Cells(idx, 9) = sDoName
            ThisWorkbook.Worksheets("MICS").Cells(idx, 16) = sDoName1
            ThisWorkbook.Worksheets("MICS").Cells(idx, 17) = sDoName2
            ThisWorkbook.Worksheets("MICS").Cells(idx, 10) = LDinst & "/" & sLnName & "/" & sDoName
            ThisWorkbook.Worksheets("MICS").Cells(idx, 11) = sDoType
            ThisWorkbook.Worksheets("MICS").Cells(idx, 12) = sCdc
            ThisWorkbook.Worksheets("MICS").Cells(idx, 13) = DoNode.getAttribute("desc")
            ' определение prescond выполняется через обращение к NSD-схеме, и по наличию суффикса Ex в имени типа DO
            bIsEx = (InStr(1, sDoType, sCDCsmall & "Ex", vbTextCompare) > 0) ' Флаг что разработчик считал это как Ex
            bIsMulti = False ' дефолтный флаг применения OMulti
            ' ищем в NSD по полному наименованию - с учетом sLnClass!!!
            
            sMOCE = FindPresCond(sLnClass, sDoName, sCdc)
            
            ' если не нашли - надо убрать цифры в конце имени и поискать снова, возожно это OMulty
            If sMOCE = "" Then
                sDoName = sDoName1
                bIsMulti = bHasMulti ' разработчик сделал OMulti
                sMOCE = FindPresCond(sLnClass, sDoName, sCdc)
            End If
            ' проверяем варианты соответствия заданного Ex и определений NSD
            posmul = InStr(1, sMOCE, "multi", vbTextCompare)
            If bIsEx Then
                If sMOCE = "" Then
                    sMOCE = "E": sMsg = ""
                    Else
                    sMsg = sMOCE & " - не является EX": sMOCE = "ERROR"
                    waserror = True
                End If
                Else
                If sMOCE = "" Then
                    sMOCE = "ERROR": sMsg = "Объект " & sDoName & " не стандартный. Должно быть Ex?"
                    waserror = True
                    Else
                    sMOCE = Left(sMOCE, 1): sMsg = ""
                    If Not ISIN(sMOCE, Array("M", "O", "C", "E")) Then sMOCE = "C"
                End If
            End If
            ' проверяем корректность применения Multi
            If (sMsg = "") And bIsMulti And (posmul = 0) And (sMOCE <> "E") Then
                sMOCE = "ERROR": sMsg = "Индексация объекта запрещена"
                waserror = True
            End If
                        
            ThisWorkbook.Worksheets("MICS").Cells(idx, 14) = sMOCE
            ThisWorkbook.Worksheets("MICS").Cells(idx, 15) = sMsg
                        
            idx = idx + 1
          Next DoNode
        End If
      Next LN
    ParseLns = idx
End Function

Function FindPresCond(sClassName, sDoName, sCdc) As String
' если сразу не нашли - надо проверять в абстрактных родительских узлах заданного класса
    Dim ownclassname As String
    ownclassname = sClassName
    Do
        Set nsddef = NSDdoc.SelectSingleNode("//LNClass[(@name='" & ownclassname & "')]/DataObject[(@name='" & sDoName & "' and @type='" & sCdc & "')]")
        If nsddef Is Nothing Then Set nsddef = NSDdoc.SelectSingleNode("//AbstractLNClass[(@name='" & ownclassname & "')]/DataObject[(@name='" & sDoName & "' and @type='" & sCdc & "')]")
        If nsddef Is Nothing Then
            Set ownclass = NSDdoc.SelectSingleNode("//LNClass[(@name='" & ownclassname & "')]")
            If ownclass Is Nothing Then ownclassname = "" Else ownclassname = ownclass.getAttribute("base")
        End If
    Loop Until (Not (nsddef Is Nothing)) Or (ownclassname = "")
    
    If nsddef Is Nothing Then FindPresCond = "" Else FindPresCond = nsddef.getAttribute("presCond")
    
End Function

Function ShowError(XMLDOMParseError)
    mess = _
    "parseError.errorCode: " & XMLDOMParseError.ErrorCode & vbCrLf & _
    "parseError.filepos: " & XMLDOMParseError.filepos & vbCrLf & _
    "parseError.line: " & XMLDOMParseError.Line & vbCrLf & _
    "parseError.linepos: " & XMLDOMParseError.linepos & vbCrLf & _
    "parseError.reason: " & XMLDOMParseError.reason & vbCrLf & _
    "parseError.srcText: " & XMLDOMParseError.srcText & vbCrLf & _
    "parseError.url: " & XMLDOMParseError.Url & vbCrLf
    MsgBox mess
End Function

Function ISIN(x, StringSetElementsAsArray)
    ISIN = InStr(1, Join(StringSetElementsAsArray, "'"), _
    x, vbTextCompare) > 0
End Function

Function GetUsage(CDC, ByRef sord) As String
    res = "Unknown": sord = 99
    If ISIN(CDC, Array("SPS", "DPS", "INS", "ENS", "ACT", "ACD", "SEC", "BCR", "HST", "VSS")) Then
        res = "Status"
        sord = 20
        End If
    If ISIN(CDC, Array("MV", "CMV", "SAV", "WYE", "DEL", "SEQ", "HMV", "HWYE", "HDEL")) Then
        res = "Measurements"
        sord = 25
        End If
    If ISIN(CDC, Array("SPC", "DPC", "INC", "ENC", "BSC", "ISC", "APC", "BAC")) Then
        res = "Controls"
        sord = 30
        End If
    If ISIN(CDC, Array("SPG", "ING", "ENG", "ORG", "TSG", "CUG", "VSG", "ASG", "CURVE", "CSG")) Then
        res = "Settings"
        sord = 40
        End If
    If ISIN(CDC, Array("DPL", "LPL", "CSD")) Then
        res = "Description"
        sord = 10
        End If
    GetUsage = res
End Function

Function getlngroup(lnclass) As String
    Select Case Left(lnclass, 1)
        Case "A"
        getlngroup = "Automatic control"
        Case "C"
        getlngroup = "Supervisory Control"
        Case "D"
        getlngroup = "Distributed energy resources"
        Case "F"
        getlngroup = "Functional blocks"
        Case "G"
        getlngroup = "Generic function references"
        Case "H"
        getlngroup = "Hydro power"
        Case "I"
        getlngroup = "Interfacing and archiving"
        Case "K"
        getlngroup = "Mechanical and non-electrical primary equipment"
        Case "L"
        getlngroup = "System logical nodes"
        Case "M"
        getlngroup = "Metering and measurement"
        Case "P"
        getlngroup = "Protection functions"
        Case "Q"
        getlngroup = "Power quality events detection related"
        Case "R"
        getlngroup = "Protection related functions"
        Case "S"
        getlngroup = "Supervision and monitoring"
        Case "T"
        getlngroup = "Instrument transformer and sensors"
        Case "W"
        getlngroup = "Wind power"
        Case "X"
        getlngroup = "Switchgear"
        Case "Y"
        getlngroup = "Power transformer and related functions"
        Case "Z"
        getlngroup = "Further (power system) equipment"
        Case Else
        getlngroup = "Reserved"
    End Select
End Function

Function getlnsort(lnclass) As String
  If Left(lnclass, 1) = "L" Then
    getlnsort = "0" & lnclass
    Else
    getlnsort = lnclass
  End If
End Function

Sub AssociateData()
  Dim RootPath, ProjectPath As String

  RootPath = ActiveWorkbook.Names("BaseFolderPath").RefersToRange.Value & "\"
  ProjectPath = RootPath & ActiveWorkbook.Names("ProjectName").RefersToRange.Value & "\"
  
  Dim ARMBook As Workbook
  Set ARMBook = Workbooks.Open(Filename:=ProjectPath & "armdef.txt", UpdateLinks:=False, ReadOnly:=True)
  ThisWorkbook.Activate
  Dim fcell As Object
  
  MismatchFlag = False
  sWarnMsg = "Расхождения таблицы ассоциаций с моделью:"
  
  ThisWorkbook.Worksheets("MICS").Unprotect Password:="111"
  maxrow = ThisWorkbook.Worksheets("MICS").Cells.SpecialCells(xlLastCell).Row
  ThisWorkbook.Worksheets("MICS").Protect Password:="111", AllowSorting:=True, AllowFiltering:=True, AllowFormattingColumns:=True

  For nrow = 2 To maxrow
  If ThisWorkbook.Worksheets("MICS").Cells(nrow, 10).Value <> "" Then
    sTxt = ""
    Set fcell = ARMBook.ActiveSheet.Columns(1).Find _
        (ThisWorkbook.Worksheets("MICS").Cells(nrow, 10).Value & "/", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    act = "this": first = "this"
    If fcell Is Nothing Then
        MismatchFlag = True
        sWarnMsg = sWarnMsg & vbNewLine & "'" & ThisWorkbook.Worksheets("MICS").Cells(nrow, 10).Value & "'"
    Else: first = fcell.Address
    End If
    Do While Not (act = first)
        If (ARMBook.ActiveSheet.Cells(fcell.Row, 4) <> "") _
           And (ARMBook.ActiveSheet.Cells(fcell.Row, 3) <> "QUALITY") _
           And (InStr(1, ARMBook.ActiveSheet.Cells(fcell.Row, 4), "_", vbTextCompare) = 0) Then
           If Not (sTxt = "") Then sTxt = sTxt & "; "
           sTxt = sTxt & ARMBook.ActiveSheet.Cells(fcell.Row, 4)
        End If
        Set fcell = ARMBook.ActiveSheet.Columns(1).FindNext(fcell)
        act = fcell.Address
    Loop
    ThisWorkbook.Worksheets("MICS").Cells(nrow, 15) = sTxt
    End If
  Next nrow
  ARMBook.Close False
  ThisWorkbook.Save
  If MismatchFlag Then MsgBox sWarnMsg, vbInformation
End Sub
