Attribute VB_Name = "CleanTemplateModule"
Sub CleanTemplate()
Dim RootPath, ProjectPath, SourceFileName, ICDFileName, NSDFileName As String

  RootPath = ActiveWorkbook.Names("BaseFolderPath").RefersToRange.Value & "\"
  ProjectPath = RootPath & ActiveWorkbook.Names("ProjectName").RefersToRange.Value & "\"
  ICDFileName = ProjectPath & ActiveWorkbook.Names("ICDName").RefersToRange.Value & ".icd"

  Set doc = CreateObject("MSXML.DOMDocument")
  doc.async = False
  doc.Load ICDFileName
  If doc.parseError.ErrorCode Then
    ShowError doc.parseError
  Else
    
    lncnt = 0: docnt = 0: dacnt = 0: enumcnt = 0
    
    Set TemplatesRoot = doc.SelectSingleNode("//DataTypeTemplates")
    
    For Each LNtype In doc.getElementsByTagName("LNodeType")
        If LNtype.getAttribute("lnClass") = "LLN0" Then
            Set ImpNode = doc.SelectSingleNode("//LN0[@lnType='" & LNtype.getAttribute("id") & "']")
        Else
            Set ImpNode = doc.SelectSingleNode("//LN[@lnType='" & LNtype.getAttribute("id") & "']")
        End If
        If ImpNode Is Nothing Then
            TemplatesRoot.RemoveChild LNtype
            lncnt = lncnt + 1
        End If
    Next
    
    Do
      wasdel = False
      For Each DOtype In doc.getElementsByTagName("DOType")
        Set ImpDO = doc.SelectSingleNode("//DO[@type='" & DOtype.getAttribute("id") & "']")
        Set ImpSDO = doc.SelectSingleNode("//SDO[@type='" & DOtype.getAttribute("id") & "']")
        If (ImpDO Is Nothing) And (ImpSDO Is Nothing) Then
            TemplatesRoot.RemoveChild DOtype
            wasdel = True
            docnt = docnt + 1
        End If
      Next
    Loop While wasdel
    
    Do
      wasdel = False
      For Each DAtype In doc.getElementsByTagName("DAType")
        Set ImpDA = doc.SelectSingleNode("//DA[@type='" & DAtype.getAttribute("id") & "']")
        Set ImpBDA = doc.SelectSingleNode("//BDA[@type='" & DAtype.getAttribute("id") & "']")
        If (ImpDA Is Nothing) And (ImpBDA Is Nothing) Then
            TemplatesRoot.RemoveChild DAtype
            wasdel = True
            dacnt = dacnt + 1
        End If
      Next
    Loop While wasdel
    
    For Each Enumtype In doc.getElementsByTagName("EnumType")
        Set ImpDA = doc.SelectSingleNode("//DA[@type='" & Enumtype.getAttribute("id") & "']")
        Set ImpBDA = doc.SelectSingleNode("//BDA[@type='" & Enumtype.getAttribute("id") & "']")
        If (ImpDA Is Nothing) And (ImpBDA Is Nothing) Then
            TemplatesRoot.RemoveChild Enumtype
            enumcnt = enumcnt + 1
        End If
    Next
        
    doc.Save (ICDFileName)
  
  End If
  deletelist = "Удалено из DataTypeTemplates:" & vbCrLf _
    & "LNodeTypes: " & lncnt & vbCrLf _
    & "DOTypes: " & docnt & vbCrLf _
    & "DATypes: " & dacnt & vbCrLf _
    & "EnumTypes: " & enumcnt
  MsgBox deletelist
  Set doc = Nothing
End Sub

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
