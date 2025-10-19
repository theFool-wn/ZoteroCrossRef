Attribute VB_Name = "ZoteroCrossRef"
' ZoteroCrossRef
' Establish Cross-references between Zotero citations and the bibliography
' Do not support Bookmark-type citations
' Supported citation styles:
'     Numeric     : [1,3,5-9]  |  1,3,5-9  |  [1],[3],[5-9]
'     Author-Year : (Smith, 2002; Li et al., 2025)
' Note: Automatically handle Chinese and English commas and parentheses.
'       If the hyperlink color and underline are not set correctly, just run the macro again.
'       It is recommended to run the macro after the final draft is completed
'           (i.e., citations and references are complete, and the citation format is confirmed),
'           and to save a copy before running it.
'
' Created by WangNan, 2025.10.18 - 2025.10.19
' Revised by WangNan, 2025.10.19
'
' Contract: wang.nan@buaa.edu.cn / me@wangnan.net
' Reference: https://github.com/altairwei/ZoteroLinkCitation
'            https://blog.csdn.net/Bearingz/article/details/146242667
'            https://blog.csdn.net/eternity_memory/article/details/150343285

Option Explicit

Public Sub ZoteroCrossRef()
    Dim nStart&, nEnd&
    nStart = Selection.Start
    nEnd = Selection.End
    
    Dim userChoice As VbMsgBoxResult
    userChoice = MsgBox("请选择引用样式：" & vbCrLf & vbCrLf & _
                        "是 - 顺序编码（Numeric）" & vbCrLf & _
                        "否 - 作者-年份（Author-Year）", _
                        vbYesNoCancel + vbQuestion, "选择引用样式")
    
    If userChoice = vbCancel Then Exit Sub
    
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    If Not ActiveDocument.Bookmarks.Exists("Zotero_Bibliography") Then
        ActiveWindow.View.ShowFieldCodes = True
        Selection.Find.ClearFormatting
        With Selection.Find
            .Text = "^d ADDIN ZOTERO_BIBL"
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute
        With ActiveDocument.Bookmarks
            .Add Range:=Selection.Range, Name:="Zotero_Bibliography"
            .DefaultSorting = wdSortByName
            .ShowHidden = True
        End With
        ActiveWindow.View.ShowFieldCodes = False
    End If
    
    If userChoice = vbYes Then
        ProcessNumberCitations
    Else
        ProcessAuthorYearCitations
    End If
    
    
    ActiveDocument.Range(nStart, nEnd).Select
    Application.ScreenUpdating = True
    MsgBox "Zotero交叉引用链接完成！", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "处理过程中出现错误: " & Err.Description, vbCritical
End Sub

Private Sub ProcessNumberCitations()
    Dim fieldCode As String
    Dim Paper_index As Integer
    Dim titles() As String
    Dim Num As String
    Dim NumParts() As String
    Dim dashParts() As String
    Dim field As field
    Dim NumPart As Variant
    
    For Each field In ActiveDocument.Fields
        If InStr(field.code, "ADDIN ZOTERO_ITEM") > 0 Then
            fieldCode = field.code
            Paper_index = 0
            titles = GetTitles(fieldCode)
            
            field.Select
            Num = Selection.Range.Text
            Num = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Num, " ", ""), "[", ""), "]", ""), "，", ","), "—", "-"), "‐", "-"), "–", "-"), "(", ""), ")", "")
            NumParts = Split(Num, ",")
            
            For Each NumPart In NumParts
                If InStr(NumPart, "-") > 0 Then
                    dashParts = Split(NumPart, "-")
                    Paper_index = Paper_index + 1
                    InsertCrossRef dashParts(0), field, titles(Paper_index)
                    Paper_index = Paper_index + CLng(dashParts(1)) - CLng(dashParts(0))
                    InsertCrossRef dashParts(1), field, titles(Paper_index)
                Else
                    Paper_index = Paper_index + 1
                    InsertCrossRef NumPart, field, titles(Paper_index)
                End If
            Next NumPart
        End If
    Next field
End Sub

Private Sub ProcessAuthorYearCitations()
    Dim linkChoice As VbMsgBoxResult
    linkChoice = MsgBox("请选择链接方式：" & vbCrLf & vbCrLf & _
                        "是 - 链接整个引用" & vbCrLf & _
                        "否 - 只链接年份", _
                        vbYesNo + vbQuestion, "选择链接方式")
    
    Dim fieldCode As String
    Dim Paper_index As Integer
    Dim titles() As String
    Dim AuthorYear As String
    Dim AuthorYearParts() As String
    Dim CommaParts() As String
    Dim field As field
    Dim part As Variant
    
    For Each field In ActiveDocument.Fields
        If InStr(field.code, "ADDIN ZOTERO_ITEM") > 0 Then
            fieldCode = field.code
            Paper_index = 0
            titles = GetTitles(fieldCode)
            
            field.Select
            AuthorYear = Selection.Range.Text
            AuthorYearParts = Split(Replace(Replace(Replace(Replace(Replace(AuthorYear, "；", ";"), "(", ""), ")", ""), "（", ""), "）", ""), ";")
            
            If linkChoice = vbYes Then
                For Each part In AuthorYearParts
                    Paper_index = Paper_index + 1
                    InsertCrossRef Trim(part), field, titles(Paper_index)
                Next part
            Else
                For Each part In AuthorYearParts
                    CommaParts = Split(Replace(part, "，", ","), ",")
                    Paper_index = Paper_index + 1
                    InsertCrossRef Trim(part), field, titles(Paper_index), Trim(CommaParts(UBound(CommaParts)))
                Next part
            End If
        End If
    Next field
End Sub

Private Function GetTitles(fieldCode As String) As String()
    Dim n1 As Long
    Dim n2 As Long
    Dim count As Integer
    Dim titles() As String
    Dim title As String
    
    count = 0
    Do While InStr(fieldCode, """title"":""") > 0
        n1 = InStr(fieldCode, """title"":""") + Len("""title"":""")
        n2 = InStr(Mid(fieldCode, n1, Len(fieldCode) - n1), """,""") - 1 + n1
        title = Mid(fieldCode, n1, n2 - n1)
        count = count + 1
        ReDim Preserve titles(0 To count)
        titles(count) = title
        fieldCode = Mid(fieldCode, n2 + 1, Len(fieldCode) - n2 - 1)
    Loop
    
    GetTitles = titles
End Function

Private Sub InsertCrossRef(ByVal RefText As String, ByVal field As field, ByVal title As String, Optional ByVal Year As String = "")
    Dim illegalChars As String
    Dim i As Integer
    Dim result As String
    Dim titleAnchor As String
    
    result = title
    illegalChars = " -—‐!@#$%^&*()+=[]{}|;:',.<>?/`~""\，。；：？！（）【】"
    For i = 1 To Len(illegalChars)
        result = Replace(result, Mid(illegalChars, i, 1), "_")
    Next i
    titleAnchor = Left(result, 35) & "_" & SimpleHash(result)
    
    If Not ActiveDocument.Bookmarks.Exists(titleAnchor) Then
        Selection.GoTo What:=wdGoToBookmark, Name:="Zotero_Bibliography"
        Selection.Find.ClearFormatting
        With Selection.Find
            .Text = Left(title, 255)
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindAsk
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute
        Selection.Paragraphs(1).Range.Select
    
        With ActiveDocument.Bookmarks
            .Add Range:=Selection.Range, Name:=titleAnchor
            .DefaultSorting = wdSortByName
            .ShowHidden = True
        End With
    End If
    
    Dim findRange As Range
    Set findRange = field.result.Duplicate
    findRange.Collapse wdCollapseStart
    With findRange.Find
        .ClearFormatting
        .Text = RefText
        .Forward = True
        .Wrap = wdFindStop
        If .Execute Then
            If Year = "" Then
                If findRange.Hyperlinks.count = 0 Then
                    ActiveDocument.Hyperlinks.Add Anchor:=findRange, Address:="", SubAddress:=titleAnchor, TextToDisplay:=RefText
                End If
            Else
                With findRange.Find
                    .ClearFormatting
                    .Text = Year
                    .Forward = True
                    .Wrap = wdFindStop
                    If .Execute Then
                        If findRange.Hyperlinks.count = 0 Then
                            ActiveDocument.Hyperlinks.Add Anchor:=findRange, Address:="", SubAddress:=titleAnchor, TextToDisplay:=Year
                        End If
                    End If
                End With
            End If
        End If
    End With
    With findRange.Font
        .Underline = wdUnderlineNone
        .Color = wdColorAutomatic
    End With
End Sub

Private Function SimpleHash$(ByVal s$)
    Dim i&, h&
    For i = 1 To Len(s)
        h = h + Asc(Mid(s, i, 1)) * i
    Next i
    h = h Mod 10000
    If h < 0 Then h = h + 10000
    SimpleHash = Format$(h, "0000")
End Function


