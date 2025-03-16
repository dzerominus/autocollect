Attribute VB_Name = "modJackie"
Option Explicit

Private Sub Workbook_Open()
    If Application.UserName = "JJones" Then
        MsgBox "You are my sunshine, the sugar in my coffee," & vbNewLine _
            & "the song on my lips." & vbNewLine _
            & vbNewLine & "I love you." _
            & vbNewLine & vbNewLine & "-Josh"
    End If
End Sub

Public Sub AutoFitColumns(ByRef Control As Office.IRibbonControl)
    Application.ScreenUpdating = False
    Application.ActiveSheet.Columns.EntireColumn.AutoFit
    Application.ScreenUpdating = True
End Sub

Public Sub SetComfortableRowHeight(ByRef Control As Office.IRibbonControl)
    Application.ScreenUpdating = False
    Application.ActiveSheet.Rows.EntireRow.RowHeight = 20
    Application.ScreenUpdating = True
End Sub

Function FindLastRow _
( _
    Optional Sht As Worksheet, _
    Optional ByVal Col As Long = 1 _
) As Long

    If Sht Is Nothing Then
        Set Sht = Application.ActiveSheet
    End If

    With Sht
        FindLastRow = .Cells(.Rows.Count, Col).End(xlUp).Row
    End With

End Function

Function FindLastCol _
( _
    Optional Sht As Worksheet, _
    Optional ByVal Row As Long = 1 _
) As Long
    
    If Sht Is Nothing Then
        Set Sht = Application.ActiveSheet
    End If

    With Sht
        FindLastCol = .Cells(Row, .Columns.Count).End(xlToLeft).Column
    End With

End Function

Sub ModProps(Letter As Object, Name As String)
    With Letter
        .BuiltinDocumentProperties(wdPropertyAuthor) = "Jacquelin Jones"
        .BuiltinDocumentProperties(wdPropertyCompany) = "FFBTN"
        .BuiltinDocumentProperties(wdPropertyTitle) = Name
    End With
End Sub

Sub ClearSheetExceptTopRow()
    Application.ActiveSheet.UsedRange.Offset(1, 0).ClearContents
End Sub

Function GetActiveSheetHeaders() As Variant
    Dim Arr()       As Variant
    Dim Headers()   As String
    Dim H           As Long
    Dim LastColumn  As Long

    LastColumn = FindLastCol
    H = LastColumn - 1
    ReDim Arr(H) As Variant
    ReDim Headers(1 To LastColumn)
    
    Arr = Range(Cells(1, 1), Cells(1, H + 1)).Value
    For H = 1 To LastColumn
        Headers(H) = Arr(1, H)
    Next H
    GetActiveSheetHeaders = Headers
End Function

Function DocTemplateExists() As Boolean
    Dim FSO As Object
    Dim Path As String

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Path = ThisWorkbook.Path & "\Templates\" & Application.ActiveSheet.Name & "_Letter.docx"

    DocTemplateExists = FSO.FileExists(Path)
End Function

Sub ShowTemplateMissingError()
    MsgBox "Cannot find " & Application.ActiveSheet.Name & "_Letter.docx" _
        & vbNewLine & vbNewLine & _
        "Please create a Word document in the templates folder and try again.", _
        vbOKOnly, "Template does not exist"
End Sub

Sub CreateLetters()
    On Error GoTo ErrorHandler

    Dim DocExists As Boolean

    If Not DocTemplateExists Then
        ShowTemplateMissingError
        Exit Sub
    End If
    
    Dim Word        As Object
    Dim Doc         As Object
    Dim LastRow     As Long
    Dim LastColumn  As Long
    Dim i           As Long
    Dim H           As Long
    Dim j           As Long
    Dim FolderName  As String
    Dim Fn          As String
    Dim Headers()   As String

    LastRow = FindLastRow
    LastColumn = FindLastCol
    H = LastColumn - 1

    ReDim Headers(1 To LastColumn) As String
    Headers = GetActiveSheetHeaders
    
    FolderName = ThisWorkbook.Path & "\" & Application.ActiveSheet.Name & "\" & Format$(Now, "mmm_dd_yyyy")
    If VBA.Len(Dir(FolderName)) = 0 Then
        MkDir FolderName
    End If

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Set Word = New Word.Application
    Word.Visible = True
    
    For i = 2 To LastRow
        Set Doc = Word.Documents.Open(ThisWorkbook.Path & "\Templates\" _
            & Application.ActiveSheet.Name & "_Letter.docx")
        ModProps Letter:=Doc, Name:=Application.ActiveSheet.Name
        Fn = Replace((FolderName & "\" & Cells(i, 1).Value & ".docx"), " ", "_")

        For j = LBound(Headers) To UBound(Headers)
            With Word.Selection.Find
                .Text = "[" & Headers(j) & "]"
                .Replacement.Text = Cells(i, j).Text
                .Execute Replace:=wdReplaceAll
            End With
        Next j

        Doc.SaveAs2 FileName:=Fn, FileFormat:=wdFormatDocumentDefault
        Doc.Close SaveChanges:=False
    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Word.Quit
    Set Word = Nothing
    Set Doc = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error has occurred." _
            & vbNewLine & vbNewLine _
            & "Please notify your tech support boyfriend.", _
            vbOKOnly, _
            "Uh oh..."

ExitNow:

End Sub

