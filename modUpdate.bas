Attribute VB_Name = "modUpdate"
Option Explicit

Function GetLatestReleaseTag() As String
    Dim Json        As Object
    Dim Req         As Object
    Dim html_url()  As String
    Dim Res         As String
    
    Set Req = CreateObject("WinHttp.WinHttpRequest.5.1")
    With Req
        .Open "GET", ThisWorkbook.CustomDocumentProperties("Latest_Tag_URL"), False
        .Send
        Res = .ResponseText
    End With
    Set Req = Nothing

    Set Json = JsonConverter.ParseJson(Res)
    html_url = VBA.Split(Json("html_url"), "/")
    GetLatestReleaseTag = html_url(UBound(html_url))
End Function

Function GetLatestReleaseDownloadURL() As String
    Dim URL As String
    Dim Tag As String

    Tag = GetLatestReleaseTag
    URL = ThisWorkbook.CustomDocumentProperties("Update_URL") & "_" & Tag & ".bas"

    GetLatestReleaseDownloadURL = URL
End Function

Function GetCurrentModuleVersion() As String
    GetCurrentModuleVersion = ThisWorkbook.CustomDocumentProperties("Current_Version")
End Function

Function DownloadLatestRelease() As Boolean
    On Error GoTo ErrorHandler

    Const ForWrite  As Integer = 2
    Const CanCreate As Boolean = VBA.vbTrue
    Const Tristate  As Integer = -2

    Dim URL         As String
    Dim FilePath    As String
    Dim Res         As String
    Dim Req         As Object
    Dim FSO         As Object
    Dim File        As Object

    URL = GetLatestReleaseDownloadURL
    Set Req = CreateObject("WinHttp.WinHttpRequest.5.1")

    With Req
        .Open "GET", URL, False
        .Send

        If .Status = 200 Then
            Set FSO = CreateObject("Scripting.FileSystemObject")
            FilePath = ThisWorkbook.Path & "\modJackie.bas"
            FSO.CreateTextFile FilePath
            Set File = FSO.OpenTextFile(FilePath, ForWrite, CanCreate, Tristate)
            File.Write .ResponseText
            File.Close
        End If
    End With

    Set Req = Nothing
    DownloadLatestRelease = True
    Exit Function

ErrorHandler:
    Debug.Print "An error occured while downloading update."
    DownloadLatestRelease = False

End Function

Function ReplaceModule() As Boolean
    On Error GoTo ErrorHandler

    Dim LatestTag   As String
    Dim Project     As VBIDE.VBProject
    Dim Components  As VBIDE.VBComponents
    Dim CurrentMod  As VBIDE.VBComponent
    Dim LatestMod   As VBIDE.VBComponent
    Dim CodeModule  As VBIDE.CodeModule
    
    LatestTag = GetLatestReleaseTag
    Set Project = ThisWorkbook.VBProject
    Set Components = Project.VBComponents
    Set CurrentMod = Components("modJackie")
    Components.Remove CurrentMod
    Set LatestMod = Components.Import(ThisWorkbook.Path & "\modJackie.bas")
    ThisWorkbook.CustomDocumentProperties("Current_Version") = LatestTag

    ReplaceModule = True
    Exit Function
   
ErrorHandler:
    Debug.Print "An error occured while applying the update."
    ReplaceModule = False

End Function

Function IsModuleUpdateAvailable() As Boolean
    Dim IsAvailable As Integer
    IsAvailable = VBA.StrComp(GetCurrentModuleVersion, GetLatestReleaseTag, vbTextCompare)

    If IsAvailable = 0 Then
        IsModuleUpdateAvailable = False
    Else
        IsModuleUpdateAvailable = True
    End If
End Function

Public Sub IsUpdateButtonEnabled(Control As IRibbonControl, ByRef ReturnValue)
    Dim IsAvailable As Boolean
    IsAvailable = IsModuleUpdateAvailable
    If IsAvailable = True Then
        Debug.Print "Module update is available.  Enabling update button."
    End If
    ReturnValue = IsAvailable
End Sub

Public Sub UpdateToLatestRelease(Control As IRibbonControl)
    Dim IsDownloaded As Boolean
    Dim IsReplaced   As Boolean

    IsDownloaded = DownloadLatestRelease
    If IsDownloaded = True Then
        IsReplaced = ReplaceModule

        If IsReplaced = True Then
            Application.StatusBar = "Now using version " & GetCurrentModuleVersion & "!"
            MsgBox "Latest release updated successfully." _
                & vbNewLine & vbNewLine & "AutoCollect has been updated.", _
                vbOKOnly + vbInformation

        Else
            MsgBox "Latest release downloaded successfully but could not be applied.", vbOKOnly + vbExclamation
        End If

    Else
        MsgBox "An error occured while downloading update.", vbOKOnly + vbCritical
    End If

End Sub
