Attribute VB_Name = "modUpdate"
Option Explicit

Private Function GetLatestReleaseTag() As String
    Dim JSON        As Object
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

    Set JSON = JsonConverter.ParseJson(Res)
    html_url = VBA.Split(JSON("html_url"), "/")
    GetLatestReleaseTag = html_url(UBound(html_url))
End Function

Private Function GetLatestReleaseDownloadURL() As String
    Dim URL As String
    Dim Tag As String

    Tag = GetLatestReleaseTag
    URL = ThisWorkbook.CustomDocumentProperties("Update_URL") & "_" & Tag & ".bas"

    GetLatestReleaseDownloadURL = URL
End Function

Private Function GetCurrentModuleVersion() As String
    GetCurrentModuleVersion = ThisWorkbook.CustomDocumentProperties("Current_Version")
End Function

Private Function DownloadLatestRelease() As Boolean
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

Private Function ReplaceModule() As Boolean
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
    Kill ThisWorkbook.Path & "\modJackie.bas"
    Exit Function
   
ErrorHandler:
    Debug.Print "An error occured while applying the update."
    ReplaceModule = False

End Function

Private Function IsModuleUpdateAvailable() As Boolean
    Dim CurrentV    As Integer
    Dim LatestV     As Integer
    
    CurrentV = CInt(VBA.Join(VBA.Split(modUpdate.GetCurrentModuleVersion, "."), ""))
    LatestV = CInt(VBA.Join(VBA.Split(modUpdate.GetLatestReleaseTag, "."), ""))
    
    IsModuleUpdateAvailable = LatestV > CurrentV
End Function

Public Sub IsUpdateButtonEnabled(Control As IRibbonControl, ByRef ReturnValue)
    ReturnValue = IsModuleUpdateAvailable
End Sub

Private Sub UpdateButton_OnAction(Control As IRibbonControl)
    If Not IsModuleUpdateAvailable Then
        Exit Sub
    End If

    UpdateToLatestRelease
End Sub

Sub Ribbon_OnLoad(RibbonUI As IRibbonUI)
    If IsModuleUpdateAvailable Then
        Application.StatusBar = "A new version is available..."
    Else
        Application.StatusBar = "Version " & GetCurrentModuleVersion
    End If
End Sub

Public Sub UpdateToLatestRelease()
    If DownloadLatestRelease = True Then
        If ReplaceModule = True Then
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
