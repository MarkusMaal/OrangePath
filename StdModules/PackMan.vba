'
' Package Manager for OrangePath/OS
'

' If you want to use another package repository, change this value here
Const Repository As String = "http://nossl.markusmaal.ee/op_packages/"

Function FindPackage(ByVal Keywords As String)
    Success = Fetch("?cmd=list&keyword=" & Keywords, "packages.list")
    If Success Then
        Const ForReading = 1, ForWriting = 2, ForAppending = 8
        Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
        Set fS = CreateObject("Scripting.FileSystemObject")
        Set F = fS.OpenTextFile(Environ("TEMP") & "\packages.list", ForReading, True, TristateFalse)
        fContent = F.ReadAll
        F.Close
        FindPackage = fContent
    Else
        FindPackage = "Couldn't fetch package list"
    End If
End Function

Function DownloadPackage(ByVal PackageName As String)
    Success = Fetch("?cmd=install&pkg=" & PackageName, "App" & PackageName & ".pptm")
    If Success = True Then
        Slide12.Shapes("FirmwareSource").TextFrame.TextRange.Text = Environ("TEMP") & "\App" & PackageName & ".pptm"
        DownloadPackage = "Package downloaded successfully!"
    Else
        DownloadPackage = "Package not found - do you have a working internet connection?"
    End If
End Function


' source: https://stackoverflow.com/questions/17877389/how-do-i-download-a-file-using-VBA-without-internet-explorer
Function Fetch(ByVal URL As String, ByVal Filename As String)
    Dim myURL As String
    myURL = Repository & URL
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False, "username", "password"
    WinHttpReq.send
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile Environ("TEMP") & "\" & Filename, 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
        Fetch = True
    Else
        Fetch = False
    End If
End Function
