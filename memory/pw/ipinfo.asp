<%
On Error Resume Next

'Set oUser = GetObject("WinNT://salesify.net/salvador.ang,user")
'Response.Write (oUser.Name)
'set oUser=nothing
'Response.End
dim un
un = Request.QueryString("un")
rst = Request.QueryString("reset")
ip = Request.QueryString("ip")
host = Request.QueryString("host")


dim filesys, filetxt
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set filesys = CreateObject("Scripting.FileSystemObject")
Set filetxt = filesys.OpenTextFile("c:\inetpub\wwwroot\pw\ipinfo.txt", ForAppending, True)
dim xx 
if ip="" then
   xx = now() & "," & Request.ServerVariables("remote_addr") & "," & un
else
   xx = now() & "," & host & "," & ip & "," & un
end if
if (rst<>"") then xx = xx & "," & rst

filetxt.WriteLine(xx)
filetxt.Close 
set filetxt=nothing
set filesys=nothing

'Response.Write un & "<BR>" & rst
'Response.End

'strUser = InputBox("Please enter a user name.","Unlock User")
'strUser = "salvador.ang"
strUser = un

If ip = vbNullString then
   'MsgBox "Either Cancel was selected or you did not enter a user name.", 16, "User Unlock"
   Response.Write "Nothing to do without the IP!<BR>"
   Response.End
   'WScript.Quit
End If

set xShell = Server.CreateObject("WScript.Shell")
fname = day(date) & month(date) & year(date) & hour(Time) & minute(Time) & second(Time)
outFile = "c:\inetpub\www\pw\ipinfo\" & fname & ".txt"
ret = xShell.Run("c:\Windows\System32\wbem\WMIC.exe /NODE: " & ip & " COMPUTERSYSTEM GET USERNAME > " & outFile, 0, true)
set xShell = nothing

'Set oShell = Server.CreateObject("WScript.Shell")
'sExecStr = "c:\windows\cmd.exe /c c:\Windows\System32\wbem\WMIC.exe /NODE: " & ip & " COMPUTERSYSTEM GET USERNAME > " & outFile
'Set oExec = oShell.Exec(sExecStr)
'Do
'      tmpStr = oExec.StdOut.ReadAll()
'Loop While Not oExec.Stdout.atEndOfStream

'RetCode = oExec.stderr.readall
'Response.Write RetCode

'Set oShell = nothing

'Set fso  = CreateObject("Scripting.FileSystemObject")
'Set file = fso.OpenTextFile(outFile, 1)
'outText = file.ReadAll
'Response.Write outText
'file.Close

Response.Write "Hello"



'outFile = "c:\inetpub\www\pw\info\" & fname & ".txt"
'strCmd = "dir >" & outFile
'xShell.Run strCmd, 1, true
'Set xShell = Nothing
'Set Fso = Server.CreateObject("Scripting.FileSystemObject")
'  If Fso.FileExists(outFile) Then
'    Set Fsf = Fso.OpenTextFile(outFile)
'    outText = Fsf.readall
'    Set Fsf = Nothing
'    Fso.DeleteFile(outFile)
'  end if
'Set Fso = Nothing
'Response.Write outText

%>