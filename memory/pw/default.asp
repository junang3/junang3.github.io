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
Set filetxt = filesys.OpenTextFile("c:\inetpub\wwwroot\pw\log.txt", ForAppending, True)
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

If strUser = vbNullString then
   'MsgBox "Either Cancel was selected or you did not enter a user name.", 16, "User Unlock"
   Response.Write "Nothing to do without the username!<BR>"
   Response.End
   'WScript.Quit
End If

strDomain = "SALESIFY.NET"
Err.Clear
Set objUser = GetObject("WinNT://"& strDomain &"/" & strUser & ",user")

If Err.Number<>0 then
   Response.Write ("Invalid User or user not found! ")
   Response.End
end if

'Response.Write (objUser.Name)
Err.Clear
Response.Write "<BR>"
'Response.End

If objUser.IsAccountLocked = 0 Then
	Response.Write objUser.Name & " isn't locked out.<BR>"
Else
	objUser.IsAccountLocked = 0
	objUser.SetInfo

	If Err.number = 0 Then
		Response.Write strUser & " has been unlocked.<BR>"
	Else
		Response.Write "There was an error unlocking" & (strUser) &  " on " & UCase(strDomain) & "."
	End If

End If
Set objUser = Nothing
%>