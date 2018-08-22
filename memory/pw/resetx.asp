<%
On Error Resume Next

Set oUser = GetObject("WinNT://salesify.net/salvador.ang,user")
Response.Write (oUser.Name)
set oUser=nothing
Response.End



dim un
un = Request.QueryString("un")
rst = Request.QueryString("reset")
ip = Request.QueryString("ip")
host = Request.QueryString("host")


dim filesys, filetxt
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set filesys = CreateObject("Scripting.FileSystemObject")
Set filetxt = filesys.OpenTextFile("c:\inetpub\wwwroot\pw\reset.log", ForAppending, True)
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

'-----------------

'objUser.IsAccountLocked = 0
'Set objIADS = GetObject("WinNT:").OpenDSObject("WinNT://domain", "Administrator", sDomainPassword, ADS_SECURE_AUTHENTICATION)
'Set objIADSUser = objIADS.GetObject("user", sUserID)
'objIADSUser.ChangePassword sOldPassword, sNewPassword
''Alternatively: objIADSUser.SetPassword sNewPassword
'objIADSUser.SetInfo

'-----------------


If objUser.  = 0 Then
	Response.Write objUser.Name & " isn't locked out.<BR>I can't reset what isn't locked out!!<BR>Maybe the password has expired?<BR>"
Else
	'objUser.IsAccountLocked = 0
	objIADSUser.SetPassword "dell@123"
	objUser.SetInfo

	If Err.number = 0 Then
		Response.Write strUser & " password has been reset to dell@123.<BR>"
	Else
		Response.Write "There was an error resetting the password of " & (strUser) &  " on " & UCase(strDomain) & "."
	End If

End If
Set objUser = Nothing
%>