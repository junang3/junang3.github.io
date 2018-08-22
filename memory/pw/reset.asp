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

'if IsUserGroupMember(GetDistinguishedName(strUser), "Domain Admins") then
'  response.write err.description
'  respons.write "<BR><BR>" & vbCrLF
'  response.write "You can't reset an administrator user!"
'  response.end
'end if 
admins = "|aakash.sharma|administrator|ajay.lichade|arnold.dsilva|ashish.barmase|bhagwan.solanki|cloudberry|admin|kapil.jadhav|mahesh.nirmal|nilesh.salvi|roby.carriedo|sachin.garg|salvador.ang|sanjay.korana|sophos|utm|vrunesh.botre|"

if instr(1, admins, "|" & lcase(strUser) & "|") then
response.write "You can't reset an administrator!"
response.end
end if

authip = "|192.168.0.233|192.168.1.33|"

if instr(1, authip, "|" & ip & "|") = false then
  response.write "You (" & ip & ") are not AUTHORIZED to reset any password<BR>"
  response.write "This has been logged..<BR>"
  response.end
end if



If objUser.IsAccountLocked = 0 Then
	'Response.Write objUser.Name & " isn't locked out.<BR>"
	Response.Write objUser.Name & " isn't locked out.<BR>I can't reset what isn't locked out!!<BR>Maybe the password has expired?<BR>"
Else
	objUser.IsAccountLocked = 0
	objUser.SetPassword "dell@123"
	objUser.SetInfo

	If Err.number = 0 Then
		Response.Write strUser & " password has been reset to dell@123<BR>Password has to be changed right away!"
	Else
		Response.Write "There was an error resetting the password for " & (strUser) &  " on " & UCase(strDomain) & "."
	End If

End If
Set objUser = Nothing
%>