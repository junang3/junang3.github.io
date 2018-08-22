  <%
Public Function H2B(Value )
  ' Return a byte value as a two-digit hex string.
  s = ""
  if (Value < &H10) then
    s = "0"
  end if
  H2b = s & hex(value)
End Function  
  
Private Function encode64(s)
    Dim i, strRet
	For i=1 To Len(s)
		strRet = strRet & h2b(asc(Mid(s,i,1)))
	Next     
	'Remove last space at end.
    encode64 = strRet
End Function
  
Private Function Stream_BinaryToString(Binary)
  Const adTypeText = 2
  Const adTypeBinary = 1
  Dim BinaryStream 'As New Stream
  Set BinaryStream = CreateObject("ADODB.Stream")
  BinaryStream.Type = adTypeBinary
  BinaryStream.Open
  BinaryStream.Write Binary
  BinaryStream.Position = 0
  BinaryStream.Type = adTypeText
  BinaryStream.CharSet = "us-ascii"
  Stream_BinaryToString = BinaryStream.ReadText
  Set BinaryStream = Nothing
End Function
  
'******************************************************************************
'** Function:     Encrypt
'** Version:      1.0
'** Created:      20-1-2009 22:35
'** Author:       Adriaan Westra
'** E-mail:         
'**
'** Purpose / Comments:
'**
'**      Encrypt a string in to an encrypted string
'**
'** Change Log :
'**
'** 20-1-2009 22:36 : Initial Version
'**
'** Arguments :  
'**
'**   strEncrypt   :   string to be encrypted
'**   strKey       :   string used as encryption key
'**   intSeed      :   integer to make the encryption random
'**
'** Returns   :
'**
'**   an encrypted string
'**         
'******************************************************************************
Function Encrypt( strEncrypt, strKey, intSeed)
  Rnd(-1)
  Randomize intSeed
  intRnd =  Int( ( Len(strKey) - 1 + 1 ) * Rnd + 1 )
  
  arrEncrypt = String2Asc(strEncrypt)
  arrKey = String2Asc(strKey)
  
  For intI = 0 to UBound( arrEncrypt ) - 1
      
      intPointer = intI + intRnd
      If intPointer > UBound(arrKey) Then
         intPointer = intPointer -  ((UBound(arrKey) + 1 ) * Int(intPointer / (UBound(arrKey) + 1)))
      End If
      
      intCalc = arrEncrypt(intI) + arrKey(intPointer)
      
      If intCalc > 256 Then
      	intCalc = intCalc - 256 
      End If
      strEncrypted = strEncrypted & Chr(intCalc)
  Next
  encrypt = strEncrypted
End Function
'******************************************************************************
'** Function:     Decrypt
'** Version:      1.0
'** Created:      20-1-2009 22:35
'** Author:       Adriaan Westra
'** E-mail:         
'**
'** Purpose / Comments:
'**
'**      Decrypt an encrypted string
'**
'** Change Log :
'**
'** 20-1-2009 22:36 : Initial Version
'**
'** Arguments :  
'**
'**   strDecrypt   :   string to be Decrypted
'**   strKey       :   string used as encryption key
'**   intSeed      :   integer used to make the encryption random
'**
'** Returns   :
'**
'**   A Decrypted string
'**         
'******************************************************************************
Function Decrypt( strDecrypt, strKey, intSeed)
  Rnd(-1)
  Randomize intSeed
  intRnd =  Int( ( Len(strKey) - 1 + 1 ) * Rnd + 1 )
  
  arrDecrypt = String2Asc(strDecrypt)
  arrKey = String2Asc(strKey)
    
  For intI = 0 to UBound( arrDecrypt ) - 1
      
      intPointer = intI + intRnd
      If intPointer > UBound(arrKey) Then
         intPointer = intPointer -  ((UBound(arrKey) + 1 ) * Int(intPointer / (UBound(arrKey) + 1)))
      End If
      
      intCalc = arrDecrypt(intI) - arrKey(intPointer)
      
      If intCalc < 0 Then
      	intCalc = intCalc + 256 
      End If
      strDecrypted = strDecrypted & Chr(intCalc)
  Next
  Decrypt = strDecrypted
End Function

  
  '---------------
  
  
  dim un
  dim pw
  dim dc
  
  
  'edit this 20180822
  un = "..."
  pw = "..."

  a = encode64(encrypt(pw))
  response.write un + "<BR>"
  response.write pw + "<BR>"
  response.write a + "<BR>"
  response.end

  'un = "ЩЮТeвиФЩиЬТк"
  'pw = "еСЭнСЭЯк`СЯЮ"
  
  dim a
  
  a = encrypt(un,"junang3",6)
  response.write a
  response.write "<BR>"
  
  a = base64encode(a)
  response.write a
  response.write "<BR>"
  
  a = base64decode(a)
  response.write a
  response.write "<BR>"
   
  
  a = decrypt(a,"junang3",6)
  
  response.write a
  
  response.end
  
  'un = "VHonZWM/IlQ/bydh"
  'pw = "WSc/LSc/WWFgJ1l6"
  
'::ЩЮТeвиФЩиЬТк::
'::еСЭнСЭЯк`СЯЮ::
'::VHonZWM/IlQ/bydh::
'::WSc/LSc/WWFgJ1l6::
  
  
  dc = "SALESIFY"
  
  'dim unen
  'dim pwen
 ' 
 ' unen = encrypt(un,"20170908")
 ' pwen = encrypt(un,"20170908")

'  response.write "::" & unen & "::"
'  response.write "<BR>"
'  response.write "::" & pwen & "::"
'  response.write "<BR>"
  
  
'  unen = Base64Encode(unen)
'  pwen = Base64Encode(pwen)
'  '
'  response.write "::" & unen & "::"
'  response.write "<BR>"
'  response.write "::" & pwen & "::"
'  response.write "<BR>"
'  
'  response.end
  
  dim domainController : domainController = "phdc01.salesify.net"
  dim ldapPort : ldapPort = 389
  dim startOu : startOu = "DC=SALESIFY,DC=NET"
  
  

  Function CheckLogin( szUserName, szPassword)
    CheckLogin = False
    szUserName = trim( "" &  szUserName)
    dim oCon : Set oCon = Server.CreateObject("ADODB.Connection")
    oCon.Provider = "ADsDSOObject"
    oCon.Properties("User ID") = szUserName
    oCon.Properties("Password") = szPassword
    oCon.Open "ADProvider"
    dim oCmd : Set oCmd = Server.CreateObject("ADODB.Command")
    Set oCmd.ActiveConnection = oCon

    ' let's look for the mail address of a non exitsting user
    dim szDummyQuery : szDummyQuery = "(&(objectCategory=person)(samaccountname=DeGaullesC))"
    dim szDummyProperties : szDummyProperties = "mail"
    dim cmd : cmd = "<" & "LDAP://" & domainController & ":" & ldapPort & _
                        "/" & startOu & ">;" & szDummyQuery & ";" & szDummyProperties & ";subtree"
    oCmd.CommandText = cmd
    oCmd.Properties("Page Size") = 100
    on error resume next
    dim rs : Set rs = oCmd.Execute
    if err.Number = 0 then
      CheckLogin = true
      call rs.Close()
      set rs = nothing
    end if
    on error goto 0
    set oCmd = nothing
  End Function

  ' perform test
  'dim res : res = CheckLogin( dc & "\" & un, pw)
  'response.write decode(base64decode(un),"20170908")
  response.write decrypt(un,"20170908")
  'dim res : res = CheckLogin( dc & "\" & decode(base64decode(un),"20170908"), decode(base64decode(pw),"20170908"))
  if res then
    Response.Write( "Login ok")
  else
    Response.Write( "Login failed")
  end if
  

 	
  

 
  %>
