<!-- #include file="common.asp" -->

<%

on error resume next

Dim obj

strUser = request("user")
strPass1 = request("pass1")
strPass2 = request("pass2")

if strPass1 <> strPass2 then response.redirect "userpassword.asp?user=" & strUser

response.write "User: " & strUser & "<P>" & vbCrLf

Set obj = GetObject("LDAP://" & ldapEncode(strUser))
strName = mid(obj.name, instr(obj.name, "=") + 1)

obj.SetPassword strPass1
obj.pwdLastSet = 0
obj.SetInfo

Set obj = GetObject("WinNT://" + strUser)
obj.IsAccountLocked = False
obj.SetInfo

select case err.number
	case -2147022651
		response.write "<B>Error: password does not meet complexity requirements</B>"
	case 70
		response.write "<B>Error: Permission denied</B>"
	case else
		response.redirect "userdetails.asp?user=" & strUser
end select

Set obj = Nothing


%>

</TABLE>
</BODY>