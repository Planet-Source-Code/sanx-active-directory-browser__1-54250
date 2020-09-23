<!-- #include file="common.asp" -->

<%

On Error Resume Next

Dim obj

strUser = request("user")
lstGroups = split(request("lstGroups"), "###")

Set obj = GetObject("LDAP://" & ldapEncode(strUser))
if isArray(obj.memberOf) then
	for each grp in obj.memberOf
		Set grpObj = GetObject("LDAP://" & ldapEncode(grp))
		grpObj.PutEx ADS_PROPERTY_DELETE, "member", Array(strUser)
		grpObj.SetInfo
		set grpObj = nothing
	next
end if

for each grp in lstGroups
	Set grpObj = GetObject("LDAP://" & ldapEncode(grp))
	grpObj.PutEx ADS_PROPERTY_APPEND, "member", Array(strUser)
	grpObj.SetInfo
	set grpObj = nothing
next

Set obj = Nothing

response.redirect "userdetails.asp?user=" & MakeURL(strUser)

%>