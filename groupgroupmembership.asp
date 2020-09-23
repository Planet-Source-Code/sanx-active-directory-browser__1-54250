<!-- #include file="common.asp" -->

<%

On Error Resume Next

Dim obj

strGroup = request("group")
lstGroups = split(request("lstGroups"), "###")

Set obj = GetObject("LDAP://" & ldapEncode(strGroup))
if isArray(obj.memberOf) then
	for each grp in obj.memberOf
		Set grpObj = GetObject("LDAP://" & ldapEncode(grp))
		grpObj.PutEx ADS_PROPERTY_DELETE, "member", Array(strGroup)
		grpObj.SetInfo
		set grpObj = nothing
	next
end if

for each grp in lstGroups
	Set grpObj = GetObject("LDAP://" & ldapEncode(grp))
	grpObj.PutEx ADS_PROPERTY_APPEND, "member", Array(strGroup)
	grpObj.SetInfo
	set grpObj = nothing
next

Set obj = Nothing

response.redirect "groupdetails.asp?group=" & MakeURL(strGroup)

%>