<!-- #include file="common.asp" -->

<%

Dim obj

strUser = request("user")
strAction = request("action")
strContainer = request("container")

select case strAction
	case "edit"
		Set obj = GetObject("LDAP://" & ldapEncode(strUser))
		updateInfo "sAMAccountName"
	case "new"
		Set objContainer = GetObject("LDAP://" & ldapEncode(strContainer))
		Set obj = objContainer.Create("user", "cn=" & request("sAMAccountName"))
		obj.SetInfo
		updateInfo "sAMAccountName"
		strUser = "CN=" & request("sAMAccountName") & "," & strContainer
end select

updateInfo "givenName"
updateInfo "initials"
updateInfo "sn"
updateInfo "displayName"
updateInfo "description"
updateInfo "physicalDeliveryOfficeName"
updateInfo "mail"
updateInfo "wWWHomePage"
updateInfo "otherTelephone"
updateInfo "streetAddress"
updateInfo "l"
updateInfo "st"
updateInfo "postalCode"
updateInfo "c"
updateInfo "userWorkstations"
updateInfo "profilePath"
updateInfo "scriptPath"
updateInfo "homeDrive"
updateInfo "homePhone"
updateInfo "pager"
updateInfo "mobile"
updateInfo "facsimileTelephoneNumber"
updateInfo "ipPhone"
updateInfo "title"
updateInfo "department"
updateInfo "company"
updateInfo "manager"

obj.SetInfo
Set obj = Nothing

sub updateInfo(field)

	dim strTemp
	
	strTemp = request(field)
	
	if not (isNull(strTemp) or isEmpty(strTemp) or len(strTemp) = 0) then
		obj.Put field, strTemp
	else
		obj.PutEx ADS_PROPERTY_CLEAR, field, vbNullString
		obj.SetInfo
	end if

end sub

response.redirect "userdetails.asp?user=" & MakeURL(strUser)

%>