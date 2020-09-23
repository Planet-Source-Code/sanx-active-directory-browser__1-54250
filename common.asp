<%

strDomainRoot = "DC=sanx,DC=local"

Const ADS_PROPERTY_CLEAR = 1
Const ADS_PROPERTY_UPDATE = 2
Const ADS_PROPERTY_APPEND = 3
Const ADS_PROPERTY_DELETE = 4

function ResolveObjectName(strClassName)

	select case strClassName
		case "builtinDomain"
			ResolveObjectName = "Domain"
		case "computer"
			ResolveObjectName = "Computer"
		case "container"
			ResolveObjectName = "Container"
		case "dnsNode"
			ResolveObjectName = "DNS Node"
		case "group"
			ResolveObjectName = "Group"
		case "infrastructureUpdate"
			ResolveObjectName = "Infrastructure Update"
		case "lostAndFound"
			ResolveObjectName = "Lost and Found"
		case "organizationalUnit"
			ResolveObjectName = "Organisational Unit"
		case "user"
			ResolveObjectName = "User"
		case "contact"
			ResolveObjectName = "User"
		case "printQueue"
			ResolveObjectName = "Print Queue"
		case "msDS-QuotaContainer"
			ResolveObjectName = "Quota Specification"
		case "domainPolicy"
			ResolveObjectName = "Domain Policy"
		case "dfsConfiguration"
			ResolveObjectName = "DFS Configuration"
		case "rpcContainer"
			ResolveObjectName = "RPC Container"
		case "nTFRSSettings"
			ResolveObjectName = "File Replication Configuration"
		case "rIDManager"
			ResolveObjectName = "RID Manager"
		case "fileLinkTracking"
			ResolveObjectName = "File Tracking Container"
		case "secret"
			ResolveObjectName = "Secret"
		case "groupPolicyContainer"
			ResolveObjectName = "Group Policy Container"
		case "domainDNS"
			ResolveObjectName = "DNS Domain"
		case "classStore"
			ResolveObjectName = "Class Store"
		case "ipsecPolicy"
			ResolveObjectName = "IPSEC Policy"
		case "ipsecISAKMPPolicy"
			ResolveObjectName = "IPSEC ISAKMP Policy"
		case "ipsecFilter"
			ResolveObjectName = "IPSEC Filter"
		case "ipsecNegotiationPolicy"
			ResolveObjectName = "IPSEC Negotiation Policy"
		case "foreignSecurityPrincipal"
			ResolveObjectName = "Foreign Security Principal"
		case "rIDSet"
			ResolveObjectName = "RID Set"
		case "nTFRSReplicaSet"
			ResolveObjectName = "File Replication Set"
		case "nTFRSMember"
			ResolveObjectName = "File Replication Member Object"
		case "nTFRSSubscriptions"
			ResolveObjectName = "File Replication Subscriptions"
		case "nTFRSSubscriber"
			ResolveObjectName = "File Replication Subscriber"
		case else
			ResolveObjectName = strClassName
	end select
	
end function

function doRow(title, data)

if len(data) > 0 then
	response.write "<TR><TD CLASS=""dataHead"" VALIGN=""top"">" & title & ":</TD>" & _
	vbCrLf & "<TD CLASS=""data"" VALIGN=""top"">" & data & "</TD></TR>" & vbCrLf
end if

end function

function trimDesc(description)

	if len(description) > 20 then trimDesc = left(description, 20) & "..."

end function

function ldapEncode(adspath)

	adspath = replace(adspath, "\", "\5C")
	adspath = replace(adspath, vbCr, "\0D")
	adspath = replace(adspath, vbLf, "\0A")
	adspath = replace(adspath, """", "\22")
	adspath = replace(adspath, "#", "\23")
	adspath = replace(adspath, "+", "\2B")
	adspath = replace(adspath, "/", "\2F")
	adspath = replace(adspath, ";", "\3B")
	adspath = replace(adspath, "<", "\3C")
	'adspath = replace(adspath, "=", "\3D")
	adspath = replace(adspath, ">", "\3E")
	ldapEncode = adspath
	
end function

function MakeURL(url)

	MakeURL = Server.URLPathEncode(url)
	MakeURL = replace(MakeURL, "/", "%2F")
	MakeURL = replace(MakeURL, "\", "%5C")

end function

function resolvegroups(data)

	dim grp, grpObj, temp()

	if isArray(data) then
		redim temp(ubound(data))
		for count = lbound(data) to ubound(data)
			Set grp = GetObject("LDAP://" & data(count))
			temp(count) = "<A HREF=""groupdetails.asp?group=" & grp.distinguishedName & """>" & grp.samAccountName & "</A>"
			set grp = nothing
		next
	elseif VarType(data) = 8 then
		redim temp(0)
		Set grp = GetObject("LDAP://" & data)
		temp(0) = "<A HREF=""groupdetails.asp?group=" & grp.distinguishedName & """>" & grp.samAccountName & "</A>"
		set grp = nothing
	end if
	
	resolvegroups = temp

end function

function resolveusers(data)

	dim usr, temp()

	if isArray(data) then
		redim temp(ubound(data))
		for count = lbound(data) to ubound(data)
			Set usr = GetObject("LDAP://" & data(count))
			select case usr.class
				case "user"
					temp(count) = "<A HREF=""userdetails.asp?user=" & data(count) & """>" & mid(usr.name, instr(usr.name, "=") + 1) & "</A>"
				case "group"
					temp(count) = "<A HREF=""groupdetails.asp?group=" & data(count) & """>" & mid(usr.name, instr(usr.name, "=") + 1) & "</A>"
				case "computer"
					temp(count) = "<A HREF=""computerdetails.asp?computer=" & data(count) & """>" & mid(usr.name, instr(usr.name, "=") + 1) & "</A>"
				case else
					temp(count) = mid(usr.name, instr(usr.name, "=") + 1)
			end select
		next
	elseif VarType(data) = 8 then
		redim temp(0)
		Set usr = GetObject("LDAP://" & data)
		select case usr.class
			case "user"
					temp(0) = "<A HREF=""userdetails.asp?user=" & data & """>" & mid(usr.name, instr(usr.name, "=") + 1) & "</A>"
				case "group"
					temp(0) = "<A HREF=""groupdetails.asp?group=" & data & """>" & mid(usr.name, instr(usr.name, "=") + 1) & "</A>"
				case "computer"
					temp(0) = "<A HREF=""computerdetails.asp?computer=" & data & """>" & mid(usr.name, instr(usr.name, "=") + 1) & "</A>"
				case else
					temp(0) = mid(usr.name, instr(usr.name, "=") + 1)
		end select
	end if

	resolveusers = temp

	set usr = nothing

end function

function resolveuser(data)

	dim usr
	
	if len(data) = 0 then resolveuser = "": exit function
	
	set usr = GetObject("LDAP://" & data)
	resolveuser = "<A HREF=""userdetails.asp?user=" & data & """>" & mid(usr.name, instr(usr.name, "=") + 1) & "</A>"

	set usr = nothing

end function

function Bool2String(data)

	if data = true then
		bool2string = "<SPAN STYLE=""font-weight: bold; color: red;"">Yes</SPAN>"
	else
		bool2string = "No"
	end if

end function

function ListGroups()

	Dim Connect
	dim strGroups

    Set Connect = CreateObject("ADODB.Connection")
	Connect.Provider = "ADsDSOObject"
	Connect.Open "DS Query"

    Query = "<LDAP://" & strDomainRoot & ">;(objectCategory=group);distinguishedName,adspath;subtree"
	
    Set RecordSet = Connect.Execute(Query)
	
	strGroups = ""
	If not(RecordSet.EOF And RecordSet.BOF) Then
		While Not RecordSet.EOF
			strGroups = strGroups & RecordSet.Fields("distinguishedName") & "###"
			RecordSet.MoveNext
		Wend
	End If
	strGroups = left(strGroups, len(strGroups) - 3)
	
	ListGroups = split(strGroups, "###")

end function

sub WriteSearch
%>
<FORM METHOD="GET" ACTION="search.asp">
<SPAN STYLE="font-size: 10pt; font-weight: bold;">Search</SPAN><BR>
<INPUT TYPE="text" NAME="searchstr" STYLE="width: 175px;"><BR>
<B>Object Type:</B><BR>
<SELECT NAME="objectType" STYLE="width: 175px;">
	<OPTION SELECTED VALUE="any">Any</OPTION>
	<OPTION VALUE="user">User</OPTION>
	<OPTION VALUE="group">Group</OPTION>
	<OPTION VALUE="computer">Computer</OPTION>
	<OPTION VALUE="printer">Printer</OPTION>
	<OPTION VALUE="contact">Contact</OPTION>
</SELECT><P>
<INPUT TYPE="submit" VALUE="  Search  "></FORM><P>
<A HREF="advSearch.asp">Advanced Search</A><P>
<%
end sub
%>

<HTML>
	<HEAD>
		<TITLE>Aqua Software Active Directory Browser</TITLE>
		<LINK HREF="styles.css" REL="stylesheet" TYPE="text/css">
	</HEAD>
	
<SCRIPT LANGUAGE="JavaScript">

</SCRIPT>

<BODY ALIGN="center">