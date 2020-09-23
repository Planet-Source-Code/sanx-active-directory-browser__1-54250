<!-- #include file="common.asp" -->

<%

Dim obj

strGroup = request("group")

aryContainer = split(strGroup, ",")
Set obj = GetObject("LDAP://" & ldapEncode(strGroup))

strname = mid(obj.name, instr(obj.name, "=") + 1)
strsamAccountName = obj.samAccountName
strdescription = obj.description
strmail = obj.mail

Set objHash = CreateObject("Scripting.Dictionary")
objHash.Add "ADS_GROUP_TYPE_GLOBAL_GROUP", &h2
objHash.Add "ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP", &h4
objHash.Add "ADS_GROUP_TYPE_UNIVERSAL_GROUP", &h8
objHash.Add "ADS_GROUP_TYPE_SECURITY_ENABLED", &h80000000
intgroupType = obj.groupType
If intgroupType AND objHash.Item("ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP") Then
  strGroupScope = "Domain Local Group"
ElseIf intGroupType AND objHash.Item("ADS_GROUP_TYPE_GLOBAL_GROUP") Then
  strGroupScope = "Global Group"
ElseIf intGroupType AND objHash.Item("ADS_GROUP_TYPE_UNIVERSAL_GROUP") Then
  strGroupScope = "Universal Group"
End If
If intgroupType AND objHash.Item("ADS_GROUP_TYPE_SECURITY_ENABLED") Then
  strGroupType = "Security"
Else
  strGroupType = "Distribution"
End If

strmember = join(resolveusers(obj.member), "<BR>")
strmemberOf = join(resolvegroups(obj.memberOf), "<BR>")

%>
<TABLE WIDTH="800px" STYLE="border: 1px solid gray;" CELLSPACING=0 CELLPADDING=0 BORDER=0>
	<TR>
		<TD WIDTH="200px" ALIGN="center" STYLE="background: #74AAD7;" VALIGN="top"><IMG SRC="iconGroup.gif" ALT="" WIDTH="128" HEIGHT="128" BORDER="0"></TD>
		<TD WIDTH="600px" ROWSPAN=4 VALIGN="top">
			<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 STYLE="width: 100%" ROWS=2 COLS=1>
				<TR>
					<TD STYLE="background: #74AAD7; padding: 10px; width: 100%" VALIGN="top">
						<SPAN STYLE="font-size: 14pt;"><%= strsamAccountName %></SPAN><BR>
						<SPAN STYLE="font-size: 10pt;">Group</SPAN>
					</TD>
				</TR>
				<TR>
					<TD STYLE="padding-top: 5px; padding-bottom: 5px;" VALIGN="top">
						<TABLE WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0><%

doRow "Object Name", strName
doRow "Account Name", strsamAccountName
doRow "Group Scope", strGroupScope
doRow "Group Type", strGroupType
doRow "Description", strdescription
doRow "Mail Address", strmail
doRow "Members", strmember
doRow "Group Membership", strmemberOf

Set obj = Nothing

%>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		<TR>
		<TD WIDTH="200px" STYLE="padding: 10px; background: #74AAD7;" VALIGN="top">
			<DIV STYLE="font-size: 10pt; font-weight: bold; padding-bottom: 5px;">Group Tasks</DIV>
			<A HREF="groupGroups.asp?group=<%=strGroup%>">Group membership</A><P>
			<A HREF="deleteObject.asp?obj=<%=strGroup%>">Delete group</A>
		</TD>
	</TR>
	<TR>
		<TD WIDTH="200px" STYLE="padding: 10px; background: #74AAD7;" VALIGN="top">
			<DIV STYLE="font-size: 10pt; font-weight: bold; padding-bottom: 5px;">Parent Containers</DIV>
			<A HREF="index.asp">Domain Root</A><BR>
			<%
			intRootComponents = ubound(split(strDomainRoot, ",")) + 1
			intContainerComponents = ubound(aryContainer)
			intIndent = 3
			for count = intRootComponents to (intContainerComponents - 1)
				aryContainerName = split(aryContainer(intContainerComponents - count), "=")
				strContainerURL = ""
				for count2 = (intContainerComponents - count) to intContainerComponents
					strContainerURL = strContainerURL & aryContainer(count2)
					if count2 < intContainerComponents then strContainerURL = strContainerURL & ","
				next
				for count2 = 1 to intIndent
					response.write "&nbsp;"
				next
				response.write "-&nbsp;<A HREF=""index.asp?container=" & strContainerURL & _
							   """>" & aryContainerName(1) & "</A><BR>" & vbCrLf
				intIndent = intIndent + 3
			next
			%>
		</TD>
	</TR>
	<TR>
		<TD WIDTH="200px" STYLE="padding: 10px; background: #74AAD7;" VALIGN="top">
		<%WriteSearch%>
		</TD>
	</TR>
</TABLE>

</BODY>