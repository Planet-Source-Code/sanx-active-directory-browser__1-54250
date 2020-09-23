<!-- #include file="common.asp" -->

<%

Dim obj

strCmp = request("computer")

aryContainer = split(strCmp, ",")
Set obj = GetObject("LDAP://" & ldapEncode(strCmp))

strName = mid(obj.name, instr(obj.name, "=") + 1)
strdnsHostName = obj.dnsHostName
strdescription = obj.description
Set objHash = CreateObject("Scripting.Dictionary")
objHash.Add "ADS_UF_TRUSTED_FOR_DELEGATION", &h80000
objHash.Add "ADS_UF_WORKSTATION_TRUST_ACCOUNT", &h1000
objHash.Add "ADS_UF_SERVER_TRUST_ACCOUNT", &h2000
intuserAccountControl = obj.Get("userAccountControl")
If intuserAccountControl AND objHash.Item("ADS_UF_TRUSTED_FOR_DELEGATION") Then
  strDelegrationTrust = CStr(true)
Else
  strDelegrationTrust = CStr(false)
End If
If intuserAccountControl AND objHash.Item("ADS_UF_SERVER_TRUST_ACCOUNT") Then
  strRole =  "Domain Controller"
Else
  strRole =  "Workstation or Server"
End If
stroperatingSystem = obj.operatingSystem
stroperatingSystemVersion = obj.operatingSystemVersion
stroperatingSystemServicePack = obj.operatingSystemServicePack
strmemberOf = join(resolvegroups(obj.memberOf), "<BR>")
strlocation = obj.location
%>
<TABLE WIDTH="800px" STYLE="border: 1px solid gray;" CELLSPACING=0 CELLPADDING=0 BORDER=0>
	<TR>
		<TD WIDTH="200px" ALIGN="center" STYLE="background: #74AAD7;" VALIGN="top"><IMG SRC="iconComputer.gif" ALT="" WIDTH="128" HEIGHT="128" BORDER="0"></TD>
		<TD WIDTH="600px" ROWSPAN=3 VALIGN="top">
			<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 STYLE="width: 100%" COLS=1 ROWS=2>
				<TR>
					<TD STYLE="background: #74AAD7; padding: 10px; width: 100%" VALIGN="top">
						<SPAN STYLE="font-size: 14pt;"><%= strName %></SPAN><BR>
						<SPAN STYLE="font-size: 10pt;">Computer</SPAN>
					</TD>
				</TR>
				<TR>
					<TD STYLE="padding-top: 5px; padding-bottom: 5px;" VALIGN="top">
						<TABLE WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0><%
doRow "Object Name", strName
doRow "DNS Host Name", strdnshostname
doRow "Description", strdescription
doRow "Trusted for Delegation", strDelegrationTrust
doRow "Domain Role", strrole
doRow "OS", stroperatingSystem
doRow "OS Version", stroperatingSystemVersion
doRow "OS Service Pack", stroperatingSystemServicePack
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
			<DIV STYLE="font-size: 10pt; font-weight: bold; padding-bottom: 5px;">Computer Tasks</DIV>
			<A HREF="index.asp?container=<%=strCmp%>">View as container</A><BR>
			<A HREF="deleteObject.asp?obj=<%=strCmp%>">Delete account</A>
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