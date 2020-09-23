<!-- #include file="common.asp" -->

<%

Dim dom
Dim ou
Dim cmp
dim strCell(4)

if len(request("container")) < 1 then
	strContainer = strDomainRoot
else
	strContainer = request("container")
end if

aryContainer = split(strContainer, ",")
strTemp = split(aryContainer(0), "=")
strOUName = strTemp(1)
%>

<TABLE WIDTH="800px" STYLE="border: 1px solid gray;" CELLSPACING=0 CELLPADDING=0 BORDER=0>
	<TR>
		<TD WIDTH="200px" ALIGN="center" STYLE="background: #74AAD7;" VALIGN="top"><IMG SRC="iconOU.gif"></TD>
		<TD WIDTH="600px" ROWSPAN=4 VALIGN="top">
			<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH="100%" STYLE="width: 100%" COLS=1 ROWS=2>
				<TR>
					<TD WIDTH="100%" STYLE="background: #74AAD7; padding: 10px; width: 100%" VALIGN="top">
						<SPAN STYLE="font-size: 14pt;"><%= strOUName %></SPAN><BR>
						<SPAN STYLE="font-size: 10pt;">Organisational Unit</SPAN>
					</TD>
				</TR>
				<TR>
					<TD STYLE="padding-top: 5px; padding-bottom: 5px;" VALIGN="top">
						<TABLE WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0>
				<%
				
Set dom = GetObject("LDAP://" & ldapEncode(strContainer))
For Each obj In dom
	strName = mid(obj.name, instr(obj.name, "=") + 1)
	strClass = obj.class
	strDescription = trimDesc(obj.Description)

	response.write "<TR>" & vbcrlf
	select case strClass
		case "organizationalUnit", "container", "builtinDomain", "lostAndFound"
			strCell(0) = "<A HREF=""index.asp?container=" & MakeURL(obj.name) & "," & strContainer & """>" & strName  & "</A>"
			strCell(1) = ResolveObjectName(strClass)
			strCell(2) = strDescription
		case else
			select case strClass
				case "user"
					strCell(0) = "<A HREF=""userdetails.asp?user=" & MakeURL(obj.name) & "," & strContainer & """>" & strName  & "</A>"
					strCell(1) = ResolveObjectName(strClass)
					strCell(2) = strDescription
				case "computer"
					strCell(0) = "<A HREF=""computerdetails.asp?computer=" & MakeURL(obj.name) & "," & strContainer & """>" & strName  & "</A>"
					strCell(1) = ResolveObjectName(strClass)
					strCell(2) = strDescription
				case "group"
					strCell(0) = "<A HREF=""groupdetails.asp?group=" & MakeURL(obj.name) & "," & strContainer & """>" & strName  & "</A>"
					strCell(1) = ResolveObjectName(strClass)
					strCell(2) = strDescription
				case "contact"
					strCell(0) = "<A HREF=""contactdetails.asp?contact=" & MakeURL(obj.name) & "," & strContainer & """>" & strName  & "</A>"
					strCell(1) = ResolveObjectName(strClass)
					strCell(2) = strDescription
				case "printQueue"
					aryTemp = split(strName, "-")
					strTemp = ""
					for count = (lbound(aryTemp) + 1) to ubound(aryTemp)
						strTemp = strTemp & aryTemp(count)
						if count < ubound(aryTemp) then strTemp = strTemp & "-"
					next
					strCell(0) = "<A HREF=""printerdetails.asp?printer=" & MakeURL(obj.Name & "," & strContainer) & """>" & strTemp & " on " & aryTemp(0) & "</A>"
					strCell(1) = ResolveObjectName(strClass)
					strCell(2) = strDescription
				case else
					strCell(0) = strName
					strCell(1) = ResolveObjectName(strClass)
					strCell(2) = strDescription
			end select
	end select
	select case strClass
		case "builtinDomain","computer","container","dnsNode","group", _
			 "infrastructureUpdate","lostAndFound","organizationalUnit", _
			 "user", "printQueue", "contact"
			response.write "<TD WIDTH=16 VALIGN=""top"" CLASS=""data"">" & _
						    "<IMG SRC=""" & strClass & ".gif"" BORDER=0></TD>"
		case else
			response.write "<TD WIDTH=16 VALIGN=""top"" CLASS=""data"">" & _
						   "<IMG SRC=""default.gif"" BORDER=0></TD>"
	end select
	for count = lbound(strCell) to ubound(strCell)
		if len(strCell(count)) = 0 then strCell(count) = "&nbsp;"
		if count = 0 then
			response.write "<TD CLASS=""dataHead"" VALIGN=""top"">" & strCell(count) & "</TD>" & vbCrLf
		else
			 response.write "<TD CLASS=""data"" VALIGN=""top"">" & strCell(count) & "</TD>" & vbCrLf
		end if
	next
	response.write "</TR>" & vbcrlf
Next

Set dom = Nothing
Set obj = Nothing
Set cmp = Nothing

%>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD VALIGN="top" WIDTH="200px" STYLE="padding: 10px; background: #74AAD7;">
		<DIV STYLE="font-size: 10pt; font-weight: bold; padding-bottom: 5px;">Tasks</DIV>
		<A HREF="newUser.asp?container=<%=MakeURL(strContainer)%>">New User</A>
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
		<TD VALIGN="top" WIDTH="200px" STYLE="padding: 10px; background: #74AAD7;">
		<%WriteSearch%>
		</TD>
	</TR>
</TABLE>

</BODY>