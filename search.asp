<!-- #include file="common.asp" -->

<%

Dim dom
Dim ou
Dim cmp
dim strWildcardQuery
dim Query
dim strCell(3)

strSearch = request("searchstr")
strObjectType = request("objectType")
blnAccDisabled = request("accDisabled")

if blnAccDisabled = "true" then blnAccDisabled = true else blnAccDisabled = false

strContainer = strDomainRoot

    Set Connect = CreateObject("ADODB.Connection")
	Connect.Provider = "ADsDSOObject"
	Connect.Open "DS Query"

    Query = "<LDAP://" & strDomainRoot & ">;(&"
	if len(strSearch) > 0 then Query = Query & "(cn=" & strSearch & "*)"
	if blnAccDisabled then
		Query = Query & "(&(userAccountControl:1.2.840.113556.1.4.803:=2)"
	end if
	Select case strObjectType
		case "any"
			Query = Query & "(|(objectCategory=user)(objectCategory=organizationalUnit)" & _
							"(objectCategory=computer)(objectCategory=group)" & _
							"(objectCategory=container)(objectCategory=contact)" & _
							"(objectCategory=builtinDomain))"
		case "user"
			Query = Query & "(objectCategory=user)"
		case "computer"
			Query = Query & "(objectCategory=computer)"
		case "contact"
			Query = Query & "(&(objectCategory=contact)(objectCategory=person))"
		case "group"
			Query = Query & "(objectCategory=group)"
		case "printer"
			Query = Query & "(objectCategory=printQueue)"
	End Select
	if blnAccDisabled then
		Query = Query & ")"
	end if
	Query = Query & ");distinguishedName,adspath;subtree"
	
    Set RecordSet = Connect.Execute(Query)
	%>

<TABLE WIDTH="800px" STYLE="border: 1px solid gray;" CELLSPACING=0 CELLPADDING=0 BORDER=0>
	<TR>
		<TD WIDTH="200px" ALIGN="center" STYLE="background: #74AAD7;" VALIGN="top"><IMG SRC="iconSearch.gif"></TD>
		<TD WIDTH="600px" ROWSPAN=3 VALIGN="top">
			<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH="100%" STYLE="width: 100%" COLS=1 ROWS=2>
				<TR>
					<TD WIDTH="100%" STYLE="background: #74AAD7; padding: 10px; width: 100%" VALIGN="top">
						<SPAN STYLE="font-size: 14pt;">Search Results</SPAN><BR>
						<SPAN STYLE="font-size: 10pt;">&nbsp;</SPAN>
					</TD>
				</TR>
				<TR>
					<TD STYLE="padding-top: 5px; padding-bottom: 5px;" VALIGN="top">
						<TABLE WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0>
				<%

    If not(RecordSet.EOF And RecordSet.BOF) Then
		While Not RecordSet.EOF
            Set obj = GetObject("LDAP://" & ldapEncode(RecordSet.Fields("distinguishedName")))
				aryTemp = split(obj.distinguishedName, ",")
				strObjName = aryTemp(lbound(aryTemp))
				strName = mid(strObjName, instr(strObjName, "=") + 1)
				strClass = obj.class
				strDescription = trimDesc(obj.Description)

				response.write "<TR>" & vbcrlf
				select case strClass
					case "organizationalUnit", "container", "builtinDomain", "lostAndFound"
						strCell(0) = "<A HREF=""index.asp?container=" & MakeURL(obj.distinguishedname) & """>" & strName  & "</A>"
						strCell(1) = ResolveObjectName(strClass)
						strCell(2) = strDescription
					case else
						select case strClass
							case "user"
								strCell(0) = "<A HREF=""userdetails.asp?user=" & MakeURL(obj.distinguishedname) & """>" & strName  & "</A>"
								strCell(1) = ResolveObjectName(strClass)
								strCell(2) = strDescription
							case "computer"
								strCell(0) = "<A HREF=""computerdetails.asp?computer=" & MakeURL(obj.distinguishedname) & """>" & strName  & "</A>"
								strCell(1) = ResolveObjectName(strClass)
								strCell(2) = strDescription
							case "group"
								strCell(0) = "<A HREF=""groupdetails.asp?group=" & MakeURL(obj.distinguishedname) & """>" & strName  & "</A>"
								strCell(1) = ResolveObjectName(strClass)
								strCell(2) = strDescription
							case "printQueue"
								aryTemp = split(strName, "-")
								strTemp = ""
								for count = (lbound(aryTemp) + 1) to ubound(aryTemp)
									strTemp = strTemp & aryTemp(count)
									if count < ubound(aryTemp) then strTemp = strTemp & "-"
								next
								strCell(0) = "<A HREF=""printerdetails.asp?printer=" & MakeURL(obj.distinguishedName) & """>" & strTemp & " on " & aryTemp(0) & "</A>"
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
						response.write "<TD WIDTH=16 VALIGN=""top"" STYLE=""padding-left: 5px;"">" & _
									    "<IMG SRC=""" & strClass & ".gif"" BORDER=0></TD>"
					case else
						response.write "<TD WIDTH=16 VALIGN=""top"" STYLE=""padding-left: 5px;"">" & _
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

            RecordSet.MoveNext
        WEnd
    End If

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
		<%WriteSearch%>
		</TD>
	</TR>
</TABLE>
</BODY>