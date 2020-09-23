<!-- #include file="common.asp" -->

<%

Dim obj
dim printJobs()

strPrinter = request("printer")
aryContainer = split(strPrinter, ",")

Set obj = GetObject("LDAP://" & ldapEncode(strPrinter))

strTemp = split(obj.Name, "=")
if isArray(strTemp) then strName = strTemp(1)
strdisplayName = obj.displayName
strdescription = obj.description
strlocation = obj.location
strmodel = obj.model
strpath = obj.PrinterPath
strstatus = ResolvePQStatusCode(CInt("&h" & obj.Status))
set objPJ = obj.PrintJobs

intPJCount = 0
for each pj in objPj
	intPJCount = intPJCount + 1
	redim preserve printJobs(intPJCount -1)
	printJobs(intPJCount-1) = "<B>Job</B>: " & pj.Description & "<BR><B>" & "Total Pages</B>: " & _
	pj.TotalPages & "<BR><B>Size</B>: " & (pj.size / 1024) & " kb" & "<BR><B>User</B>: " & _
	"<A HREF=""search.asp?searchstr=" & pj.User & """>" & pj.User  & "</A>"
next

%>
<TABLE WIDTH="800px" STYLE="border: 1px solid gray;" CELLSPACING=0 CELLPADDING=0 BORDER=0>
	<TR>
		<TD WIDTH="200px" ALIGN="center" STYLE="background: #74AAD7;" VALIGN="top"><IMG SRC="iconPrinter.gif" ALT="" WIDTH="128" HEIGHT="128" BORDER="0"></TD>
		<TD WIDTH="600px" ROWSPAN=3 VALIGN="top">
			<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 STYLE="width: 100%" COLS=1 ROWS=2>
				<TR>
					<TD STYLE="background: #74AAD7; padding: 10px; width: 100%" VALIGN="top">
						<SPAN STYLE="font-size: 14pt;"><%= strName %></SPAN><BR>
						<SPAN STYLE="font-size: 10pt;">Printer</SPAN>
					</TD>
				</TR>
				<TR>
					<TD STYLE="padding-top: 5px; padding-bottom: 5px;" VALIGN="top">
						<TABLE WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0><%

doRow "Display Name", strdisplayName
doRow "Description", strdescription
doRow "Location", strlocation
doRow "Model", strmodel
doRow "Share Path", strPath
doRow "Status", strstatus
doRow "Waiting Print Jobs", join(printJobs, "<P>")

Set objPJ = nothing
Set obj = Nothing

%>
						</TABLE>
					</TR>
				</TD>
			</TABLE>
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

<%

function ResolvePJStatusCode(status)

select case status
	Case 1
		ResolvePJStatusCode = "Job paused"
	Case 2
		ResolvePJStatusCode = "Error"
	Case 4
		ResolvePJStatusCode = "Deleting"
	Case 8
		ResolvePJStatusCode = "Spooling"
	Case 16
		ResolvePJStatusCode = "Printing"
	Case 32
		ResolvePJStatusCode = "Printer offline"
	Case 64
		ResolvePJStatusCode = "Out of paper"
	Case 128
		ResolvePJStatusCode = "Printed"
	Case 256
		ResolvePJStatusCode = "Deleted"
	Case else
		ResolvePJStatusCode = CStr(status)
End Select

End Function

function ResolvePQStatusCode(status)

select case status
	Case 0
		ResolvePQStatusCode = "Online"
	Case 1
		ResolvePQStatusCode = "Paused"
End Select

End Function

%>

</TABLE>
</BODY>