<!-- #include file="common.asp" -->

<%
'on error resume next

dim strObj, strParent, strClass, strName

strCN = request("obj")
aryPath = split(strCN, ",")
strConfirm = request("confirm")

strObj = aryPath(lbound(aryPath))
for count = (lbound(aryPath)+1) to ubound(aryPath)
	strParent = strParent & aryPath(count)
	If count < ubound(aryPath) then strParent = strParent & ","
next

strName = aryPath(lbound(aryPath))

if strConfirm = "yes" then

	Set obj = GetObject("LDAP://" & ldapEncode(strCN))
	strClass = obj.Class
	set obj = nothing
	
	Set obj = GetObject("LDAP://" & ldapEncode(strParent))
	obj.Delete strClass, strName
	
	response.write strParent & "<P>" & strName & "<P>" & strClass
	
	select case err.number
		case 70
			response.write "<B>Error: Permission denied</B>"
		case else
			response.redirect "index.asp?container=" & strParent
	end select
	
	Set obj = Nothing

else
	%>
<TABLE WIDTH="800px" STYLE="border: 1px solid gray;" CELLSPACING=0 CELLPADDING=0 BORDER=0>
	<TR>
		<TD WIDTH="200px" ALIGN="center" STYLE="background: #74AAD7;" VALIGN="top"><IMG SRC="iconDelete.gif" ALT="" WIDTH="128" HEIGHT="128" BORDER="0"></TD>
		<TD WIDTH="600px" ROWSPAN=3 VALIGN="top">
			<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 STYLE="width: 100%" COLS=1 ROWS=2>
				<TR>
					<TD STYLE="background: #74AAD7; padding: 10px; width: 100%" VALIGN="top">
						<SPAN STYLE="font-size: 14pt;"><%= mid(strName, instr(strName, "=")+1 )%></SPAN><BR>
						<SPAN STYLE="font-size: 10pt;">Delete Confirmation</SPAN>
					</TD>
				</TR>
				<TR>
					<TD STYLE="padding: 5px;" VALIGN="top">
							Are you sure you wish to delete this object?<P>
							<FORM METHOD="post" ACTION="deleteObject.asp">
							<INPUT TYPE="hidden" NAME="obj" VALUE="<%=strCN%>">
							<INPUT TYPE="hidden" NAME="confirm" VALUE="yes">
							<INPUT TYPE="submit" NAME="submit" VALUE="Delete it!">&nbsp;&nbsp;<INPUT TYPE="button" NAME="cancel" VALUE="Cancel" ONCLICK="document.location= '<%="index.asp?container=" & MakeURL(strParent)%>';"></FORM>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH="200px" STYLE="padding: 10px; background: #74AAD7;" VALIGN="top">
			&nbsp;
		</TD>
	</TR>
</TABLE>

<%
end if
%>
</BODY>
