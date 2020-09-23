<!-- #include file="common.asp" -->

<SCRIPT LANGUAGE="JavaScript">

function addGroup() {
	var ag = document.forms[0].elements["ag"];
	var cg = document.forms[0].elements["cg"];
	var newOpt = new Option;
	
	if(!(ag.selectedIndex == -1)) {
		newOpt.text = ag.options[ag.selectedIndex].text;
		newOpt.value = ag.options[ag.selectedIndex].value;
		
		ag.options[ag.selectedIndex] = null;
		cg.options[cg.options.length] = newOpt;
	}
}

function removeGroup() {
	var ag = document.forms[0].elements["ag"];
	var cg = document.forms[0].elements["cg"];
	var newOpt = new Option;
	
	if(!(cg.selectedIndex == -1)) {
		newOpt.text = cg.options[cg.selectedIndex].text;
		newOpt.value = cg.options[cg.selectedIndex].value;
		
		cg.options[cg.selectedIndex] = null;
		ag.options[ag.options.length] = newOpt;
	}
}

function getList() {
	var cg = document.forms[0].elements["cg"];
	var strGroups = "";
	
	for(var count=0; count <= cg.options.length -1; count++) {
		strGroups = strGroups + cg.options[count].value;
		if(count!=(cg.options.length -1)) { strGroups = strGroups + "###"; }
	}
	document.forms[0].elements["lstgroups"].value = strGroups;
	return true;		
}

</SCRIPT>

<%

Dim obj

strGroup = request("group")
aryContainer = split(strGroup, ",")

Set obj = GetObject("LDAP://" & ldapEncode(strGroup))
strName = mid(obj.name, instr(obj.name, "=") + 1)

%>
<TABLE WIDTH="800px" STYLE="border: 1px solid gray;" CELLSPACING=0 CELLPADDING=0 BORDER=0>
	<TR>
		<TD WIDTH="200px" VALIGN="top" STYLE="background: #74AAD7">
			<TABLE WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0 COLS=1>
				<TR>
					<TD WIDTH="200px" HEIGHT="128px" ALIGN="center" STYLE="background: #74AAD7;" VALIGN="top"><IMG SRC="iconGroup.gif" ALT="" WIDTH="128" HEIGHT="128" BORDER="0"></TD>
					</TD>
				</TR>
			</TABLE>
		</TD>
		<TD WIDTH="600px" ROWSPAN=3 VALIGN="top">
			<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 STYLE="width: 100%" COLS=1 ROWS=2>
				<TR>
					<TD STYLE="background: #74AAD7; padding: 10px; width: 100%" VALIGN="top">
						<SPAN STYLE="font-size: 14pt;"><%= strName %></SPAN><BR>
						<SPAN STYLE="font-size: 10pt;">Group Membership</SPAN>
					</TD>
				</TR>
				<TR>
					<TD STYLE="padding-top: 5px; padding-bottom: 5px;" VALIGN="top">
						<FORM METHOD="post" ACTION="groupgroupmembership.asp"  ONSUBMIT="getList();">
						<INPUT TYPE="hidden" NAME="group" VALUE="<%=strGroup%>">
						<INPUT TYPE="hidden" NAME="lstgroups" ID="lstgroups" VALUE="">
						<TABLE WIDTH="100%" BORDER=0 CELLSPACING=5 CELLPADDING=0 COLS=3 ROWS=2>
							<TR>
								<TD VALIGN="top" WIDTH="40%">
								<B>Available groups:</B><BR>
								<SELECT NAME="availgroups" SIZE="10" STYLE="width: 100%" ID="ag">
								<%
								aryGroups = ListGroups()
								if isArray(obj.memberOf) then
									strConfigGroups =  "###" & join(obj.MemberOf, "###") & "###"
								else
									strConfigGroups = ""
								end if
								for each group in aryGroups
									Set grpObj = GetObject("LDAP://" & ldapEncode(group))
									aryTemp = split(grpObj.distinguishedName, ",")
									strObjName = aryTemp(lbound(aryTemp))
									strName = mid(strObjName, instr(strObjName, "=") + 1)
									if ((instr(strConfigGroups, "###" & grpObj.distinguishedName & "###") = 0) and (not(grpObj.distinguishedName = obj.distinguishedName))) then
										response.write "<OPTION VALUE=""" & _
										ldapEncode(grpObj.distinguishedName) & _
										""">" & strName & "</OPTION>" & vbCrLf
									end if
								next
								set grpObj = nothing
								%>
								</SELECT>
								</TD>
								<TD WIDTH="*" ALIGN="center" VALIGN="middle">
									<INPUT TYPE="button" NAME="add" VALUE="Add >>" STYLE="width: 100px; margin-bottom: 3px;" ONCLICK="addGroup();"><BR>
									<INPUT TYPE="button" NAME="remove" VALUE="<< Remove" STYLE="width: 100px" ONCLICK="removeGroup();">
								</TD>
								<TD VALIGN="top" WIDTH="40%">
								<B>Selected groups:</B><BR>
								<SELECT NAME="configgroups" SIZE="10" STYLE="width: 100%" ID="cg">
								<%
								if isArray(obj.memberOf) then
									for each group in obj.MemberOf
										Set grpObj = GetObject("LDAP://" & ldapEncode(group))
										aryTemp = split(grpObj.distinguishedName, ",")
										strObjName = aryTemp(lbound(aryTemp))
										strName = mid(strObjName, instr(strObjName, "=") + 1)
										response.write "<OPTION VALUE=""" & ldapEncode(grpObj.distinguishedName) & """>" & strName & "</OPTION>" & vbCrLf
									next
								end if
								set grpObj = nothing
								%>
								</SELECT>
								</TD>
							</TR>
							<TR>
								<TD COLSPAN=3 ALIGN="center"><INPUT TYPE="submit" VALUE="Update Group Membership"></TD>
							</TR>
						</TABLE>
					</FORM>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>

</BODY>