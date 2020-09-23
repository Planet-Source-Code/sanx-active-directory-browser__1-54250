<!-- #include file="common.asp" -->

<SCRIPT LANGUAGE="JavaScript">

function writeName(obj){
	if (obj.value == "") {
		obj.value = document.forms["edit"].elements["gn"].value + " " + document.forms["edit"].elements["sn"].value;
	}
}

</SCRIPT>

<%

Dim obj

strUser = request("user")
aryContainer = split(strUser, ",")

Set obj = GetObject("LDAP://" & ldapEncode(strUser))

strName = mid(obj.name, instr(obj.name, "=") + 1)
strgivenName = obj.givenName
strinitials = obj.initials
strsn = obj.sn
strdisplayName = obj.displayName
strdescription = obj.description
strphysicalDeliveryOfficeName = obj.physicalDeliveryOfficeName
strtelephoneNumber = obj.telephoneNumber
strmail = obj.mail
strwWWHomePage = obj.wWWHomePage
strstreetAddress = obj.streetAddress
strl = obj.l
strst = obj.st
strpostalCode = obj.postalCode
strc = obj.c
strsAMAccountName = obj.sAMAccountName
strprofilePath = obj.profilePath
strscriptPath = obj.scriptPath
strhomeDrive = obj.homeDrive
strhomePhone = obj.homePhone
strpager = obj.pager
strmobile = obj.mobile
strfacsimileTelephoneNumber = obj.facsimileTelephoneNumber
stripPhone = obj.ipPhone
strtitle = obj.title
strdepartment = obj.department
strcompany = obj.company
strmanager = obj.manager

Set obj = Nothing
%>
<TABLE WIDTH="800px" STYLE="border: 1px solid gray;" CELLSPACING=0 CELLPADDING=0 BORDER=0>
	<TR>
		<TD WIDTH="200px" VALIGN="top" STYLE="background: #74AAD7">
			<TABLE WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0 COLS=1>
				<TR>
					<TD WIDTH="200px" HEIGHT="128px" ALIGN="center" STYLE="background: #74AAD7;" VALIGN="top"><IMG SRC="iconUser.gif" ALT="" WIDTH="128" HEIGHT="128" BORDER="0"></TD>
					</TD>
				</TR>
				<TR>
					<TD WIDTH="200px" STYLE="padding: 10px; background: #74AAD7;" VALIGN="top">
						<DIV STYLE="font-size: 10pt; font-weight: bold; padding-bottom: 5px;">User Tasks</DIV>
						<A HREF="userdetails.asp?user=<%=MakeURL(strUser)%>">Cancel</A></A>
					</TD>
				</TR>
			</TABLE>
		</TD>
		<TD WIDTH="600px" ROWSPAN=3 VALIGN="top">
			<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 STYLE="width: 100%" COLS=1 ROWS=2>
				<TR>
					<TD STYLE="background: #74AAD7; padding: 10px; width: 100%" VALIGN="top">
						<SPAN STYLE="font-size: 14pt;"><%= strName %></SPAN><BR>
						<SPAN STYLE="font-size: 10pt;">Edit User</SPAN>
					</TD>
				</TR>
				<TR>
					<TD STYLE="padding: 3px;" VALIGN="top">
						<FORM METHOD="get" ACTION="updateUser.asp" NAME="edit">
						<INPUT TYPE="hidden" NAME="user" VALUE="<%=strUser%>">
						<INPUT TYPE="hidden" NAME="action" VALUE="edit">
						<TABLE WIDTH="100%" BORDER=0 CELLSPACING=5 CELLPADDING=0>
							<TR>
								<TD VALIGN="top" COLSPAN=2 CLASS="sectionhead" STYLE="padding-top: 0px;">General</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Given name:<BR><INPUT TYPE="Text" NAME="givenName" VALUE="<%=strGivenName%>" ID="gn">
								</TD>
								<TD VALIGN="top">Initials:<BR><INPUT TYPE="Text" NAME="initials" VALUE="<%=strinitials%>">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Last name:<BR><INPUT TYPE="Text" NAME="sn" VALUE="<%=strsn%>" ID="sn">
								</TD>
								<TD VALIGN="top" >Display name:<BR><INPUT TYPE="Text" NAME="displayname" VALUE="<%=strdisplayname%>" ONFOCUS="writeName(this);" ID="dn">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2>Description:<BR><INPUT TYPE="Text" NAME="description" VALUE="<%=strDescription%>" SIZE=77>
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Physical office:<BR><INPUT TYPE="Text" NAME="physicaldeliveryofficename" VALUE="<%=strphysicaldeliveryofficename%>">
								</TD>
								<TD VALIGN="top">Telephone number:<BR><INPUT TYPE="Text" NAME="telehonenumber" VALUE="<%=strtelephonenumber%>">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Mail address:<BR><INPUT TYPE="Text" NAME="mail" VALUE="<%=strmail%>">
								</TD>
								<TD VALIGN="top">Homepage:<BR><INPUT TYPE="Text" NAME="wwwhomepage" VALUE="<%=strwwwhomepage%>">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2 CLASS="sectionhead">Address</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2>Street Address:<BR><INPUT TYPE="Text" NAME="streetaddress" VALUE="<%=strstreetaddress%>" SIZE=77>
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">City:<BR><INPUT TYPE="Text" NAME="l" VALUE="<%=strl%>">
								</TD>
								<TD VALIGN="top">State / Province:<BR><INPUT TYPE="Text" NAME="st" VALUE="<%=strst%>">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Country:<BR><INPUT TYPE="Text" NAME="c" VALUE="<%=strc%>" MAXLENGTH=2>
								</TD>
								<TD VALIGN="top">Post code:<BR><INPUT TYPE="Text" NAME="postalcode" VALUE="<%=strpostalcode%>">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2 CLASS="sectionhead">Account</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Account name:<BR><INPUT TYPE="Text" NAME="samaccountname" VALUE="<%=strsamaccountname%>">
								</TD>
								<TD VALIGN="top">Login script:<BR><INPUT TYPE="Text" NAME="scriptpath" VALUE="<%=strscriptpath%>">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Profile path:<BR><INPUT TYPE="Text" NAME="profilepath" VALUE="<%=strprofilepath%>">
								</TD>
								<TD VALIGN="top">Home drive:<BR><INPUT TYPE="Text" NAME="homedrive" VALUE="<%=strhomedrive%>">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2 CLASS="sectionhead">Telephone Numbers</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Home phone:<BR><INPUT TYPE="Text" NAME="homephone" VALUE="<%=strhomephone%>">
								</TD>
								<TD VALIGN="top">Mobile phone:<BR><INPUT TYPE="Text" NAME="mobile" VALUE="<%=strmobile%>">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Fax:<BR><INPUT TYPE="Text" NAME="facsimiletelephonenumber" VALUE="<%=strfacsimiletelephonenumber%>">
								</TD>
								<TD VALIGN="top">Pager:<BR><INPUT TYPE="Text" NAME="pager" VALUE="<%=strpager%>">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">IP phone:<BR><INPUT TYPE="Text" NAME="ipphone" VALUE="<%=stripphone%>">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2 CLASS="sectionhead">Organisation</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Title:<BR><INPUT TYPE="Text" NAME="title" VALUE="<%=strtitle%>">
								</TD>
								<TD VALIGN="top">Department:<BR><INPUT TYPE="Text" NAME="department" VALUE="<%=strdepartment%>">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2>Company:<BR><INPUT TYPE="Text" NAME="company" VALUE="<%=strcompany%>" SIZE="77">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2 CLASS="submit"><INPUT TYPE="submit" VALUE="  Apply  ">
								</TD>
							</TR>
						</TABLE>
						</FORM>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>

<SCRIPT LANUGUAGE="Javascript">

if(document.forms["edit"].elements["dn"].value == "") {
	document.forms["edit"].elements["dn"].value = document.forms["edit"].elements["gn"].value + " " + document.forms["edit"].elements["sn"].value;
}

</SCRIPT>

</BODY>