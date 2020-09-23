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

strContainer = request("container")
aryContainer = split(strContainer, ",")
strOUName = right(aryContainer(lbound(aryContainer)), len(aryContainer(lbound(aryContainer))) - 3)

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
						<A HREF="index.asp?container=<%=MakeURL(strContainer)%>">Cancel</A></A>
					</TD>
				</TR>
			</TABLE>
		</TD>
		<TD WIDTH="600px" ROWSPAN=3 VALIGN="top">
			<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 STYLE="width: 100%" COLS=1 ROWS=2>
				<TR>
					<TD STYLE="background: #74AAD7; padding: 10px; width: 100%" VALIGN="top">
						<SPAN STYLE="font-size: 14pt;">New User</SPAN><BR>
						<SPAN STYLE="font-size: 10pt;">Create in: <%=strOUName%></SPAN>
					</TD>
				</TR>
				<TR>
					<TD STYLE="padding: 3px;" VALIGN="top">
						<FORM METHOD="get" ACTION="updateUser.asp" NAME="edit">
						<INPUT TYPE="hidden" NAME="container" VALUE="<%=strContainer%>">
						<INPUT TYPE="hidden" NAME="action" VALUE="new">
						<TABLE WIDTH="100%" BORDER=0 CELLSPACING=5 CELLPADDING=0>
							<TR>
								<TD VALIGN="top" COLSPAN=2 CLASS="sectionhead">Account</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Account name:<BR><INPUT TYPE="Text" NAME="samaccountname">
								</TD>
								<TD VALIGN="top">Login script:<BR><INPUT TYPE="Text" NAME="scriptpath">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Profile path:<BR><INPUT TYPE="Text" NAME="profilepath">
								</TD>
								<TD VALIGN="top">Home drive:<BR><INPUT TYPE="Text" NAME="homedrive">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2 CLASS="sectionhead" STYLE="padding-top: 0px;">General</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Given name:<BR><INPUT TYPE="Text" NAME="givenName" ID="gn">
								</TD>
								<TD VALIGN="top">Initials:<BR><INPUT TYPE="Text" NAME="initials">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Last name:<BR><INPUT TYPE="Text" NAME="sn" ID="sn">
								</TD>
								<TD VALIGN="top" >Display name:<BR><INPUT TYPE="Text" NAME="displayname" ONFOCUS="writeName(this);" ID="dn">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2>Description:<BR><INPUT TYPE="Text" NAME="description" SIZE=77>
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Physical office:<BR><INPUT TYPE="Text" NAME="physicaldeliveryofficename">
								</TD>
								<TD VALIGN="top">Telephone number:<BR><INPUT TYPE="Text" NAME="telehonenumber">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Mail address:<BR><INPUT TYPE="Text" NAME="mail">
								</TD>
								<TD VALIGN="top">Homepage:<BR><INPUT TYPE="Text" NAME="wwwhomepage">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2 CLASS="sectionhead">Address</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2>Street Address:<BR><INPUT TYPE="Text" NAME="streetaddress" SIZE=77>
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">City:<BR><INPUT TYPE="Text" NAME="l">
								</TD>
								<TD VALIGN="top">State / Province:<BR><INPUT TYPE="Text" NAME="st">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Country:<BR><INPUT TYPE="Text" NAME="c" MAXLENGTH=2>
								</TD>
								<TD VALIGN="top">Post code:<BR><INPUT TYPE="Text" NAME="postalcode">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2 CLASS="sectionhead">Telephone Numbers</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Home phone:<BR><INPUT TYPE="Text" NAME="homephone">
								</TD>
								<TD VALIGN="top">Mobile phone:<BR><INPUT TYPE="Text" NAME="mobile">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Fax:<BR><INPUT TYPE="Text" NAME="facsimiletelephonenumber">
								</TD>
								<TD VALIGN="top">Pager:<BR><INPUT TYPE="Text" NAME="pager">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top">IP phone:<BR><INPUT TYPE="Text" NAME="ipphone">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2 CLASS="sectionhead">Organisation</TD>
							</TR>
							<TR>
								<TD VALIGN="top">Title:<BR><INPUT TYPE="Text" NAME="title">
								</TD>
								<TD VALIGN="top">Department:<BR><INPUT TYPE="Text" NAME="department">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2>Company:<BR><INPUT TYPE="Text" NAME="company" SIZE="77">
								</TD>
							</TR>
							<TR>
								<TD VALIGN="top" COLSPAN=2 CLASS="submit"><INPUT TYPE="submit" VALUE="  Create  ">
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