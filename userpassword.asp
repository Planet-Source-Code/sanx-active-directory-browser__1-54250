<!-- #include file="common.asp" -->

<%

Dim obj

strUser = request("user")

Set obj = GetObject("LDAP://" & ldapEncode(strUser))
strName = mid(obj.name, instr(obj.name, "=") + 1)

%>

<TABLE WIDTH="800px" STYLE="border: 1px solid gray;" CELLSPACING=0 CELLPADDING=0 BORDER=0>
	<TR>
		<TD WIDTH="200px" ALIGN="center" STYLE="background: #74AAD7;" VALIGN="top"><IMG SRC="iconUser.gif" ALT="" WIDTH="128" HEIGHT="128" BORDER="0"></TD>
		<TD WIDTH="600px" ROWSPAN=3 VALIGN="top">
			<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 STYLE="width: 100%" COLS=1 ROWS=2>
				<TR>
					<TD STYLE="background: #74AAD7; padding: 10px; width: 100%" VALIGN="top">
						<SPAN STYLE="font-size: 14pt;"><%=strName%></SPAN><BR>
						<SPAN STYLE="font-size: 10pt;">Reset Password</SPAN>
					</TD>
				</TR>
				<TR>
					<TD STYLE="padding-top: 5px; padding-bottom: 5px;" VALIGN="top">
						<TABLE WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD WIDTH="100%" CLASS="data" VALIGN="top">
									<FORM METHOD="post" ACTION="updatepassword.asp">
									<INPUT TYPE="hidden" NAME="user" VALUE="<%=strUser%>">
									Enter password:<BR>
									<INPUT TYPE="password" NAME="pass1"><P>
									Confirm password:<BR>
									<INPUT TYPE="password" NAME="pass2"><P>
									<INPUT TYPE="submit" VALUE="Change password">
									</FORM>						
								</TD>
							</TR>
<%

Set obj = Nothing

%>

						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</BODY>