<!-- #include file="common.asp" -->

<TABLE WIDTH="800px" STYLE="border: 1px solid gray;" CELLSPACING=0 CELLPADDING=0 BORDER=0>
	<TR>
		<TD WIDTH="200px" ALIGN="center" STYLE="background: #74AAD7;" VALIGN="top"><IMG SRC="iconSearch.gif" ALT="" WIDTH="128" HEIGHT="128" BORDER="0"></TD>
		<TD WIDTH="600px" ROWSPAN=3 VALIGN="top">
			<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 STYLE="width: 100%" COLS=1 ROWS=2>
				<TR>
					<TD STYLE="background: #74AAD7; padding: 10px; width: 100%" VALIGN="top">
						<SPAN STYLE="font-size: 14pt;">Advanced Search</SPAN><BR>
						<SPAN STYLE="font-size: 10pt;">&nbsp;</SPAN>
					</TD>
				</TR>
				<TR>
					<TD STYLE="padding-top: 5px; padding-bottom: 5px;" VALIGN="top">
						<TABLE WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD WIDTH="100%" CLASS="data" VALIGN="top">
									<FORM METHOD="GET" ACTION="search.asp">
									<SPAN STYLE="font-size: 10pt; font-weight: bold;">Search</SPAN><BR>
									<INPUT TYPE="text" NAME="searchstr" STYLE="width: 175px;"><P>
									<B>Object Type:</B><BR>
									<SELECT NAME="objectType" STYLE="width: 175px;">
										<OPTION SELECTED VALUE="any">Any</OPTION>
										<OPTION VALUE="user">User</OPTION>
										<OPTION VALUE="group">Group</OPTION>
										<OPTION VALUE="computer">Computer</OPTION>
										<OPTION VALUE="printer">Printer</OPTION>
										<OPTION VALUE="contact">Contact</OPTION>
									</SELECT><P>
									<B>Attributes:</B><BR>
									<INPUT TYPE="checkbox" NAME="accDisabled" VALUE="true">&nbsp;&nbsp;Account Disabled<P>
									<INPUT TYPE="submit" VALUE="  Search  "></FORM>				
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</BODY>