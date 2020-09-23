<!-- #include file="common.asp" -->

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
strotherTelephone = obj.otherTelephone
strurl = obj.url
strstreetAddress = obj.streetAddress
strl = obj.l
strst = obj.st
strpostalCode = obj.postalCode
strc = obj.c
strpostOfficeBox = obj.postOfficeBox
struserPrincipalName = obj.userPrincipalName
strdc = obj.dc
strsAMAccountName = obj.sAMAccountName
struserWorkstations = obj.userWorkstations
strAccountLocked = Bool2String(obj.IsAccountLocked)
strprofilePath = obj.profilePath
strscriptPath = obj.scriptPath
strhomeDrive = obj.homeDrive
strhomePhone = obj.homePhone
strpager = obj.pager
strmobile = obj.mobile
strfacsimileTelephoneNumber = obj.facsimileTelephoneNumber
stripPhone = obj.ipPhone
strinfo = obj.info
strotherHomePhone = obj.otherHomePhone
strotherPager = obj.otherPager
strotherMobile = obj.otherMobile
strotherFacsimileTelephoneNumber = obj.otherFacsimileTelephoneNumber
strotherIpPhone = obj.otherIpPhone
strtitle = obj.title
strdepartment = obj.department
strcompany = obj.company
strmanager = resolveuser(obj.manager)
strdirectReports = obj.directReports
strmemberOf = join(resolvegroups(obj.memberOf), "<BR>")
strwhenCreated = obj.whenCreated
strwhenChanged = obj.whenChanged
'strlastlogin = obj.lastLogin
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
						<A HREF="editUser.asp?user=<%=MakeURL(strUser)%>">Edit account</A><BR>
						<A HREF="usergroups.asp?user=<%=MakeURL(strUser)%>">Group Membership</A><BR>
						<A HREF="userpassword.asp?user=<%=MakeURL(strUser)%>">Reset password</A><P>
						<A HREF="deleteObject.asp?obj=<%=MakeURL(strUser)%>">Delete account</A>
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
		</TD>
		<TD WIDTH="600px" ROWSPAN=3 VALIGN="top">
			<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 STYLE="width: 100%" COLS=1 ROWS=2>
				<TR>
					<TD STYLE="background: #74AAD7; padding: 10px; width: 100%" VALIGN="top">
						<SPAN STYLE="font-size: 14pt;"><%= strName %></SPAN><BR>
						<SPAN STYLE="font-size: 10pt;">User</SPAN>
					</TD>
				</TR>
				<TR>
					<TD STYLE="padding-top: 5px; padding-bottom: 5px;" VALIGN="top">
						<TABLE WIDTH="100%" BORDER=0 CELLSPACING=0 CELLPADDING=0><%

doRow "Object Name", strName
doRow "Given Name", strGivenName
doRow "Last Name", strsn
doRow "Display Name", strdisplayName
doRow "Description", strdescription
doRow "Physical Delivery Office", strphysicalDeliveryOfficeName
doRow "Telephone", strtelephoneNumber
doRow "Mail Address", strmail
doRow "Homepage", strwWWHomePage
doRow "Other Telephone", strotherTelephone
doRow "URLs", strurl
doRow "Street Address", strstreetAddress
doRow "City", strl
doRow "St", strst
doRow "Post Code", strpostalCode
doRow "Country", strc
doRow "PO Box", strpostOfficeBox
doRow "UserPrincipal Name", struserPrincipalName
doRow "DC", strdc
doRow "Account Name", strsAMAccountName
doRow "Workstations", struserWorkstations
doRow "Account Locked", strAccountLocked
doRow "Profile Path", strprofilePath
doRow "Login Script", strscriptPath
doRow "Home Drive Letter", strhomeDrive
doRow "Home Phone", strhomePhone
doRow "Pager", strpager
doRow "Mobile", strmobile
doRow "Fax Number", strfacsimileTelephoneNumber
doRow "IP Phone", stripPhone
doRow "Info", strinfo
doRow "Other Home Phone", strotherHomePhone
doRow "Other Pager", strotherPager
doRow "Other Mobile", strotherMobile
doRow "Other Fax Number", strotherFacsimileTelephoneNumber
doRow "Other IP Phone", strotherIpPhone
doRow "Job Title", strtitle
doRow "Department", strdepartment
doRow "Company", strcompany
doRow "Manager", strmanager
doRow "Reports to", strdirectReports
doRow "Object Created", strwhenCreated
doRow "Object Last Changed", strwhenChanged
doRow "Last Login", strlastlogin
doRow "Group Membership", strmemberOf

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