<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on Settings button
	buttonswitch = 5
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Find your profile
	set rsProfile = server.CreateObject("adodb.recordset")
	sql = "execute FindUserbyID " & session("idUser")
	set rsProfile = dbConnection.Execute(sql)
	
	'Get a list of the Customers
	'if you are a super users then you list all of the customers if not then just list the customers that you have access to.
	set rsCustomers = server.CreateObject("adodb.recordset")
	if session("idAccess") = "A" then
		sql = "execute ListCustomerNamesandIDs"
	else
		sql = "execute ListCustomerNamesandIDsbyAccess " & session("idUser")
	end if
	set rsCustomers = dbConnection.Execute(sql)
	
	'check to see if anything was requested
	if request("idCustomer") = "" then
		idCustomer = rsCustomers("idCustomer")
	else
		idCustomer = cint(request("idCustomer"))
	end if
	
	'look up the access level for the User if he is not a SuperUser
	if session("idAccess") <> "A" then
		set rsAccess = server.CreateObject("adodb.recordset")
		sql = "execute FindAdminAccessbyCustomerandUser " & session("idUser") & "," & idCustomer
		set rsAccess = dbConnection.Execute(sql)
		'load the idAccess into a variable
		tempidAccess = rsAccess("idAccess")
		'close the recordset
		rsAccess.Close
		set rsAccess = nothing		
	end if
	
	'Find your team members
	set rsTeam = server.CreateObject("adodb.recordset")
	sql = "execute ListTeamMembersbyCustomer " & idCustomer
	set rsTeam = dbConnection.Execute(sql)
	
	'Find your Customer Information
	set rsCustomer = server.CreateObject("adodb.recordset")
	sql = "execute FindCustomerbyID " & idCustomer
	set rsCustomer = dbConnection.Execute(sql)	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="settings.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/settings-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top"><table width="625" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15"><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="610"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15" colspan="2"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="2">
				    <table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td colspan="3"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Account Settings</strong></font></td>
                      </tr>
                      <tr>
                        <td colspan="3" bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">Listed below are all of the Pools that you have access to view. If you have administration access to any of the pools, links will appear to modify the information.</font></td>
                      </tr>
                      <tr>
                        <td colspan="3" height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                      </tr>
                      <tr bgcolor="#6699cc"> 
                        <td width="50%" align="right"><font size="1" face="Arial, Helvetica, sans-serif" color="#ffffff">Pool</font></td>
						<td><font size="1" face="Arial, Helvetica, sans-serif"><select name="idCustomer" size="1">
<%
	if not rsCustomers.eof then
		do until rsCustomers.eof
%>
								<option value="<%=rsCustomers("idCustomer")%>" <%if idCustomer = rsCustomers("idCustomer") then%>selected<%end if%>><%=trim(rsCustomers("chrCustomer"))%></option>
<%
		rsCustomers.movenext
		loop
	end if
%>
							</select></font></td>
							<td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><input type="submit" name="Submit" value="Submit"></font></td>
                      	</tr>
                    	</table>
				  </td>
                </tr>
                <tr> 
                  <td height="15" colspan="2"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr valign="top"> 
                  <td width="50%">
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><img src="images/doublearrows.gif" width="30" height="13"></td>
                        <td width="100%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>My Profile</strong>&nbsp;<font size="1">(<a href="editprofile.asp">edit Profile</a>)</font></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Name: <%=trim(rsProfile("chrFirst")) & " " & trim(rsProfile("chrLast"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Phone: <%=trim(rsProfile("chrPhone"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Fax: <%=trim(rsProfile("chrFax"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Email: <%=trim(rsProfile("chrEmail"))%></font></td>
                      </tr>
                      <tr> 
                        <td colspan="2">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td><img src="images/doublearrows.gif" width="30" height="13"></td>
                        <td width="100%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Team Members </strong><%if tempidAccess < "P" then%><font size="1">(<a href="addteammember.asp?idCustomer=<%=idCustomer%>">add Team Member</a>)</font><%end if%></font></td>
                      </tr>
<%
	if not rsTeam.EOF then
		do until rsTeam.EOF
%>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsTeam("chrLast")) & ", " & trim(rsTeam("chrFirst"))%> - <%=trim(rsTeam("chrAccess"))%>&nbsp;<%if tempidAccess < "P" then%>(<A HREF="deleteaccount.asp?idAccount=<%=rsTeam("idAccount")%>&idCustomer=<%=idCustomer%>">Remove</A>)<%end if%></font></td>
                      </tr>
<%
		rsTeam.MoveNext
		loop
	end if
%>
                    </table>
                  </td>
                  <td width="50%">
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><img src="images/doublearrows.gif" width="30" height="13"></td>
                        <td width="100%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Customer Information </strong><%if tempidAccess < "P" then%><font size="1">(<a href="editcustomer.asp?idCustomer=<%=idCustomer%>">edit information</a>)</font><%end if%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%" bgcolor="#f5f5f5" align="center"><font size="1" face="Arial, Helvetica, sans-serif">Contact Information</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Company: <%=trim(rsCustomer("chrCustomer"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Name: <%=trim(rsCustomer("chrCName"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Email: <%=trim(rsCustomer("chrCEmail"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Address: <%=trim(rsCustomer("chrCAddress"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Address: <%=trim(rsCustomer("chrCAddress2"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">City: <%=trim(rsCustomer("chrCCity"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">State: <%=trim(rsCustomer("chrCState"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Zip: <%=trim(rsCustomer("chrCZip"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Phone: <%=trim(rsCustomer("chrCPhone"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Fax: <%=trim(rsCustomer("chrCFax"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%" bgcolor="#f5f5f5" align="center"><font size="1" face="Arial, Helvetica, sans-serif">Billing Information</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Company: <%=trim(rsCustomer("chrCustomer"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Name: <%=trim(rsCustomer("chrBName"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Email: <%=trim(rsCustomer("chrBEmail"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Address: <%=trim(rsCustomer("chrBAddress"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Address: <%=trim(rsCustomer("chrBAddress2"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">City: <%=trim(rsCustomer("chrBCity"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">State: <%=trim(rsCustomer("chrBState"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Zip: <%=trim(rsCustomer("chrBZip"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Phone: <%=trim(rsCustomer("chrBPhone"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">Fax: <%=trim(rsCustomer("chrBFax"))%></font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15" colspan="2"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <!-- #Begin bottom part -->
    <!-- #include file="includes/bottom.htm" -->
  </table>
  </form>
</body>
</html>
<%
	rsProfile.Close
	set rsProfile = nothing
	rsTeam.Close
	set rsTeam = nothing
	rsCustomer.Close
	set rsCustomer = nothing
	dbConnection.Close
	set dbConnection = nothing
%>