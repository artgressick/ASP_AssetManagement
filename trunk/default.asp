<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp?idUrl=default.asp"
	end if
	
	if session("idAccess") < "O" then
		Response.Redirect "default-super.asp"
	elseif session("idAccess") = "O" then
		Response.Redirect "default-manager.asp"
	else
		Response.Redirect "default-user.asp"
	end if 
	
	'turn on order button
	buttonswitch = 1
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Get the Statistics
	set rsStats = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute StatisticsbyAdmin"
	else
		sql = "execute StatisticsbyUser " & session("idUser")
	end if
	set rsStats = dbConnection.Execute(sql)
	
	'Get a list of orders that need approval
	set rsApproval = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListCartsforApprovalAll"
	else
		sql = "execute ListCartsforApprovalAllAccess " & session("idUser")
	end if
	set rsApproval = dbConnection.Execute(sql)
	
	'get a list of orders that will ship in 7 days
	set rsShipping = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListShippingin7DaysAll"
	else
		sql = "execute ListShippingin7DaysAllAccess " & session("idUser")
	end if
	set rsShipping = dbConnection.Execute(sql)
	
	'get a list of orders that will return in 7 days
	set rsReturning = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListReturningin7DaysAll"
	else
		sql = "execute ListReturningin7DaysAllAccess " & session("idUser")
	end if
	set rsReturning = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/home-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top"><table width="625" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15"><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="610"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Welcome!</strong></font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">This 
                          is the homepage of the new techIT Solutions Asset Management System.  
                          &nbsp;To navigate through the site, please use the black tabs above and the
						  navigation links on the left. &nbsp;These links will change in accordance with the
						  black tab you have selected. &nbsp;Should you have any questions, or encounter any
						  system anomalies, please contact your Pool Manager or your techIT Solutions Account Manager.</font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="300" valign="top">
						  <table width="100%" border="0" cellspacing="0" cellpadding="1">
                            <tr> 
                              <td bgcolor="#c0c0c0">
								<table width="100%" border="0" cellspacing="0" cellpadding="3">
                                  <tr> 
                                    <td colspan="2" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Quick Reports</strong></font></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#ffffff" width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportwhatsavailable.asp">What's Available</a></font></td>
                                    <td bgcolor="#ffffff" width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportinventory.asp">Inventory Report</a></font></td>
                                  </tr>
                                </table>
                              </td>
                            </tr>
                          </table>
                        </td>
                        <td width="10">&nbsp;</td>
                        <td width="300" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="1">
                            <tr> 
                              <td bgcolor="#c0c0c0"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                                  <tr> 
                                    <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Inventory Statistics</strong></font></td>
									<td bgcolor="#f5f5f5" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportcharts.asp">View Graph</a></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Total Assets: <%=rsStats("intTotal")%></font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportinventory.asp?idCustomer=0&idStatus=8&idWarehouse=0">Damaged</a>: <%=rsStats("intBroken")%></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportinventory.asp?idCustomer=0&idStatus=1&idWarehouse=0">Assets Ready</a>: <%=rsStats("intReady")%></font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportinventory.asp?idCustomer=0&idStatus=9&idWarehouse=0">Lost</a>: <%=rsStats("intLost")%></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportinventory.asp?idCustomer=0&idStatus=2&idWarehouse=0">Assets Out</a>: <%=rsStats("intOut")%></font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportinventory.asp?idCustomer=0&idStatus=7&idWarehouse=0">Internal Use</a>: <%=rsStats("intInternal")%></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportinventory.asp?idCustomer=0&idStatus=3&idWarehouse=0">Need Turning</a>: <%=rsStats("intTurning")%></font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportinventory.asp?idCustomer=0&idStatus=6&idWarehouse=0">Out of System</a>: <%=rsStats("intOutSystem")%></font></td>
                                  </tr>
                                </table></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr bgcolor="#6699cc"> 
                        <td height="25" colspan="6"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong>&nbsp;Carts Awaiting Approval</strong></font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Cart</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Customer</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Ship</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Departure</font></td>
                        <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">Assets&nbsp;</font></td>
                        <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif">Options&nbsp;</font></td>
                      </tr>
<%
	bgcolor = 0
	if rsApproval.EOF then
%>
                      <tr align="center"> 
                        <td height="20" colspan="6"><font size="1" face="Arial, Helvetica, sans-serif">There are no carts that need approval.</font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="6"><font size="1" face="Arial, Helvetica, sans-serif"><img src="images/c0c0c0dot.gif" width="1" height="1"></font></td>
                      </tr>
<%
	else
		do until rsApproval.EOF
		if bgswitch = 1 then
			bgcolor = "#f5f5f5"
			bgswitch = 0
		else
			bgcolor = "#ffffff"
			bgswitch = 1
		end if
%>
                      <tr bgcolor="<%=bgcolor%>"> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsApproval("chrCart"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsApproval("chrCustomer"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsApproval("dtShip"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsApproval("dtDeparture"),2)%></font></td>
                        <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=rsApproval("intOrdered")%></font></td>
                        <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><A HREF="approvecart.asp?idCart=<%=rsApproval("idCart")%>">View Details</a>&nbsp;</font></td>
                      </tr>
					  <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="6"><font size="1" face="Arial, Helvetica, sans-serif"><img src="images/c0c0c0dot.gif" width="1" height="1"></font></td>
                      </tr>
<%
		rsApproval.MoveNext
		loop
	end if
%>
                    </table></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="25" colspan="7" bgcolor="#6699cc">&nbsp;<font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong>Shipping in 7 Days</strong><font size="1">&nbsp;&nbsp;(<a href="reportshipping.asp" class="titlelink"><U>Click here to view all orders that need to be shipped.</U></a>)</font></font></td>
                      </tr>
                      <tr> 
                        <td height="20" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Cart</font></td>
                        <td height="20" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Customer</font></td>
                        <td height="20" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Ship Date</font></td>
                        <td height="20" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Arrival</font></td>
                        <td height="20" align="center" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">Assets&nbsp;</font></td>
                        <td height="20" align="right" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">Options&nbsp;</font></td>
                      </tr>
<%
	bgcolor = 0
	if rsShipping.EOF then
%>
                      <tr> 
                        <td height="20" colspan="6" align="center"><font size="1" face="Arial, Helvetica, sans-serif">There are no Carts that need to be shipped in 7 days.</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="6" bgcolor="#c0c0c0"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsShipping.EOF
		if bgswitch = 1 then
			bgcolor = "#ffffff"
			bgswitch = 0
		else
			bgcolor = "#f5f5f5"
			bgswitch = 1
		end if
%>
                      <tr bgcolor="<%=bgcolor%>">
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsShipping("chrCart"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsShipping("chrCustomer"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsShipping("dtShip"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsShipping("dtArrival"),2)%></font></td>
                        <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=rsShipping("intAssets")%>&nbsp;</font></td>
                        <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><a href="pulllist.asp?idCart=<%=rsShipping("idCart")%>" target="_blank">Pull List</a>&nbsp;</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="6" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif"><img src="images/c0c0c0dot.gif" width="1" height="1"></font></td>
                      </tr>
<%
		rsShipping.MoveNext
		loop
	end if
%>
                    </table></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="20" colspan="6" bgcolor="#6699cc"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong>&nbsp;Due Back in 7 Days</strong><font size="1">&nbsp;&nbsp;(<a href="reportcartsreturning.asp" class="titlelink"><U>Click here to view all returning orders.</U></a>)</font></font></td>
                      </tr>
                      <tr> 
                        <td height="20" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Cart</font></td>
                        <td height="20" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Customer</font></td>
                        <td height="20" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Departure</font></td>
                        <td height="20" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Return</font></td>
                        <td height="20" align="center" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">Assets&nbsp;</font></td>
                        <td height="20" align="right" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">Options&nbsp;</font></td>
                      </tr>
<%
	bgcolor = 0
	if rsReturning.EOF then
%>
                      <tr> 
                        <td height="20" colspan="6" align="center"><font size="1" face="Arial, Helvetica, sans-serif">There are no carts due back in 7 days.</font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="6"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsReturning.EOF
		if bgswitch = 1 then
			bgcolor = "#ffffff"
			bgswitch = 0
		else
			bgcolor = "#f5f5f5"
			bgswitch = 1
		end if
%>
                      <tr bgcolor="<%=bgcolor%>">
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsReturning("chrCart"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsReturning("chrCustomer"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsReturning("dtDeparture"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsReturning("dtReturn"),2)%></font></td>
                        <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=rsReturning("intAssets")%>&nbsp;</font></td>
                        <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><a href="viewcart.asp?idCart=<%=rsReturning("idCart")%>">View Details</a>&nbsp;</font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="6"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
                      </tr>
<%
		rsReturning.MoveNext
		loop
	end if
%>
                    </table></td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <!-- #Begin bottom part -->
    <!-- #include file="includes/bottom.htm" -->
  </table>
</body>
</html>
<%
	rsStats.Close
	set rsStats = nothing
	rsApproval.Close
	set rsApproval = nothing
	rsShipping.Close
	set rsShipping = nothing
	rsReturning.Close
	set rsReturning = nothing
	dbConnection.Close
	set dbConnection = nothing
%>