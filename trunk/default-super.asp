<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp?idUrl=default.asp"
	end if
	
	'turn on order button
	buttonswitch = 1
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'check to see if they entered anything for the radio buttons
	if request("idOWarehouse") = "" or request("idIWarehouse") = "" then
		'now check for any cookies
		if Request.Cookies.Item("idOWarehouse") = "" then
			'we need to write a cookie
			Response.Cookies.Item("idOWarehouse") = 0
			Response.Cookies.Item("idOWarehouse").Expires = #December 31, 2005#
			'load warehouse
			idOWarehouse = 0
		else
			idOWarehouse = Request.Cookies.Item("idOWarehouse")
		end if
		if Request.Cookies.Item("idIWarehouse") = "" then
			'we need to write a cookie
			Response.Cookies.Item("idIWarehouse") = 0
			Response.Cookies.Item("idIWarehouse").Expires = #December 31, 2005#
			'load warehouse
			idIWarehouse = 0
		else
			idIWarehouse = Request.Cookies.Item("idIWarehouse")
		end if
	else
		'update the cookie
		Response.Cookies.Item("idOWarehouse") = request("idOWarehouse")
		Response.Cookies.Item("idOWarehouse").Expires = #December 31, 2005#
		'update the cookie
		Response.Cookies.Item("idIWarehouse") = request("idIWarehouse")
		Response.Cookies.Item("idIWarehouse").Expires = #December 31, 2005#
		'get the request variables and store locally
		idOWarehouse = request("idOWarehouse")
		idIWarehouse = request("idIWarehouse")
	end if	
		
	'Get the Statistics
	set rsStats = server.CreateObject("adodb.recordset")
	sql = "execute StatisticsbyAdmin"
	set rsStats = dbConnection.Execute(sql)
	
	'Get a list of orders that need approval
	set rsNew = server.CreateObject("adodb.recordset")
	sql = "execute ListCartsPreArrivalbyWarehouse " & idOWarehouse
	set rsNew = dbConnection.Execute(sql)
	
	'get a list of orders that will return in 7 days
	set rsReturning = server.CreateObject("adodb.recordset")
	sql = "execute ListCartsReturningbyWarehouse " & idIWarehouse
	set rsReturning = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="default-super.asp">
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
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>techIT Dashboard</strong></font></td>
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
                                    <td colspan="2" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Quick Reports/Information</strong></font></td>
                                  </tr>
                                  <tr> 
                                    <td bgcolor="#ffffff" width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportwhatsavailable.asp">What's Available</a></font></td>
                                    <td bgcolor="#ffffff" width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportinventory.asp">Inventory Report</a></font></td>
                                  </tr>
								  <tr> 
                                    <td bgcolor="#ffffff" width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="corp-pricing.pdf">Corporate Events Pool Fees</a></font></td>
                                    <td bgcolor="#ffffff" width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="packingsheets.asp">Asset Packing Sheets</a></font></td>
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
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr bgcolor="#6699cc"> 
                        <td height="25" colspan="7"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong>&nbsp;New/Awaiting Approval Carts</strong><br>
                        <font size="1">&nbsp;Carts in <font color="#ffff00">YELLOW</font> have not been approved and are within 3 days of shipping.<br>
                        &nbsp;Carts in <font color="#ff0000">RED</font> have not been approved and are within 1 day of shipping.</font></td>
                      </tr>
                      <tr bgcolor="#5b5b5b"> 
                        <td height="25" colspan="7">
						  <table width="100%" border="0" cellspacing="0" cellpadding="3">
							<tr>
							  <td><font size="2" face="Arial, Helvetica, sans-serif" color="#ffffff"><strong>Warehouse Filter</strong></font></td>
							  <td><font size="2" face="Arial, Helvetica, sans-serif" color="#ffffff"><input type="radio" name="idOWarehouse" value="0" <%if cint(request("idOWarehouse")) = 0 then%>checked<%end if%>>&nbsp;All Warehouses</font></td>
							  <td><font size="2" face="Arial, Helvetica, sans-serif" color="#ffffff"><input type="radio" name="idOWarehouse" value="1" <%if cint(request("idOWarehouse")) = 1 then%>checked<%end if%>>&nbsp;Livermore</font></td>
							  <td><font size="2" face="Arial, Helvetica, sans-serif" color="#ffffff"><input type="radio" name="idOWarehouse" value="2" <%if cint(request("idOWarehouse")) = 2 then%>checked<%end if%>>&nbsp;Knoxville</font></td>
							  <td><font size="2" face="Arial, Helvetica, sans-serif" color="#ffffff"><input type="submit" name="Submit" value="Filter"></td>
							</tr>
						  </table>
                        </td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Cart</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Pool</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Ship</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Arrival</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Status</font></td>
                        <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">Assets&nbsp;</font></td>
                      </tr>
<%
	bgcolor = 0
	if rsNew.EOF then
%>
                      <tr align="center"> 
                        <td height="20" colspan="7"><font size="1" face="Arial, Helvetica, sans-serif">No carts have been entered or all carts are returning.</font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="7"><font size="1" face="Arial, Helvetica, sans-serif"><img src="images/c0c0c0dot.gif" width="1" height="1"></font></td>
                      </tr>
<%
	else
		do until rsNew.EOF
		'this is a RED warning
		if (rsNew("idStatus") = 1 or rsNew("idStatus") = 2) and datediff("d",date,rsNew("dtShip")) <= 1 then
			bgcolor = "#ff0000"
			fcolor="#ffff00"
		'this is a yellow warning
		elseif (rsNew("idStatus") = 1 or rsNew("idStatus") = 2) and datediff("d",date,rsNew("dtShip")) <= 3 then
			bgcolor = "#ffff00"
			fcolor="#000000"
		else
			bgcolor = "#ffffff"
			fcolor = "#000000"
		end if
%>
                      <tr bgcolor="<%=bgcolor%>">
<%
	'check to see what kind of Order type this is going to be
	select case rsNew("idType")
		case 1
			img = "circle.gif"
			alt = "Standard Cart"			
		case 2
			img = "square.gif"
			alt = "Internal Use Cart"
		case 3
			img = "animalprint.gif"
			alt = "Out of System"
	end select
	'check to see if there is a Show to Show and use lightning
	if cint(rsNew("idShow2Show")) = 1 then
		img = "lightning.gif"
		alt = "Show to Show Cart"
	end if
%>
						<td height="20" align="center" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif"><IMG SRC="images/<%=img%>" alt="<%=alt%>"></font></td>
<%
		'this is for the approval process
		if rsNew("idStatus") = 1 or rsNew("idStatus") = 2 then
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<A HREF="approvecart.asp?idCart=<%=rsNew("idCart")%>"><%=trim(rsNew("chrCart"))%></a></font></td>
<%
		else
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<A HREF="viewcart.asp?idCart=<%=rsNew("idCart")%>"><%=trim(rsNew("chrCart"))%></a></font></td>
<%
		end if
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=trim(rsNew("chrCustomer"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=formatdatetime(rsNew("dtShip"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=formatdatetime(rsNew("dtArrival"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=trim(rsNew("chrCartStatus"))%></font></td>
                        <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=rsNew("intOrdered")%></font></td>
                      </tr>
					  <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="7"><font size="1" face="Arial, Helvetica, sans-serif"><img src="images/c0c0c0dot.gif" width="1" height="1"></font></td>
                      </tr>
<%
		rsNew.MoveNext
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
                        <td height="25" colspan="7" bgcolor="#6699cc">&nbsp;<font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong>Carts Returning</strong><br>
                        <font size="1">&nbsp;Carts in <font color="#ffff00">YELLOW</font> are 4 days or more from the return date.<br>
                        &nbsp;Carts in <font color="#ff0000">RED</font> are 10 days or more from the return date and have not returned or have been checked in completely.</font></font></td>
                      </tr>
                      <tr bgcolor="#5b5b5b"> 
                        <td height="25" colspan="7">
						  <table width="100%" border="0" cellspacing="0" cellpadding="3">
							<tr>
							  <td><font size="2" face="Arial, Helvetica, sans-serif" color="#ffffff"><strong>Warehouse Filter</strong></font></td>
							  <td><font size="2" face="Arial, Helvetica, sans-serif" color="#ffffff"><input type="radio" name="idIWarehouse" value="0" <%if cint(request("idIWarehouse")) = 0 then%>checked<%end if%>>&nbsp;All Warehouses</font></td>
							  <td><font size="2" face="Arial, Helvetica, sans-serif" color="#ffffff"><input type="radio" name="idIWarehouse" value="1" <%if cint(request("idIWarehouse")) = 1 then%>checked<%end if%>>&nbsp;Livermore</font></td>
							  <td><font size="2" face="Arial, Helvetica, sans-serif" color="#ffffff"><input type="radio" name="idIWarehouse" value="2" <%if cint(request("idIWarehouse")) = 2 then%>checked<%end if%>>&nbsp;Knoxville</font></td>
							  <td><font size="2" face="Arial, Helvetica, sans-serif" color="#ffffff"><input type="submit" name="Submit" value="Filter"></td>
							</tr>
						  </table>
                        </td>
                      </tr>
                      <tr> 
                        <td height="20" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td height="20" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Cart</font></td>
                        <td height="20" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Customer</font></td>
                        <td height="20" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Departure</font></td>
                        <td height="20" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Returning</font></td>
                        <td height="20" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Status</font></td>
                        <td height="20" align="center" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif">Assets&nbsp;</font></td>
                      </tr>
<%
	bgcolor = 0
	if rsReturning.EOF then
%>
                      <tr> 
                        <td height="20" colspan="7" align="center"><font size="1" face="Arial, Helvetica, sans-serif">You have not entered any carts or all carts have been checked back in.</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="7" bgcolor="#c0c0c0"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsReturning.EOF
		'this is a RED warning
		if (rsReturning("idStatus") = 7 or rsReturning("idStatus") = 8)and datediff("d",rsReturning("dtReturn"),date) >= 10 then
			bgcolor = "#ff0000"
			fcolor="#ffff00"
		'this is a yellow warning
		elseif rsReturning("idStatus") = 7 and datediff("d",rsReturning("dtReturn"),date) >= 4 then
			bgcolor = "#ffff00"
			fcolor="#000000"
		else
			bgcolor = "#ffffff"
			fcolor = "#000000"
		end if
%>
                      <tr bgcolor="<%=bgcolor%>">
<%
	'check to see what kind of Order type this is going to be
	select case rsReturning("idType")
		case 1
			img = "circle.gif"
			alt = "Standard Cart"			
		case 2
			img = "square.gif"
			alt = "Internal Use Cart"
		case 3
			img = "animalprint.gif"
			alt = "Out of System"
	end select
	'check to see if there is a Show to Show and use lightning
	if cint(rsReturning("idShow2Show")) = 1 then
		img = "lightning.gif"
		alt = "Show to Show Cart"
	end if
%>
						<td height="20" align="center" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif"><IMG SRC="images/<%=img%>" alt="<%=alt%>"></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<A HREF="viewcart.asp?idCart=<%=rsReturning("idCart")%>"><%=trim(rsReturning("chrCart"))%></a></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=trim(rsReturning("chrCustomer"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=formatdatetime(rsReturning("dtDeparture"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=formatdatetime(rsReturning("dtReturn"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=trim(rsReturning("chrCartStatus"))%></font></td>
                        <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>"><%=rsReturning("intOrdered")%>&nbsp;</font></td>                        
                      </tr>
                      <tr> 
                        <td height="1" colspan="7" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif"><img src="images/c0c0c0dot.gif" width="1" height="1"></font></td>
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
</form>
</body>
</html>
<%
	rsStats.Close
	set rsStats = nothing
	rsNew.Close
	set rsNew = nothing
	rsReturning.Close
	set rsReturning = nothing
	dbConnection.Close
	set dbConnection = nothing
%>