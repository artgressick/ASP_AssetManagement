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
	'Get a list of orders that need approval
	set rsNew = server.CreateObject("adodb.recordset")
	sql = "execute ListCartsPreArrivalbyUser " & session("idUser")
	set rsNew = dbConnection.Execute(sql)
	
	'get a list of orders that will return in 7 days
	set rsReturning = server.CreateObject("adodb.recordset")
	sql = "execute ListCartsReturningbyUser " & session("idUser")
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
                  <td>
				  	<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Pool User Dashboard</strong></font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">Welcome to the Pool User Dashboard.
                        Listed below are all of the orders that you have placed in the system that are still open. The screen is split
                        into two sections. One for Carts that have not been shipped and the other is for returning.</font></td>
                      </tr>
                    </table>
				  </td>
                </tr>
				<tr> 
                  <td height="10"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
				  	<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif"><a href="corp-pricing.pdf">Corporate Events Pool Fees</a></font></td>
						<td bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif"><a href="packingsheets.asp">Asset Packing Sheets</a></font></td>
                      </tr>
                    </table>
				  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr bgcolor="#6699cc"> 
                        <td height="25" colspan="6"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong>&nbsp;New/Approved Carts</strong></font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
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
                        <td height="20" colspan="6"><font size="1" face="Arial, Helvetica, sans-serif">You have not entered any carts or all carts are returning.</font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="6"><font size="1" face="Arial, Helvetica, sans-serif"><img src="images/c0c0c0dot.gif" width="1" height="1"></font></td>
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
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<A HREF="viewcart.asp?idCart=<%=rsNew("idCart")%>"><%=trim(rsNew("chrCart"))%></A></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=trim(rsNew("chrCustomer"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=formatdatetime(rsNew("dtShip"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=formatdatetime(rsNew("dtArrival"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=trim(rsNew("chrCartStatus"))%></font></td>
                        <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=rsNew("intOrdered")%></font></td>
                      </tr>
					  <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="6"><font size="1" face="Arial, Helvetica, sans-serif"><img src="images/c0c0c0dot.gif" width="1" height="1"></font></td>
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
                        <td height="25" colspan="7" bgcolor="#6699cc">&nbsp;<font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong>Carts Returning</strong></font></td>
                      </tr>
                      <tr> 
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
                        <td height="20" colspan="6" align="center"><font size="1" face="Arial, Helvetica, sans-serif">You have not entered any carts or all carts have been checked back in.</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="6" bgcolor="#c0c0c0"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
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
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<A HREF="viewcart.asp?idCart=<%=rsReturning("idCart")%>"><%=trim(rsReturning("chrCart"))%></a></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=trim(rsReturning("chrCustomer"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=formatdatetime(rsReturning("dtDeparture"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=formatdatetime(rsReturning("dtReturn"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>">&nbsp;<%=trim(rsReturning("chrCartStatus"))%></font></td>
                        <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif" color="<%=fcolor%>"><%=rsReturning("intOrdered")%>&nbsp;</font></td>                        
                      </tr>
                      <tr> 
                        <td height="1" colspan="6" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif"><img src="images/c0c0c0dot.gif" width="1" height="1"></font></td>
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
	rsNew.Close
	set rsNew = nothing
	rsReturning.Close
	set rsReturning = nothing
	dbConnection.Close
	set dbConnection = nothing
%>