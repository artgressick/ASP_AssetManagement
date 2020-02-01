<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on order button
	buttonswitch = 2
	
	'First Line switch
	firstline = 0
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Get a list of the Customers
	'we need to pull a list of Customer that they can order from. Super user can order from anyone.
	set rsCustomers = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListCustomerNamesandIDs"
	else
		sql = "execute ListCustomerNamesandIDsbyManagerAccess " & session("idUser")
	end if
	set rsCustomers = dbConnection.Execute(sql)
	
	'Get a list of orders for the customers or all
	set rsOrders = server.CreateObject("adodb.recordset")
	if cint(request("idCustomer")) = 0 then
		if session("idAccess") < "O" then
			sql = "execute ListCartsforApprovalAll"
		else
			sql = "execute ListCartsforApprovalAllAccess " & session("idUser")
		end if
	else
		sql = "execute ListCartsforApprovalbyCustomer " & request("idCustomer")
	end if
	set rsOrders = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <form name="form1" method="post" action="approvecarts.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
		<!-- #include file="includes/orders-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="100%">
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="595" height="1"></td>
                </tr>
				<tr> 
                  <td>
				    <table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="50%"><strong><font size="3" face="Arial, Helvetica, sans-serif">Approve Carts</font></strong></td>
                        <td width="50%" align="right" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif"><a href="accountteam.asp">Need Help</a>?</font></td>
                      </tr>
                      <tr bgcolor="#f5f5f5"> 
                        <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">Listed below are all of the Carts that require approval. Please click on the Cart Name to begin the approval process.<br>
						You can decrease the number of carts listed by selecting the Customer / Pool from the drop-down list, and clicking on the Filter Carts button.</font></td>
                      </tr>
                    </table>
				  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
				<tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr> 
                        <td bgcolor="#5b5b5b">
						  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr bgcolor="#f5f5f5"> 
                              <td width="50%" align="right" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif">Customer / Pool</font></td>
                              <td align="center"> <font size="2" face="Arial, Helvetica, sans-serif"> 
                                <select name="idCustomer" size="1" id="idCustomer">
                                  <option value="0" <%if rsCustomers("idCustomer") = cint(request("idCustomer")) then%>selected<%end if%>>All Customers / Pools</option>
<%
		if not rsCustomers.eof then
			do until rsCustomers.EOF
%>
                                  <option value="<%=rsCustomers("idCustomer")%>" <%if rsCustomers("idCustomer") = cint(request("idCustomer")) then%>selected<%end if%>><%=trim(rsCustomers("chrCustomer"))%></option>
<%
			rsCustomers.MoveNext
			loop
		end if
%>
                                </select></font></td>
                              <td width="50%"><font size="2" face="Arial, Helvetica, sans-serif"><input type="submit" name="Submit" value="Filter Carts"></font></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
				<tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Cart Name</font></td>
						<td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Customer / Pool</font></td>
						<td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Begin Date</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;End Date</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Entered By</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Entered</font></td>
                        <td height="20" bgcolor="#6699CC" align="right"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Total Assets&nbsp;</font></td>
                      </tr>
<%
	if rsOrders.EOF then
%>
                      <tr> 
                        <td height="20" align="center" colspan="7"><font size="1" face="Arial, Helvetica, sans-serif">There are no Carts ready for Approval.</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="7" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsOrders.EOF
		if bgswitch = 1 then
			bgcolor = "#ffffff"
			bgswitch = 0
		else
			bgcolor = "#f5f5f5"
			bgswitch = 1
		end if
%>
                      <tr bgcolor="<%=bgcolor%>"> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<A HREF="approvecart.asp?idCart=<%=rsOrders("idCart")%>"><%=trim(rsOrders("chrCart"))%></a></font></td>
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrCustomer"))%></font></td>
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsOrders("dtArrival"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsOrders("dtDeparture"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrFirst")) & "&nbsp;" & trim(rsOrders("chrLast"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsOrders("dtStamp"),2)%></font></td>
                        <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=rsOrders("intOrdered")%>&nbsp;</font></td>
                      </tr>
					  <tr> 
                        <td colspan="7" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
		rsOrders.MoveNext
		loop
	end if
%>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table>
            </td>
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
          </tr>
        </table>
      </td>
    </tr>
    <!-- #Begin bottom part -->
    <!-- #include file="includes/bottom.htm" -->
  </table>
  </form>
</body>
</html>
<%
	rsOrders.Close
	set rsOrders = nothing
	rsCustomers.Close
	set rsCustomers = nothing
	dbConnection.Close
	set dbConnection = nothing
%>