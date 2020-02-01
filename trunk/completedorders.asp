<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on inventory button
	buttonswitch = 2
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'find the Category information
	set rsOrders = server.CreateObject("adodb.recordset")
	sql = "execute ListCompletedOrders"
	set rsOrders = dbConnection.Execute(sql)
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
		<!-- #include file="includes/orders-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="100%">
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="595" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr>
                        <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Close Order Process</strong></font></td>
                        <td width="50%" align="right" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr bgcolor="#f5f5f5">
						<td><font size="1" face="Arial, Helvetica, sans-serif">The orders listed below are Complete.<br>
						To close any of these orders, click on the cart name and follow the screen prompts.</font></td>
					  </tr>
					</table>
                  </td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Order/Cart</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Customer</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Ship Type</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Assets</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Reconfigured</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Special Load</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;D/D Info</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Bill To Email</font></td>
                      </tr>
<%
	if rsOrders.EOF then
%>
                      <tr> 
                        <td height="20" align="center" colspan="8"><font size="1" face="Arial, Helvetica, sans-serif">There are no completed orders at this time.</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="8" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsOrders.EOF
			if idOrder <> rsOrders("idOrder") then
				idOrder = rsOrders("idOrder")
%>
                      <tr> 
                        <td colspan="8" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
                      <tr> 
                        <td height="20" colspan="8" bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif"><A HREF="updatecompleted.asp?idOrder=<%=rsOrders("idOrder")%>"><%=trim(rsOrders("chrOrder"))%></A></font></td>
                      </tr>
<%
			end if
%>
                      <tr> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;# <%=rsOrders("idCart") & " - " & trim(rsOrders("chrCart"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrCustomer"))%></font></td>
<%
		'AEG - Severity of Cart Shipping
		select case rsOrders("idExpedite")
			case 0
				chrExpedite = "Normal"
			case 1
				chrExpedite = "Expedite"
			case 2
				chrExpedite = "Rush"
			case else
				chrExpedite = "Normal"
		end select
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=chrExpedite%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=rsOrders("intAssets")%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=rsOrders("intReconfigured")%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrSpecialLoad"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrDDName")) & "/" & trim(rsOrders("chrDDNumber"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrIEmail"))%></font></td>
                      </tr>
<%
		rsOrders.MoveNext
		loop
	end if
%>
                      <tr> 
                        <td colspan="8" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
          </tr>
        </table></td>
    </tr>
    <!-- #Begin bottom part -->
    <!-- #include file="includes/bottom.htm" -->
  </table>
</body>
</html>
<%
	rsOrders.Close
	set rsOrders = nothing
	dbConnection.Close
	set dbConnection = nothing
%>