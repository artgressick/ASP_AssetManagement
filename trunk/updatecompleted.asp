<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on order button
	buttonswitch = 2
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'find out if there are any assets left in the cart
	set rsOrder = server.CreateObject("adodb.recordset")
	sql = "execute FindOrderbyID " & request("idOrder")
	set rsOrder = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="closeorder.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/orders-nav.htm" -->
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
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Close Order: <%=trim(rsOrder("chrOrder"))%></strong></font></td>
						<td align="right"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                      </tr>
                      <tr>
                        <td bgcolor="#f5f5f5" colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">Once this has been closed it cannot be re-opened.</font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><input name="Submit" type="submit" value="Close this Order">
                        <input type="hidden" name="idOrder" value="<%=request("idOrder")%>"></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr>
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
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
	rsOrder.Close
	set rsOrder = nothing
	dbConnection.Close
	set dbConnection = nothing
%>