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
	'Find the Cart
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute FindCartbyID " & session("idLoadOut")
	set rsCart = dbConnection.Execute(sql)
	
	'Find the Asset
	set rsInventory = server.CreateObject("adodb.recordset")
	sql = "execute FindInventorybyID " & request("idInventory")
	set rsInventory = dbConnection.Execute(sql)	
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
                    <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Asset Removal</strong></font></td>
                  </tr>
                  <tr>
                    <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>You are about to remove the following asset from cart <%=trim(rsCart("chrCart"))%>.</strong></font><br>
                    <font color="#FF0000" size="2" face="Arial, Helvetica, sans-serif"><strong>This action is permanent and cannot be undone.</strong></font></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td align="center"><p><font color="#0000FF" size="3" face="Arial, Helvetica, sans-serif"><strong>Asset# <%=trim(rsInventory("chrAssNum"))%></strong></font></p>
                      <p><strong><font color="#FF0000" size="3" face="Arial, Helvetica, sans-serif">Proceed with removal?</font></strong></p>
                      <p><font size="3" face="Arial, Helvetica, sans-serif"><a href="removefromcart.asp?idCart=<%=request("idCart")%>&idInventory=<%=request("idInventory")%>">Yes</a> &nbsp;or&nbsp; <a href="finishedcheckout.asp">No</a></font></p></td>
                  </tr>
                </table></td>
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
</body>
</html>
<%
	rsCart.Close
	set rsCart = nothing
	dbConnection.Close
	set dbConnection = nothing
%>