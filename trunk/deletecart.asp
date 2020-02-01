<%@ Language=VBScript %>
<%
	'RRF - Confrimation page to Deleteing Carts
	
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on order button
	buttonswitch = 2
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Get the Cart
	set rsOrder = server.CreateObject("adodb.recordset")
	sql = "execute FindOrderbyID " & Request("idOrder")
	set rsOrder = dbConnection.Execute(sql)

	'Get the Cart
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute FindCartbyID " & Request("idCart")
	set rsCart = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="removecart2.asp">
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
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Remove 
                          Confirmation</strong></font></td>
                      </tr>
                      <tr>
                        <td bgcolor="#f5f5f5"><font color="#FF0000" size="2" face="Arial, Helvetica, sans-serif"><strong>You are about to remove the following Cart. This cannot be undone.</strong></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                         <tr> 
                        <td align="center"><font color="#000000" size="2" face="Arial, Helvetica, sans-serif"><STRONG><%=rsCart("chrCart")%></STRONG></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td align="center"><input type="submit" name="submit1" id="submit1" value="Cancel Removal">&nbsp;&nbsp;<input name="submit2" type="submit" id="submit2" value="Remove Cart">
                        <input type="hidden" name="idCart" value="<%=Request("idCart")%>"></td>
                        <input type="hidden" name="idStatus" value="<%=Request("idStatus")%>"></td>
                        <input type="hidden" name="idCustomer" value="<%=Request("idCustomer")%>"></td>
                        <input type="hidden" name="idUser" value="<%=Request("idUser")%>">
                        <input type="hidden" name="idType" value="<%=Request("idType")%>">
                        <input type="hidden" name="chrSearch" value="<%=Request("chrSearch")%>"></td>
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
  </form>
</body>
</html>
<%
	rsCart.Close
	set rsCart = nothing
	rsOrder.Close
	set rsOrder = nothing
	dbConnection.Close
	set dbConnection = nothing
%>