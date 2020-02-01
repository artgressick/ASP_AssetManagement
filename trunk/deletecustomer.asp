<%@ Language=VBScript %>
<%
	'RRF - Confrimation page to Deleteing Customer
	
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on Settings button
	buttonswitch = 5
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Get the Customer Info
	set rsCustomer = server.CreateObject("adodb.recordset")
	sql = "execute FindCustomerbyID " & Request("idCustomer")
	set rsCustomer = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <form name="deleteCustomer" method="post" action="removecustomer.asp">
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
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Remove 
                          Confirmation</strong></font></td>
                      </tr>
                      <tr>
                        <td bgcolor="#f5f5f5"><font color="#FF0000" size="2" face="Arial, Helvetica, sans-serif"><strong>You 
                          are about to remove the following Customer. This cannot be undone.</strong></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td align="center"><table width="25%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td align="center">
                          <table width="200" border="0" cellspacing="0" Cellpadding="0">
                            <tr>
                              <td><font color="#000000" size="2" face="Arial, Helvetica, sans-serif"><STRONG><%=rsCustomer("chrCustomer")%></STRONG></font></td>
                            </tr>
                            <tr>
                              <td><font color="#000000" size="2" face="Arial, Helvetica, sans-serif"><STRONG><%=rsCustomer("chrCAddress")%></STRONG></font></td>
                            </tr>
                            <tr>
                              <td><font color="#000000" size="2" face="Arial, Helvetica, sans-serif"><STRONG><%=rsCustomer("chrCAddress2")%></STRONG></font></td>
                            </tr>
                              <td><font color="#000000" size="2" face="Arial, Helvetica, sans-serif"><STRONG><%=rsCustomer("chrCCity")%>,&nbsp;<%=rsCustomer("chrCState")%>.&nbsp;<%=rsCustomer("chrCZip")%></STRONG></font></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td align="center"> <input type="submit" name="Submit2" value="Cancel This Action">&nbsp;&nbsp;<input name="Submit1" type="submit" id="Submit1" value="Remove Item">
                          <input type="hidden" name="idCustomer" value="<%=Request("idCustomer")%>"> 
                        </td>
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
	rsCustomer.Close
	set rsCustomer = nothing
	dbConnection.Close
	set dbConnection = nothing
%>