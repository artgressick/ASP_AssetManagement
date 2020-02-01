<%@ Language=VBScript %>
<%
	'RRF - Confrimation page to Deleteing Team Member
	
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
	set rsAccount = server.CreateObject("adodb.recordset")
	sql = "Select * from tblAccounts Where idAccount=" & Request("idAccount")
	set rsAccount = dbConnection.Execute(sql)

    'Find User Info
	set rsUser = server.CreateObject("adodb.recordset")
	sql = "execute FindUserbyID " & rsAccount("idUser")
	set rsUser = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <form name="deleteAccount" method="post" action="removeaccount.asp">
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
                          are about to remove the following Team Member. This cannot be undone.</strong></font></td>
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
                              <td align="center"><font color="#000000" size="2" face="Arial, Helvetica, sans-serif"><STRONG><%=rsUser("chrFirst")%>&nbsp;<%=rsUser("chrLast")%></STRONG></font></td>
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
                        <td align="center"> <input type="submit" name="Submit1" value="Cancel This Action">&nbsp;&nbsp;<input name="Submit2" type="submit" id="Submit1" value="Remove Item">
                          <input type="hidden" name="idAccount" value="<%=Request("idAccount")%>"> 
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
	rsAccount.Close
	set rsAccount = nothing
	rsUser.Close
	set rsUser = nothing
	dbConnection.Close
	set dbConnection = nothing
%>