<%@ Language=VBScript %>
<%
	session("idUser") = ""
	session.Abandon

	'RRF Temporary Fix
	if request("idURL") = "default.asp" then
		Response.Redirect "logon.asp"
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title.htm" -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#ffffff" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form action="loadaccount.asp" method="post" name="logon" id="logon">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td align="center" valign="middle">
        <table width="620" height="350" border="0" cellpadding="0" cellspacing="0">
          <!-- #include file="includes/logon-top.htm" -->
          <tr>
            <td width="10" height="330" background="images/leftblack.gif"><img src="images/leftblack.gif" width="10" height="10"></td>
            <td width="600" valign="top" bgcolor="#f5f5f5"><table width="600" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="25"><img src="images/f5f5f5dot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="600" border="0" cellspacing="0" cellpadding="2">
                      <tr> 
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Logged Off</strong></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/f5f5f5dot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="600" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="2" face="Arial, Helvetica, sans-serif">You 
                          are now logged off the Asset Management System. &nbsp;Possible reasons this has occurred may be:</font></td>
                      </tr>
					  <tr> 
                        <td><font size="2" face="Arial, Helvetica, sans-serif"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&gt;&gt;</strong>&nbsp;&nbsp;You 
                         clicked on the logoff link. 
                          </font></td>
                      </tr>
                      <tr> 
                        <td><font size="2" face="Arial, Helvetica, sans-serif"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&gt;&gt;</strong>&nbsp;&nbsp;The link
						between our webserver and your browser has been broken. 
                          </font></td>
                      </tr>
                      <tr> 
                        <td><font size="2" face="Arial, Helvetica, sans-serif"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&gt;&gt;</strong>&nbsp;&nbsp;You 
                          have stayed on the same page for longer than 20 minutes.</font></td>
                      </tr>
                      <tr>
                        <td height="20"><img src="images/f5f5f5dot.gif" width="1" height="1"></td>
                      </tr>
                      <tr>
                        <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><a href="logon.asp">Please click 
                          here to log in again</a>.<BR><BR>
                          <font color="#ffffff">Link = <%=Request.ServerVariables("URL")%>?<%=Request.ServerVariables("QUERY_STRING")%></font></font></td>
                      </tr>
                    </table> </td>
                </tr>
              </table></td>
            <td width="10" height="330" background="images/rightblack.gif"><img src="images/rightblack.gif" width="10" height="10"></td>
          </tr>
          <!-- #include file="includes/logon-bottom.htm" -->
        </table></td>
    </tr>
  </table>
</form>
</body>
</html>
