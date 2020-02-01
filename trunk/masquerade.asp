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
	'List the Active Users
	set rsUsers = server.CreateObject("adodb.recordset")
	sql = "execute ListActiveUsers"
	set rsUsers = dbConnection.Execute(sql)	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Masquerade User</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="masqueradelaunch.asp">
  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="10" height="10" align="left" valign="top"><img src="images/topleftblue.gif" width="10" height="10"></td>
            <td width="780" height="10"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td><img src="images/eamlogo.gif" width="150" height="45"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                </tr>
              </table></td>
            <td width="10" height="10" align="right" valign="top"><img src="images/toprightblue.gif" width="10" height="10"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr>
            <td bgcolor="#FFFFFF">
			  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td height="35"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td align="center"><font color="#0000FF" size="3" face="Arial, Helvetica, sans-serif"><strong>Masquerade User</strong></font></td>
                </tr>
                <tr> 
                  <td><font size="1" face="Arial, Helvetica, sans-serif">Using this option you can log into the site as another person. Please be aware that you will need to log out and back in as a SuperUser.</font></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="15">
					  <tr>
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;User<BR>
						<select name="idUser" size="1">
<%
	if not rsUsers.eof then
		do until rsUsers.eof
%>
                          <option value="<%=rsUsers("idUser")%>"><%=trim(rsUsers("chrLast")) & ", " & trim(rsUsers("chrFirst"))%></option>
<%
		rsUsers.movenext
		loop
	end if
%>
                        </select></font></td>
					  </tr>
					  <tr>
						<td><input name="Cancel" type="submit" value="Log in as User"></td>
					  </tr>
					</table>
                  </td>
                </tr>
                <tr> 
                  <td height="50"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table>
            </td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="10" height="10" align="left" valign="bottom"><img src="images/bottomleftblue.gif" width="10" height="10"></td>
            <td width="780" height="10"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td align="center"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Copyright © 2003 
                    techIT Solutions LLC. <br>
                    Asset Management Enterprise Portal 4.0 &amp; Corporate Business 
                    Intelligence as products of techIT Solutions. </font></td>
                </tr>
              </table></td>
            <td width="10" height="10" align="right" valign="bottom"><img src="images/bottomrightblue.gif" width="10" height="10"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
	rsUsers.Close
	set rsUsers = nothing
	dbConnection.close
	set dbConnection = nothing
%>