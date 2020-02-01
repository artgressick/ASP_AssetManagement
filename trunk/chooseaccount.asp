<%@ Language=VBScript %>
<%
	'Check to make sure that the user is connected to the Server
	'if session("idUser") = "" then
	'	Response.Redirect "logoff.asp"
	'end if
%>
<!-- #include file="includes/openconn.asp" -->
<%	
	'Get all of the Access levels for the User
	set rsAccounts = server.CreateObject("adodb.recordset")
	sql = "execute ListUserAccess " & session("idUser")
	set rsAccounts = dbConnection.Execute(sql)
	
	'set all of the flag to good	
	errflag = 0
	selectedflag = 0
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title.htm" -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#ffffff" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form action="loadaccount.asp" method="post" name="logon" id="logon">
<input type="hidden" name="idUrl" value="<%=request("idUrl")%>">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td align="center" valign="middle"><table width="620" height="350" border="0" cellpadding="0" cellspacing="0">
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
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Access Level</strong></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td bgcolor="#5b5b5b"><table width="600" border="0" cellspacing="1" cellpadding="3">
                      <tr>
                        <td bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif">Listed below are all of the access levels that have been assigned to your profile.<br>
						Select the radio button to the left of the access level you wish to use for this session, then click the Enter button.</font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td height="25"><img src="images/f5f5f5dot.gif" width="1" height="1"></td>
                </tr>
                <tr>
                  <td><table width="600" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="20"><img src="images/f5f5f5dot.gif" width="1" height="1"></td>
                        <td height="20"><strong><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;Access Level</font></strong></td>
						<td height="20"><strong><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;Description</font></strong></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="3" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
	if rsAccounts.EOF then
		errflag = 1
%>
                      <tr> 
                        <td height="20" align="center" bgcolor="#ffffff" colspan="3"><font size="1" face="Arial, Helvetica, sans-serif">You do not have access to any accounts at this time.</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="3" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsAccounts.EOF
%>
                      <tr> 
                        <td height="20" align="center" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif"><input type="radio" name="idAccess" value="<%=rsAccounts("idAccess")%>" <%if selectedflag = 0 then%>checked<%end if%>></font></td>
                        <td height="20" bgcolor="#ffffff"><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAccounts("chrAccess"))%></font></td>
						<td height="20" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAccounts("chrAbility"))%></font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="3" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
		if selectedflag = 0 then
			selectedflag = 1
		end if
		rsAccounts.MoveNext
		loop
	end if
%>
                    </table></td>
                </tr>
                <tr>
                  <td height="25"><img src="images/f5f5f5dot.gif" width="1" height="1"></td>
                </tr>
<%
	if errflag <> 1 then
%>
                <tr>
                  <td><table width="600" border="0" cellspacing="0" cellpadding="2">
                      <tr>
                        <td><input type="submit" name="Submit" value="Enter -&gt;"></td>
                      </tr>
                    </table></td>
                </tr>
<%
	end if
%>
                <tr>
                  <td height="25"><img src="images/f5f5f5dot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
            <td width="10" height="330" background="images/rightblack.gif"><img src="images/rightblack.gif" width="10" height="10"></td>
          </tr>
          <!-- #include file="includes/logon-bottom.htm" -->
        </table>
      </td>
    </tr>
  </table>
</form>
</body>
</html>
<%
	'close all of the database connections
	rsAccounts.Close
	set rsAccounts = nothing
	dbConnection.Close
	set dbConnection = nothing
%>