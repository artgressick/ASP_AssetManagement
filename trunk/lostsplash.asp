<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on inventory button
	buttonswitch = 3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="lostredirect.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/inventory-nav.htm" -->
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
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Lost Assets</strong></font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">This 
                          is the area for checking in and out lost assets. Please select from the 
                          below links to begin check out or check in.</font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="610" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="300" valign="top" bgcolor="#f5f5f5"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                            <tr>
                              <td align="center"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Check 
                                Out Lost Assets</strong></font></td>
                            </tr>
                            <tr>
                              <td height="25" align="center"><img src="images/f5f5f5dot.gif" width="1" height="1"></td>
                            </tr>
                            <tr>
                              <td align="center"><input type="submit" name="submit1" value="Begin Check Out"></td>
                            </tr>
                            <tr>
                              <td height="25" align="center"><img src="images/f5f5f5dot.gif" width="1" height="1"></td>
                            </tr>
                            <tr>
                              <td align="center"><font size="1" face="Arial, Helvetica, sans-serif">Please 
                                note that assets must be checked in from a Cart 
                                before you begin.</font></td>
                            </tr>
                          </table></td>
                        <td width="10">&nbsp;</td>
                        <td width="300" valign="top" bgcolor="#f5f5f5"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                            <tr>
                              <td align="center"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Check 
                                In Lost Assets</strong></font></td>
                            </tr>
                            <tr>
                              <td height="25" align="center"><img src="images/f5f5f5dot.gif" width="1" height="1"></td>
                            </tr>
                            <tr>
                              <td align="center"><input type="submit" name="submit2" value="Begin Check In"></td>
                            </tr>
                            <tr>
                              <td height="25" align="center"><img src="images/f5f5f5dot.gif" width="1" height="1"></td>
                            </tr>
                            <tr>
                              <td align="center"><font size="1" face="Arial, Helvetica, sans-serif">Please 
                                note that assets must have been checked out to loss status before you can check them back in.</font></td>
                            </tr>
                          </table></td>
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