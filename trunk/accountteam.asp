<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on order button
	buttonswitch = 1
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
		<!-- #include file="includes/home-nav.htm" -->
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
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Account 
                          Team</strong></font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">
                          The following techIT staff members are always available to assist you when needed.<br>
                           <font color="#ff0000">Please be sure to contact your Pool Manager first before contacting techIT Solutions.</font></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="300" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td bgcolor="#f5f5f5"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Account 
                                Managers</strong></font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Robert 
                                Kite (<a href="mailto:rkite@techitsolutions.com">rkite@techitsolutions.com</a>)</font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Mike 
                                Tyson (<a href="mailto:mtyson@techitsolutions.com">mtyson@techitsolutions.com</a>)</font></td>
                            </tr>
                            <tr> 
                              <td height="25"><font face="Arial, Helvetica, sans-serif"><img src="images/ffffffdot.gif" width="1" height="1"></font></td>
                            </tr>
                            <tr> 
                              <td bgcolor="#f5f5f5"><font face="Arial, Helvetica, sans-serif"><strong>Warehouse 
                                Operations</strong></font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Livermore 
                                California (<a href="mailto:warheouse@techitsolutions.com">warehouse@techitsolutions.com</a>)</font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Knoxville 
                                Tennessee (<a href="mailto:tnwarehouse@techitsolutions.com">tnwarehouse@techitsolutions.com</a>)</font></td>
                            </tr>
                          </table></td>
                        <td width="10">&nbsp;</td>
                        <td width="300" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td bgcolor="#f5f5f5"><font face="Arial, Helvetica, sans-serif"><strong>Billing 
                                &amp; Invoicing</strong></font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Lesa 
                                Preston (<a href="mailto:lpreston@techitsolutions.com">lpreston@techitsolutions.com</a>)</font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Peggy 
                                Neill (<a href="mailto:peggy@techitsolutions.com">peggy@techitsolutions.com</a>)</font></td>
                            </tr>
                            <tr> 
                              <td height="25"><font face="Arial, Helvetica, sans-serif"><img src="images/ffffffdot.gif" width="1" height="1"></font></td>
                            </tr>
                            <tr> 
                              <td bgcolor="#f5f5f5"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Web 
                                Site Problems</strong></font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif">John 
                                Gessman (<a href="mailto:jgessman@techitsolutions.com">jgessman@techitsolutions.com</a>)</font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Arthur 
                                Gressick (<a href="mailto:agressick@techitsolutions.com">agressick@techitsolutions.com</a>)</font></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
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