<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on order button
	buttonswitch = 4
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="orders.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/reports-nav.htm" -->
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
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Reports</strong></font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">Listed below are all of the reports for Orders, Inventory, Predictive and Usage reports. We will be adding reports as they are	requested.</font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="300" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="1">
                            <tr> 
                              <td bgcolor="#c0c0c0"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                                  <tr> 
                                    <td colspan="2" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Order/Cart Reports</strong></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportshipping.asp">Carts Awaiting Shipping</a></font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportcartsreturning.asp">Returning Carts</a></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportexpedite.asp">Expedited Carts</a></font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportcustomization.asp">Customized Carts</a></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportlesa.asp">Order/Cart Information</a></font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportlesa2.asp">Order/Cart Monthly Report</a></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportloaner.asp">Loaner Agreement Orders</a></font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                                  </tr>
                                </table></td>
                            </tr>
                          </table></td>
                        <td width="10">&nbsp;</td>
                        <td width="300" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="1">
                            <tr> 
                              <td bgcolor="#c0c0c0"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                                  <tr> 
                                    <td colspan="2" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Inventory Reports</strong></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><A HREF="reportinventory.asp">Inventory Report</A></font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportbroken.asp">Assets Broken</a></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportassetsadded.asp">Assets Added</a></font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportassetsreturning.asp">Assets Returning</a></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportneedturning.asp">Need Turning</a></font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportlost.asp">Assets Lost</a></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportoutofsystem.asp">Out of System</a></font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportinternaluse.asp">Internal Use</a></font></td>
                                  </tr>
                                </table></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="300" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="1">
                          <tr>
                            <td bgcolor="#c0c0c0"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                                <tr>
                                  <td colspan="2" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Predictive Reports</strong></font></td>
                                </tr>
                                <tr>
                                  <td width="50%" bgcolor="#ffffff"><A HREF="reportwhatsavailable.asp"><font size="1" face="Arial, Helvetica, sans-serif">What's Available</font></A></td>
                                  <td width="50%" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif"><a href="reportfutureactivity.asp">Future Activity</a></font></td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                        </table>
                      </td>
                      <td width="10">&nbsp;</td>
                      <td width="300" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="1">
                          <tr>
                            <td bgcolor="#c0c0c0"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                                <tr>
                                  <td colspan="2" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Usage Reports</strong></font></td>
                                </tr>
                                <tr bgcolor="#ffffff">
                                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><A HREF="reportusage.asp">Customer Usage Report</A></font></td>
                                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"></font></td>
                                </tr>
                              </table>
                            </td>
                          </tr>
                        </table>
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