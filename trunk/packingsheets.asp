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
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Packing Sheets for Corporate Events </strong></font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">Listed below are all of the packing content for each asset.</font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="100%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="1">
                            <tr> 
                              <td bgcolor="#c0c0c0"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                                  <tr> 
                                    <td colspan="2" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Order/Cart Reports</strong></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="iMacFP.pdf">iMac Flat Panel </a></font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="xServe.pdf">xServe</a></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="PowerMacG5.pdf">PowerMac G5</a> </font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="xServe-xRaid.pdf">xServe xRaid</a> </font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="PowerMacG4-White.pdf">PowerMac G4 - White</a> </font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="FibreChannelPCICard.pdf">Fibre Chanel PCI Card </a></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff"> 
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="PowerMacG4-Black.pdf">PowerMac G4 - Black</a> </font></td>
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="AirportBaseStation.pdf">Airport Basestation</a></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff">
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="iBookG4.pdf">iBook G4</a> </font></td>
                                    <td><font size="1" face="Arial, Helvetica, sans-serif"><a href="ZR45.pdf">Canon ZR45</a></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff">
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="PowerBookG4.pdf">PowerBook G4</a> </font></td>
                                    <td><font size="1" face="Arial, Helvetica, sans-serif"><a href="CanonPowerShotS400.pdf">Canon Power Shot S400</a> </font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff">
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                                    <td><font size="1" face="Arial, Helvetica, sans-serif"><a href="DVItoADCAdapter.pdf">DVI to ADC Adapter</a></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff">
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                                    <td><font size="1" face="Arial, Helvetica, sans-serif"><a href="iSight.pdf">iSight</a></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff">
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                                    <td><font size="1" face="Arial, Helvetica, sans-serif"><a href="eMagic.pdf">eMagic</a></font></td>
                                  </tr>
                                  <tr bgcolor="#ffffff">
                                    <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                                    <td><font size="1" face="Arial, Helvetica, sans-serif"><a href="SonicaTheaterUSB.pdf">Sonica Theatre USB </a></font></td>
                                  </tr>
                                </table></td>
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