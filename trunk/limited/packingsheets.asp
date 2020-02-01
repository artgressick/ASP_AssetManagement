<%@ Language=VBScript %>
<%
	if session("idUser") = "" then
		Response.Redirect "../logon.asp"
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="../includes/title.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" onLoad="document.form1.chrOrder.focus()">
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
    <!-- #include file="includes/top.htm" -->
    <tr> 
      <td width="10" background="images/leftverticalline.gif"><img src="images/leftverticalline.gif" width="10" height="10"></td>
      <td width="780"><table width="780" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td width="50%"><strong><font size="3" face="Arial, Helvetica, sans-serif">Packing Sheets for Corporate Events</font></strong></td>
                  <td width="50%" align="right" valign="bottom"><font size="2" face="Arial, Helvetica, sans-serif">&lt; <a href="default.asp">Return Home</a></font></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td bgcolor="#6699cc"><img src="images/6699ccdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">Listed below are all of the packing content for each asset.</font></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr bgcolor="#ffffff">
                <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="../iMacFP.pdf">iMac Flat Panel </a></font></td>
                <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="../xServe.pdf">xServe</a></font></td>
              </tr>
              <tr bgcolor="#ffffff">
                <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="../PowerMacG5.pdf">PowerMac G5</a> </font></td>
                <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="../xServe-xRaid.pdf">xServe xRaid</a> </font></td>
              </tr>
              <tr bgcolor="#ffffff">
                <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="../PowerMacG4-White.pdf">PowerMac G4 - White</a> </font></td>
                <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="../FibreChannelPCICard.pdf">Fibre Chanel PCI Card </a></font></td>
              </tr>
              <tr bgcolor="#ffffff">
                <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="../PowerMacG4-Black.pdf">PowerMac G4 - Black</a> </font></td>
                <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="../AirportBaseStation.pdf">Airport Basestation</a></font></td>
              </tr>
              <tr bgcolor="#ffffff">
                <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="../iBookG4.pdf">iBook G4</a> </font></td>
                <td><font size="1" face="Arial, Helvetica, sans-serif"><a href="../ZR45.pdf">Canon ZR45</a></font></td>
              </tr>
              <tr bgcolor="#ffffff">
                <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><a href="../PowerBookG4.pdf">PowerBook G4</a> </font></td>
                <td><font size="1" face="Arial, Helvetica, sans-serif"><a href="../CanonPowerShotS400.pdf">Canon Power Shot S400</a> </font></td>
              </tr>
              <tr bgcolor="#ffffff">
                <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                <td><font size="1" face="Arial, Helvetica, sans-serif"><a href="../DVItoADCAdapter.pdf">DVI to ADC Adapter</a></font></td>
              </tr>
              <tr bgcolor="#ffffff">
                <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                <td><font size="1" face="Arial, Helvetica, sans-serif"><a href="../iSight.pdf">iSight</a></font></td>
              </tr>
              <tr bgcolor="#ffffff">
                <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                <td><font size="1" face="Arial, Helvetica, sans-serif"><a href="../eMagic.pdf">eMagic</a></font></td>
              </tr>
              <tr bgcolor="#ffffff">
                <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                <td><font size="1" face="Arial, Helvetica, sans-serif"><a href="../SonicaTheaterUSB.pdf">Sonica Theatre USB </a></font></td>
              </tr>
            </table></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
          </tr>
        </table></td>
      <td width="10" background="images/rightverticalline.gif"><img src="images/rightverticalline.gif" width="10" height="10"></td>
    </tr>
    <!-- #include file="includes/bottom.htm" -->
  </table>
</body>
</html>