<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on inventory button
	buttonswitch = 3
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Find the Asset Information
	set rsInventory = server.CreateObject("adodb.recordset")
	sql = "execute FindInventoryInfobyID " & request("idInventory")
	set rsInventory = dbConnection.Execute(sql)
	
	'Find the Description
	set rsDescription = server.CreateObject("adodb.recordset")
	sql = "execute ViewDescriptionwithCategory " & rsInventory("idDescription")
	set rsDescription = dbConnection.Execute(sql)
	
	'Find the FutureActivity
	set rsFuture = server.CreateObject("adodb.recordset")
	sql = "execute ListFutureAssetActivity " & request("idInventory")
	set rsFuture = dbConnection.Execute(sql)
	
	'Find the Past Activity
	set rsPast = server.CreateObject("adodb.recordset")
	sql = "execute ListPastAssetActivity " & request("idInventory")
	set rsPast = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="viewdescription.asp">
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
                <td>
                  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td rowspan="16" align="center"><img src="assetimages/<%=rsDescription("chrLocation")%>"></td>
                    <td><font size="3" face="Arial, Helvetica, sans-serif"><strong><%=trim(rsDescription("chrItem"))%></strong></font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Item # <%=trim(rsDescription("chrItemNo"))%></font></td>
                  </tr>
<%
	if rsDescription("chrType") = "C" then
%>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Processor: <%=trim(rsDescription("chrProcessor"))%></font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Memory: <%=trim(rsDescription("chrMemory"))%></font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Hard Drive: <%=trim(rsDescription("chrHDD"))%></font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Optical Drive: <%=trim(rsDescription("chrODrive"))%></font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Second Drive: <%=trim(rsDescription("chrRStorage"))%></font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">SCSI Device: <%=trim(rsDescription("chrSCSI"))%></font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Graphics Card: <%=trim(rsDescription("chrGraphics"))%></font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Wireless Card: <%=trim(rsDescription("chrWireless"))%></font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Bluetooth: <%=trim(rsDescription("chrBluetooth"))%></font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Modem: <%=trim(rsDescription("chrModem"))%></font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">USB: <%=trim(rsDescription("chrUSB"))%></font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">FireWire: <%=trim(rsDescription("chrFireWire"))%></font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Ethernet: <%=trim(rsDescription("chrEthernet"))%></font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Operating System: <%=trim(rsDescription("chrOS"))%></font></td>
                  </tr>
<%
	end if
%>
                </table>
			   </td>
              </tr>
              <tr>
                <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td bgcolor="#6699cc"><font color="#FFFFFF" size="3" face="Arial, Helvetica, sans-serif"><strong>Notes</strong></font></td>
                  </tr>
                  <tr>
                    <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsInventory("txtNotes"))%></font></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td bgcolor="#6699cc"><font color="#FFFFFF" size="3" face="Arial, Helvetica, sans-serif"><strong>Current Status</strong></font></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="1">
                  <tr>
                    <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr bgcolor="#c0c0c0">
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Asset #</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Serial Num #</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Date In</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Date Out</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Status</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Cart/Location</font></td>
                      </tr>
                      <tr bgcolor="#ffffff">
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrAssNum"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrSerialNum"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsInventory("dtIn"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=rsInventory("dtOut")%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrInventoryStatus"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrCart"))%></font></td>
                      </tr>
                    </table></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td bgcolor="#6699cc"><strong><font color="#FFFFFF" size="3" face="Arial, Helvetica, sans-serif">Future Activity</font></strong></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="1">
                  <tr>
                    <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr bgcolor="#c0c0c0">
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Shipping</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Arrival</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Departure</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Returning</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Cart</font></td>
                      </tr>
<%
	if rsFuture.EOF then
%>
                      <tr align="center" bgcolor="#ffffff">
                        <td height="20" colspan="5"><font size="1" face="Arial, Helvetica, sans-serif">This Asset has not been assigned to any Carts at this time</font></td>
                      </tr>
<%
	else
		do until rsFuture.EOF
%>
                      <tr bgcolor="#ffffff">
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsFuture("dtShip"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsFuture("dtArrival"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsFuture("dtDeparture"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsFuture("dtReturn"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsFuture("chrCart"))%></font></td>
                      </tr>
<%
		rsFuture.MoveNext
		loop
	end if
%>
                    </table></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td bgcolor="#6699cc"><font color="#FFFFFF" size="3" face="Arial, Helvetica, sans-serif"><strong>Past Activity</strong></font></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="1">
                  <tr>
                    <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr bgcolor="#c0c0c0">
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Date</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Action</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Cart/Location</font></td>
                      </tr>
<%
	if rsPast.EOF then
%>
                      <tr align="center" bgcolor="#ffffff">
                        <td height="20" colspan="3"><font size="1" face="Arial, Helvetica, sans-serif">This Asset has no activity to report</font></td>
                      </tr>
<%
	else
		do until rsPast.EOF
%>
                      <tr bgcolor="#ffffff">
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsPast("dtStamp"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsPast("chrInventoryStatus"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsPast("chrCart"))%></font></td>
                      </tr>
<%
		rsPast.MoveNext
		loop
	end if
%>
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
<%
	rsInventory.Close
	set rsInventory = nothing
	rsDescription.Close
	set rsDescription = nothing
	rsFuture.Close
	set rsFuture = nothing
	rsPast.Close
	set rsPast = nothing
	dbConnection.Close
	set dbConnection = nothing
%>