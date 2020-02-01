<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on order button
	buttonswitch = 2
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'find the cart information
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute ViewInvoicebyCart " & request("idCart")
	set rsCart = dbConnection.Execute(sql)
	
	'find the cart information
	set rsAssets = server.CreateObject("adodb.recordset")
	sql = "execute ListInvoiceAssetsbyCart " & request("idCart")
	set rsAssets = dbConnection.Execute(sql)	
	
	'find the cart information
	set rsWarehouse = server.CreateObject("adodb.recordset")
	sql = "execute FindWarehousebyID " & rsCart("idWarehouse")
	set rsWarehouse = dbConnection.Execute(sql)
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
		<!-- #include file="includes/orders-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top">
		<table width="625" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15"><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="610">
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Approve Cart: <%=trim(rsCart("chrCart"))%></strong></font></td>
                        <td align="right" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif"><a href="approvecart.asp?idCart=<%=request("idCart")%>">Summary View</a> | <a href="accountteam.asp">Need Help</a>?</font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="10"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="#c0c0c0">
					<table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr align="center"> 
                        <td><a href="approvethiscart.asp?idCart=<%=request("idCart")%>"><img src="images/approvecartcentered.gif" border="0" WIDTH="120" HEIGHT="19"></a></td>
						<td><a href="disapprovethiscart.asp?idCart=<%=request("idCart")%>"><img src="images/disapprovecartcentered.gif" border="0" WIDTH="120" HEIGHT="19"></a></td>
<%
	if session("idAccess") < "O" then
%>
                        <td><a href="editcart.asp?idCart=<%=request("idCart")%>"><img src="images/editcart.gif" border="0" WIDTH="120" HEIGHT="19"></a></td>
<%
	else
%>
						<td><a href="editcartlimited.asp?idCart=<%=request("idCart")%>"><img src="images/editcart.gif" border="0" WIDTH="120" HEIGHT="19"></a></td>
<%
	end if
%>
						<td><a href="approval/loadcart.asp?idCart=<%=request("idCart")%>"><img src="images/changeassetscentered.gif" border="0" WIDTH="120" HEIGHT="19"></a></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
				<tr> 
                  <td bgcolor="#c0c0c0">
					<table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr align="center"> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Customer / Pool</font></td>
						<td><font size="1" face="Arial, Helvetica, sans-serif">Created By</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Created On</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Current Status</font></td>
                      </tr>
                      <tr align="center" bgcolor="#ffffff"> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCustomer"))%></font></td>
						<td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrFirst")) & " " & trim(rsCart("chrLast"))%></font></td>						
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=formatdatetime(rsCart("dtStamp"),2)%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCartStatus"))%></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
				<tr> 
                  <td bgcolor="#c0c0c0">
					<table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Ship Date</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Begin Date</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">End Date</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Return Date</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">D/D Name</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">D/D #</font></td>
                      </tr>
                      <tr bgcolor="#ffffff"> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=formatdatetime(rsCart("dtShip"),2)%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=formatdatetime(rsCart("dtArrival"),2)%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=formatdatetime(rsCart("dtDeparture"),2)%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=formatdatetime(rsCart("dtReturn"),2)%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrDDName"))%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrDDNumber"))%></font></td>
                      </tr>
                      <tr valign="top" bgcolor="#ffffff"> 
                        <td colspan="2"> <p><font size="2" face="Arial, Helvetica, sans-serif"><strong>To:</strong></font></p>
                          <p><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrOSPerson"))%><br>
                            <%=trim(rsCart("chrAddress"))%><br>
                            <%=trim(rsCart("chrAddress2"))%><br>
                            <%=trim(rsCart("chrAddress3"))%><br>
                            <%=trim(rsCart("chrAddress4"))%><br>
                            <%=trim(rsCart("chrCity")) & ", " & trim(rsCart("chrState")) & " " & trim(rsCart("chrZip"))%><br>
                            <%=trim(rsCart("chrOSPhone"))%></font></p>
                            <font size="2" face="Arial, Helvetica, sans-serif"><strong>Shipper:</strong></font><br>
                            <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCarrier"))%> - <%=trim(rsCart("chrAccount"))%></td>
                        <td colspan="2"> <p><font size="2" face="Arial, Helvetica, sans-serif"><strong>Return:</strong></font></p>
                          <p><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsWarehouse("chrWarehouse"))%> Warehouse<br>
                            <%=trim(rsWarehouse("chrAddress"))%><br>
                            <%=trim(rsWarehouse("chrAddress2"))%><br>
                            <%=trim(rsWarehouse("chrCity")) & ", " & trim(rsWarehouse("chrState")) & " " & trim(rsWarehouse("chrZip"))%><br>
                            Phone: <%=trim(rsWarehouse("chrPhone"))%></font></p></td>
                        <td colspan="2"> <p><font size="2" face="Arial, Helvetica, sans-serif"><strong>Bill To:</strong></font></p>
                          <p><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCustomer"))%><br>
                            <%=trim(rsCart("chrBAddress"))%><br>
                            <%=trim(rsCart("chrBAddress2"))%><br>
                            <%=trim(rsCart("chrBCity")) & ", " & trim(rsCart("chrBState")) & " " & trim(rsCart("chrBZip"))%><br>
                            Phone: <%=trim(rsCart("chrBPhone"))%><br>
                            Fax: <%=trim(rsCart("chrBFax"))%></font></p></td>
                      </tr>
                    </table>
                  </td>
                </tr>
				<tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
				<tr> 
                  <td bgcolor="#c0c0c0">
					<table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr align="center"> 
                        <td align="left"><font size="1" face="Arial, Helvetica, sans-serif">Reason for Loan and Notes</font></td>
                      </tr>
                      <tr align="center" bgcolor="#ffffff"> 
                        <td align="left"><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("txtNotes"))%></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr> 
                        <td bgcolor="#c0c0c0">
						  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Asset #</font></td>
                              <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item Description</font></td>
                              <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Requested By</font></td>
                              <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Requested&nbsp;</font></td>
                            </tr>
<%
	if rsAssets.EOF then
%>
                            <tr align="center"> 
                              <td height="20" colspan="6" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif">No Assets have been requested at this time.</font></td>
                            </tr>
<%
	else
		do until rsAssets.EOF
			if bgswitch = 1 then
				bgcolor = "#f5f5f5"
				bgswitch = 0
			else
				bgcolor = "#ffffff"
				bgswitch = 1
			end if
%>
                            <tr bgcolor="<%=bgcolor%>"> 
                              <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrAssNum"))%></font></td>
<%
			if rsAssets("chrType") = "C" then
%>
                              <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrItem")) & " - " & trim(rsAssets("chrProcessor"))%><br>
                              &nbsp;<%=trim(rsAssets("chrMemory")) & " - " & trim(rsAssets("chrHDD")) & " - " & trim(rsAssets("chrODrive"))%></font></td>
<%
			else
%>
                              <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrItem"))%></font></td>
<%
			end if
%>
                              <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrFirst")) & " " & trim(rsAssets("chrLast"))%></font></td>
                              <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=rsAssets("intOrdered")%></font></td>
                            </tr>
<%
			'totals
			intOrdered = intOrdered + rsAssets("intOrdered")
			rsAssets.MoveNext
		loop
%>
                            <tr bgcolor="#c0c0c0">
                              <td height="20" colspan="3" align="right"><font size="1" face="Arial, Helvetica, sans-serif">Total Assets Requested&nbsp;</font></td>
                              <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=intOrdered%>&nbsp;</font></td>
                            </tr>
<%
	end if
%>
                          </table>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <!-- #Begin bottom part -->
    <!-- #include file="includes/bottom.htm" -->
  </table>
</body>
</html>
<%
	rsCart.Close
	set rsCart = nothing
	rsAssets.Close
	set rsAssets = nothing
	rsWarehouse.Close
	set rsWarehouse = nothing
	dbConnection.Close
	set dbConnection = nothing
%>