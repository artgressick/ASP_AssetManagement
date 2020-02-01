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
                  <td height="15" colspan="3"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
				<tr> 
                  <td colspan="3">
				    <table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr bgcolor="#5b5b5b"> 
                        <td bgcolor="#5b5b5b"><font color="#ffffff" size="3" face="Arial, Helvetica, sans-serif"><strong>Cart Detail: <%=trim(rsCart("chrCart"))%></strong></font></td>
                      </tr>
                    </table>
				  </td>
                </tr>
				<tr> 
                  <td width="455" valign="top">
				   <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                    </tr>
                    <tr>
                      <td bgcolor="#c0c0c0"><table width="100%" border="0" cellspacing="1" cellpadding="3">
                        <tr align="center">
                          <td><font size="1" face="Arial, Helvetica, sans-serif">Customer</font></td>
                          <td><font size="1" face="Arial, Helvetica, sans-serif">Creator</font></td>
                          <td><font size="1" face="Arial, Helvetica, sans-serif">Created</font></td>
                          <td><font size="1" face="Arial, Helvetica, sans-serif">Status</font></td>
                        </tr>
                        <tr align="center" bgcolor="#ffffff">
                          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCustomer"))%></font></td>
                          <td><p><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrFirst")) & "&nbsp;" & trim(rsCart("chrLast"))%></font></p></td>
                          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=formatdatetime(rsCart("dtStamp"),2)%></font></td>
                          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCartStatus"))%></font></td>
                        </tr>
                      </table></td>
                    </tr>
                    <tr>
                      <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                    </tr>
                    <tr>
                      <td bgcolor="#c0c0c0"><table width="100%" border="0" cellspacing="1" cellpadding="3">
                        <tr align="center">
                          <td><font size="1" face="Arial, Helvetica, sans-serif">Last Updated</font></td>
                          <td><font size="1" face="Arial, Helvetica, sans-serif">Tech Support?</font></td>
                          <td><font size="1" face="Arial, Helvetica, sans-serif">Loaner Agreement</font></td>
                          <td><font size="1" face="Arial, Helvetica, sans-serif">Billed</font></td>
                          </tr>
                        <tr align="center" bgcolor="#ffffff">
                          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=rsCart("dtUpdated")%></font></td>
<%
	if rsCart("idSupport") = "True" then
		Support = "Yes"
	else
		Support = "No"
	end if
%>
                          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=Support%></font></td>
<%
	select case rsCart("idLoaner")
		case 0
			LoanerAgreement = "Not Required"
		case 1
			LoanerAgreement = "Sent"
		case 2
			LoanerAgreement = "Signed and Returned"
	end select
%>
                          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=LoanerAgreement%></font></td>
<%
	if rsCart("idBilled") = "True" then
		Billed = "Yes"
	else
		Billed = "No"
	end if
%>
                          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=Billed%></font></td>
                          </tr>
                      </table></td>
                    </tr>
                    <tr>
                      <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                    </tr>
                    <tr>
                      <td bgcolor="#c0c0c0"><table width="100%" border="0" cellspacing="1" cellpadding="3">
                        <tr align="center">
                          <td><font size="1" face="Arial, Helvetica, sans-serif">Order Type</font></td>
                          <td><font size="1" face="Arial, Helvetica, sans-serif">Apple Requestor</font></td>
                          </tr>
                        <tr align="center">
<%
	select case rsCart("idType")
		case 1
			OrderType = "Standard"
		case 2
			OrderType = "Internal Use"
		case 3
			OrderType = "Out of System"
	end select
%>
                          <td bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif"><%=OrderType%></font></td>
                          <td bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrARName"))%>&nbsp;(<%=trim(rsCart("chrAREmail"))%>)</font></td>
                          </tr>
                      </table></td>
                    </tr>
                  </table></td>
                  <td width="15">&nbsp;</td>
                  <td width="140" valign="top" bgcolor="#5b5b5b" align="center">
				  	<table width="100%" border="0" cellspacing="0" cellpadding="0">
<%
	if session("idAccess") < "O" then
%>
                    <tr>
                      <td height="29" align="center"><A HREF="editcart.asp?idCart=<%=rsCart("idCart")%>"><img src="images/editcart.gif" width="120" height="19" border="0"></a></td>
                    </tr>
<%
	'this is for the pool manager
	elseif session("idAccess") = "O" and cint(rsCart("idStatus")) < 4 then
%>
					<tr>
                      <td height="29" align="center"><A HREF="editcartlimited.asp?idCart=<%=rsCart("idCart")%>"><img src="images/editcart.gif" width="120" height="19" border="0"></a></td>
                    </tr>
<%
	'this is for the user
	elseif rsCart("idUser") = cint(session("idUser")) and cint(rsCart("idStatus")) < 3 then
%>
                    <tr>
                      <td height="29" align="center"><A HREF="editcartlimited.asp?idCart=<%=rsCart("idCart")%>"><img src="images/editcart.gif" width="120" height="19" border="0"></a></td>
                    </tr>
<%
	end if
%>
					<tr>
                      <td height="29" align="center"><A HREF="loanagreement.asp?idCart=<%=rsCart("idCart")%>"><img src="images/loaneragreement.gif" width="120" height="19" border="0"></a></td>
                    </tr>
<%
	if session("idAccess") < "O" then
%>
                    <tr>
                      <td height="29" align="center"><A HREF="approval/loadcart.asp?idCart=<%=rsCart("idCart")%>"><img src="images/changeassets.gif" width="120" height="19" border="0"></a></td>
                    </tr>
<%
	'this is for the pool manager
	elseif session("idAccess") = "O" and cint(rsCart("idStatus")) < 4 then
%>
					<tr>
                      <td height="29" align="center"><A HREF="approval/loadcart.asp?idCart=<%=rsCart("idCart")%>"><img src="images/changeassets.gif" width="120" height="19" border="0"></a></td>
                    </tr>
<%
	'this is for the user
	elseif rsCart("idUser") = cint(session("idUser")) and cint(rsCart("idStatus")) < 3 then
%>
					<tr>
                      <td height="29" align="center"><A HREF="ordering/loadcart.asp?idCart=<%=rsCart("idCart")%>"><img src="images/changeassets.gif" width="120" height="19" border="0"></a></td>
                    </tr>
<%
	end if
	if session("idAccess") < "O" then
%>
                    <tr>
                      <td height="29" align="center"><A HREF="unlockcart.asp?idCart=<%=rsCart("idCart")%>"><img src="images/unlockcart.gif" width="120" height="19" border="0"></a></td>
                    </tr>
<%
	'this is for the pool manager
	elseif session("idAccess") = "O" and cint(rsCart("idStatus")) < 4 then
%>
					<tr>
                      <td height="29" align="center"><A HREF="unlockcart.asp?idCart=<%=rsCart("idCart")%>"><img src="images/unlockcart.gif" width="120" height="19" border="0"></a></td>
                    </tr>
<%
	end if
%>
                    <tr>
                      <td height="29" align="center"><a href="shippinginvoice.asp?idCart=<%=request("idCart")%>" target="_blank"><img src="images/shippingmanifest.gif" width="120" height="19" border="0"></a></td>
                    </tr>
                  </table></td>
                </tr>
                <tr> 
                  <td height="20" colspan="3"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="3">
					<table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr> 
                        <td bgcolor="#c0c0c0">
						  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Asset #</font></td>
                              <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item Description</font></td>
                              <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Requested by</font></td>
                              <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Requested&nbsp;</font></td>
                              <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Shipped&nbsp;</font></td>
                              <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Returned&nbsp;</font></td>
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
                              <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=formatdatetime(rsAssets("dtOrdered"),2)%></font></td>
<%
			'Only format the date if there is a date
			if len(rsAssets("dtShipped")) <> 0 then
				dtShipped = formatdatetime(rsAssets("dtShipped"),2)
			end if
			'only format the date if there is a date
			if len(rsAssets("dtReturned")) <> 0 then
				dtReturned = formatdatetime(rsAssets("dtReturned"),2)
			end if
%>
                              <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=dtShipped%></font></td>
                              <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=dtReturned%></font></td>
                            </tr>
<%
			
			'Clear the dtShipped and dtReturned
			dtShipped = ""
			dtReturned = ""
			'totals
			intOrdered = intOrdered + rsAssets("intOrdered")
			intShipped = intShipped + rsAssets("intShipped")
			intReturned = intReturned + rsAssets("intReturned")
			rsAssets.MoveNext
		loop
%>
                            <tr bgcolor="#c0c0c0">
                              <td height="20" colspan="3" align="right"><font size="1" face="Arial, Helvetica, sans-serif">Totals&nbsp;</font></td>
                              <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=intOrdered%>&nbsp;</font></td>
                              <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=intShipped%>&nbsp;</font></td>
                              <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=intReturned%>&nbsp;</font></td>
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
                  <td height="20" colspan="3"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                  <td bgcolor="#c0c0c0" colspan="3">
					<table width="100%" border="0" cellspacing="1" cellpadding="3">
					  <tr>
					    <td><font size="1" face="Arial, Helvetica, sans-serif">Purpose of Loan</font></td>
					  </tr>
					  <tr bgcolor="ffffff">
					    <td><font size="1" face="Arial, Helvetica, sans-serif"><%=rsCart("txtNotes")%></font></td>
					  </tr>
					</table>
				  </td>
                <tr>
                </tr>
                <tr> 
                  <td height="20" colspan="3"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                  <td bgcolor="#c0c0c0" colspan="3">
					<table width="100%" border="0" cellspacing="1" cellpadding="3">
					  <tr>
					    <td><font size="1" face="Arial, Helvetica, sans-serif">Notes and Additional Information</font></td>
					  </tr>
					  <tr bgcolor="ffffff">
					    <td><font size="1" face="Arial, Helvetica, sans-serif"><%=rsCart("txtShippingNotes")%></font></td>
					  </tr>
					</table>
				  </td>
                <tr>
                </tr>
                <tr> 
                  <td height="20" colspan="3"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="#c0c0c0" colspan="3">
					<table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Ship Date</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Carrier</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Account</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">End Date</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Return VIA</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">D/D Name</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">D/D #</font></td>
                      </tr>
                      <tr bgcolor="#ffffff"> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=formatdatetime(rsCart("dtShip"),2)%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCarrier"))%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrAccount"))%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=formatdatetime(rsCart("dtDeparture"),2)%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">FedEx 2-Day</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrDDName"))%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrDDNumber"))%></font></td>
                      </tr>
                      <tr valign="top" bgcolor="#ffffff"> 
                        <td colspan="3"><font size="2" face="Arial, Helvetica, sans-serif"><strong>To:</strong></font><br>
                          <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrOSPerson"))%><br>
                            <%=trim(rsCart("chrAddress"))%><br>
                            <%=trim(rsCart("chrAddress2"))%><br>
                            <%=trim(rsCart("chrAddress3"))%><br>
                            <%=trim(rsCart("chrAddress4"))%><br>
                            <%=trim(rsCart("chrCity")) & ", " & trim(rsCart("chrState")) & " " & trim(rsCart("chrZip"))%><br>
                            <%=trim(rsCart("chrOSPhone"))%></font></td>
                        <td colspan="2"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Return:</strong></font><br>
                          <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsWarehouse("chrWarehouse"))%> Warehouse<br>
                            <%=trim(rsWarehouse("chrAddress"))%><br>
                            <%=trim(rsWarehouse("chrAddress2"))%><br>
                            <%=trim(rsWarehouse("chrCity")) & ", " & trim(rsWarehouse("chrState")) & " " & trim(rsWarehouse("chrZip"))%><br>
                            Phone: <%=trim(rsWarehouse("chrPhone"))%></font></td>
                        <td colspan="2"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Bill To:</strong></font><br>
                          <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCustomer"))%><br>
                            <%=trim(rsCart("chrBAddress"))%><br>
                            <%=trim(rsCart("chrBAddress2"))%><br>
                            <%=trim(rsCart("chrBCity")) & ", " & trim(rsCart("chrBState")) & " " & trim(rsCart("chrBZip"))%><br>
                            Phone: <%=trim(rsCart("chrBPhone"))%><br>
                            Fax: <%=trim(rsCart("chrBFax"))%></font></td>
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