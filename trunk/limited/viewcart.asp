<%@ Language=VBScript %>
<%
	if session("idUser") = "" then
		Response.Redirect "../logon.asp"
	end if
%>
<!-- #include file="../includes/openconn.asp" -->
<%	
	'Find order
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute FindCartbyID " & request("idCart")
	set rsCart = dbConnection.Execute(sql)
	
	'list all of the items for order
	set rsAssets = server.CreateObject("adodb.recordset")
	sql = "execute ListOrderedwithDescriptionsbyCart " & request("idCart")
	set rsAssets = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="../includes/title.htm" -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
    <!-- #include file="includes/top.htm" -->
    <tr> 
      <td width="10" background="images/leftverticalline.gif"><img src="images/leftverticalline.gif" width="10" height="10"></td>
      <td width="780">
		<table width="780" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td>
			  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td><font size="4" face="Arial, Helvetica, sans-serif"><strong>Comprehensive Order Information</strong></font></td>
                  <td align="right"><font size="2" face="Arial, Helvetica, sans-serif">&lt; <a href="default.asp">Return Home</a></font></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#6699cc"><img src="images/6699ccdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td>
              <table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">This page contains the important details of the order you have placed. &nbsp;If you require any changes, please notify your Pool Manager.</font></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td>
			  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td colspan="2" bgcolor="#6699cc"><font size="2" face="Arial, Helvetica, sans-serif" color="#ffffff"><strong>Cart Information</strong><font size="1"></font></font></td>
                </tr>
                <tr bgcolor="#f5f5f5"> 
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Cart Name<br>
                    <font size="2"><b><%=trim(rsCart("chrCart"))%>&nbsp;</b></font></font></td>
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Your Manager<br>
                    <font size="2"><b><%=trim(rsCart("chrManager"))%>&nbsp;</b></font></font></td>
                </tr>
                <tr bgcolor="#f5f5f5"> 
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Your Division / Department Name<br>
                    <font size="2"><b><%=trim(rsCart("chrDDName"))%>&nbsp;</b></font></font></td>
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Person to Receive Invoice<br>
                    <font size="2"><b><%=trim(rsCart("chrIPerson"))%>&nbsp;</b></font></font></td>
                </tr>
                <tr bgcolor="#f5f5f5"> 
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Your Division / Department Number<br>
                    <font size="2"><b><%=trim(rsCart("chrDDNumber"))%>&nbsp;</b></font></font></td>
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Email of Person Receiving Invoice<br>
                    <font size="2"><b><%=trim(rsCart("chrIEmail"))%>&nbsp;</b></font></font></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td colspan="2" bgcolor="#6699cc"><font color="#ffffff" size="2" face="Arial, Helvetica, sans-serif"><strong>Date Information</strong><font size="1"></font></font></td>
                </tr>
                <tr bgcolor="#f5f5f5"> 
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Delivery Date<br>
                    <font size="2"><b><%=formatdatetime(rsCart("dtArrival"),1)%>&nbsp;</b></font></font></td>
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">End Date (example: 02/01/2003)<br>
                    <font size="2"><b><%=formatdatetime(rsCart("dtDeparture"),1)%>&nbsp;</b></font></font></td>
                </tr>
                <tr bgcolor="#f5f5f5"> 
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Delivery Time, If relevant (example: 12:30 PM)<br>
                    <font size="2"><b><%=rsCart("dtArrivalTime")%>&nbsp;</b></font></font></td>
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">End Time, If relevant (example: 12:30 PM)<br>
				    <font size="2"><b><%=rsCart("dtDepartureTime")%>&nbsp;</b></font></font></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>	
          <tr> 
            <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td colspan="2" bgcolor="#6699cc"><font color="#ffffff" size="2" face="Arial, Helvetica, sans-serif"><strong>Shipping Information</strong><font size="1"></font></font></td>
                </tr>
                <tr bgcolor="#f5f5f5"> 
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">To (Booth / Receiving Party)<br>
                    <font size="2"><b><%=trim(rsCart("chrAddress"))%>&nbsp;</b></font></font></td>
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">City<br>
                    <font size="2"><b><%=trim(rsCart("chrCity"))%>&nbsp;</b></font></font></td>
                </tr>
                <tr bgcolor="#f5f5f5"> 
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Address (Company / Hotel / Venue)<br>
                    <font size="2"><b><%=trim(rsCart("chrAddress2"))%>&nbsp;</b></font></font></td>
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">State<br>
                    <font size="2"><b><%=trim(rsCart("chrState"))%>&nbsp;</b></font></font></td>
                </tr>
                <tr bgcolor="#f5f5f5"> 
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Address (PO Box / Street Address)<br>
                    <font size="2"><b><%=trim(rsCart("chrAddress3"))%>&nbsp;</b></font></font></td>
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Zip (example: 12345-1234)<br>
                    <font size="2"><b><%=trim(rsCart("chrZip"))%>&nbsp;</b></font></font></td>
                </tr>
                <tr bgcolor="#f5f5f5"> 
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Address (c/o or additional information)<br>
                    <font size="2"><b><%=trim(rsCart("chrAddress4"))%>&nbsp;</b></font></font></td>
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Country<br>
                    <font size="2"><b><%=trim(rsCart("chrCountry"))%>&nbsp;</b></font></font></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td>
			  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td colspan="2" bgcolor="#6699cc"><font color="#ffffff" size="2" face="Arial, Helvetica, sans-serif"><strong>On-Site Contact Information</strong><font size="1"></font></font></td>
                </tr>
                <tr bgcolor="#f5f5f5"> 
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Contact Name<br>
                    <font size="2"><b><%=trim(rsCart("chrOSPerson"))%>&nbsp;</b></font></font></td>
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Contact Cell Phone (exmaple: 408-555-1212)<br>
                    <font size="2"><b><%=trim(rsCart("chrOSPhone"))%>&nbsp;</b></font></font></td>
                </tr>
                <tr bgcolor="#f5f5f5"> 
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Contact Email<br>
                    <font size="2"><b><%=trim(rsCart("chrOSEmail"))%>&nbsp;</b></font></font></td>
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td>
			  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td bgcolor="#6699cc"><font color="#ffffff" size="2" face="Arial, Helvetica, sans-serif"><strong>Purpose of Equipment Loan</strong><font size="1"></font></font></td>
                </tr>
                <tr bgcolor="#f5f5f5"> 
                  <td><font size="2" face="Arial, Helvetica, sans-serif"><b><%=trim(rsCart("txtNotes"))%>&nbsp;</b></font></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td>
			  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td bgcolor="#6699cc"><font color="#ffffff" size="2" face="Arial, Helvetica, sans-serif"><strong>Setup Notes / Additional Information</strong></font></td>
                </tr>
                <tr bgcolor="#f5f5f5"> 
                  <td><font size="2" face="Arial, Helvetica, sans-serif"><b><%=trim(rsCart("txtShippingNotes"))%>&nbsp;</b></font></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
		  <tr> 
            <td>
			  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td colspan="2" bgcolor="#6699cc"><font color="#ffffff" size="2" face="Arial, Helvetica, sans-serif"><strong>Shipping Method</strong><font size="1"></font></font></td>
                </tr>
                <tr bgcolor="#f5f5f5"> 
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Carrier Name<br>
                    <font size="2"><b><%=trim(rsCart("chrCarrier"))%>&nbsp;</b></font></font></td>
                  <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Carrier Account Number<br>
                    <font size="2"><b><%=trim(rsCart("chrAccount"))%>&nbsp;</b></font></font></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td>
			  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td bgcolor="#6699cc" colspan="3"><font size="2" color="#ffffff" face="Arial, Helvetica, sans-serif"><strong>Equipment Requested</strong></font></td>
                </tr>
                <tr bgcolor="#f5f5f5">
                  <td><font size="2" face="Arial, Helvetica, sans-serif"><b>Item #</b></font></td>
                  <td><font size="2" face="Arial, Helvetica, sans-serif"><b>Asset Description</b></font></td>
                  <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><b>Requested Assets </b></font></td>
                </tr>
<%
	if rsAssets.EOF then
%>
                <tr bgcolor="#ffffff"> 
                  <td colspan="3" align="center"><font size="2" face="Arial, Helvetica, sans-serif">There are no assets in this Cart.</font></td>
                </tr>
<%
	else
		do until rsAssets.EOF
		if bgswitch = 1 then
			bgswitch = 0
			bgcolor = "#f5f5f5"
		else
			bgswitch = 1
			bgcolor = "#ffffff"
		end if
%>
                <tr bgcolor="<%=bgcolor%>">
                  <td><font size="2" face="Arial, Helvetica, sans-serif"><%=trim(rsAssets("chrItemNo"))%></font></td>
<%
			if rsAssets("chrType") = "C" then
%>
                  <td><font size="2" face="Arial, Helvetica, sans-serif"><%=trim(rsAssets("chrItem")) & " - " & trim(rsAssets("chrProcessor"))%><br>
                  <%=trim(rsAssets("chrMemory")) & " - " & trim(rsAssets("chrODrive"))%></font></td>
<%
			else
%>
                  <td><font size="2" face="Arial, Helvetica, sans-serif"><%=trim(rsAssets("chrItem"))%></font></td>
<%
			end if
%>
                  <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><%=rsAssets("intQuantity")%></font></td>
                </tr>
<%
		subtotal = subtotal + rsAssets("intQuantity")
		rsAssets.MoveNext
		loop
	end if
%>
                <tr bgcolor="#6b6b6b"> 
                  <td colspan="3" height="1"><img SRC="../images/6b6b6bdot.gif" WIDTH="1" HEIGHT="1"></td>
                </tr>
                <tr bgcolor="#ffffff"> 
                  <td colspan="2" align="right"><font size="2" face="Arial, Helvetica, sans-serif"><b>Total Assets Requested</b></font></td>
                  <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><b><%=subtotal%></b></font></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
        </table></td>
      <td width="10" background="images/rightverticalline.gif"><img src="images/rightverticalline.gif" width="10" height="10"></td>
    </tr>
    <!-- #include file="includes/bottom.htm" -->
  </table>
</body>
</html>
<%
	rsCart.Close
	set rsCart = nothing
	rsAssets.close
	set rsAssets = nothing
	dbConnection.Close
	set dbConnection = nothing
%>