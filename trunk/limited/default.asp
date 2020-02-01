<%@ Language=VBScript %>
<%
	if session("idUser") = "" then
		Response.Redirect "../logoff.asp"
	end if
%>
<!-- #include file="../includes/openconn.asp" -->
<%	
	'get the upcoming orders
	set rsOrders = server.CreateObject("adodb.recordset")
	sql = "execute ListLimitedOrdersbyUser " & session("idUser")
	set rsOrders = dbConnection.Execute(sql)
	
	'Line starter flag
	linestarter = 0
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
    <td width="780"><table width="780" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="100%" height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr> 
                <td><font size="4" face="Arial, Helvetica, sans-serif"><strong>My Orders - Homepage</strong></font></td>
                <td align="right"><a HREF="whatsavailable.asp"><img SRC="images/whatsavailable.gif" border="0" WIDTH="103" HEIGHT="19"></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="addorderintro.asp"><img src="images/neworderbutton.gif" border="0" WIDTH="75" HEIGHT="18"></a></td>
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
          <td height="10"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">Listed 
                  below are all of the orders that you have placed. All new orders will remain OPEN until approved by the Administrator. If the Order has been declined you will receive an email and it 
                  will be removed from this list. If the order has been approved then the status will be changed to Approved. If you have any questions please contact the pool Administrator.</font></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td height="10"><img src="images/ffffffdot.gif" width="1" height="1"></td>
        </tr>
		<tr>
          <td height="10"><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr>
                <td width="50%" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif"><a HREF="../corp-pricing.pdf">Apple Corporate Events Pool Fees</a></font></td>
                <td width="50%" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif"><a href="packingsheets.asp">Packing Sheets for Corporate Events</a></font></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#6699cc"> 
                <td height="20"><strong><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Order #</font></strong></td>
                <td height="20"><strong><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Order/Cart</font></strong></td>
                <td height="20"><strong><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Ship</font></strong></td>
                <td height="20"><strong><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Return</font></strong></td>
                <td height="20"><strong><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Assets</font></strong></td>
                <td height="20"><strong><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Status</font></strong></td>
                <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
              </tr>
<%
	if rsOrders.EOF then
%>
              <tr> 
                <td height="20" colspan="7" align="center"><font size="1" face="Arial, Helvetica, sans-serif">You haven't placed any orders.</font></td>
              </tr>
              <tr bgcolor="#6699cc"> 
                <td colspan="7"><img src="images/6699ccdot.gif" width="1" height="1"></td>
              </tr>
<%
	else
		do until rsOrders.EOF
			if trim(rsOrders("chrOrder")) <> chrEventTemp then
				chrEventTemp = trim(rsOrders("chrOrder"))
				if linestarter = 0 then
					linestarter = 1
				else
%>
              <tr bgcolor="#6699cc"> 
                <td colspan="7"><img src="images/6699ccdot.gif" width="1" height="1"></td>
              </tr>
<%
				end if
%>
              <tr bgcolor="#f5f5f5"> 
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;#<%=rsOrders("idOrder")%></font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrOrder"))%></font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrOrderStatus"))%></font></td>
                <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%if rsOrders("idStatus") = 1 then%><a href="addcartintro.asp?idOrder=<%=rsOrders("idOrder")%>">Add New Cart</a><%end if%>&nbsp;</font></td>
              </tr>
              <tr bgcolor="#6699cc">
                <td colspan="7"><img src="images/6699ccdot.gif" width="1" height="1"></td>
              </tr>
              <tr bgcolor="#ffffff">
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<a HREF="viewcart.asp?idCart=<%=rsOrders("idCart")%>"><%=trim(rsOrders("chrCart"))%></a></font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsOrders("dtArrival"),2)%></font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsOrders("dtDeparture"),2)%></font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=rsOrders("intAssets")%></font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrCartStatus"))%></font></td>
                <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%if rsOrders("idCartStatus") = 1 then%><a href="../ordering/loadcart.asp?idCart=<%=rsOrders("idCart")%>">Add to Cart</a><%end if%>&nbsp;</font></td>
              </tr>
<%
			else
%>
              <tr bgcolor="#6699cc">
                <td colspan="7"><img src="images/6699ccdot.gif" width="1" height="1"></td>
              </tr>
              <tr bgcolor="#ffffff">
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<a HREF="viewcart.asp?idCart=<%=rsOrders("idCart")%>"><%=trim(rsOrders("chrCart"))%></a></font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsOrders("dtArrival"),2)%></font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsOrders("dtDeparture"),2)%></font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=rsOrders("intAssets")%></font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrCartStatus"))%></font></td>
                <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%if rsOrders("idCartStatus") = 1 then%><a href="../ordering/loadcart.asp?idCart=<%=rsOrders("idCart")%>">Add to Cart</a><%end if%>&nbsp;</font></td>
              </tr>
<%
			linestarter = 0
			end if
		rsOrders.MoveNext
		'background color change
		if bgswitch = 1 then
			bgswitch = 0
		else
			bgswitch = 1
		end if
		'loop through until no records
		loop
	end if
%>
              <tr bgcolor="#6699cc">
                <td colspan="7"><img src="images/6699ccdot.gif" width="1" height="1"></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
        </tr>
      </table>
    </td>
    <td width="10" background="images/rightverticalline.gif"><img src="images/rightverticalline.gif" width="10" height="10"></td>
  </tr>
  <!-- #include file="includes/bottom.htm" -->
</table>
</body>
</html>
<%
	rsOrders.Close
	set rsOrders = nothing
	dbConnection.Close
	set dbConnection = nothing
%>