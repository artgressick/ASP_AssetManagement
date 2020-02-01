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
	'find the order information
	set rsOrder = server.CreateObject("adodb.recordset")
	sql = "execute ViewOrderPagebyID " & request("idOrder")
	set rsOrder = dbConnection.Execute(sql)
	
	'List all of the Carts for this order.
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute ListCartsbyOrder " & request("idOrder")
	set rsCart = dbConnection.Execute(sql)
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
                  <td height="15" colspan="2"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="2">
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Order Detail: <%=trim(rsOrder("chrOrder"))%></strong></font></td>
                        <td width="50%" align="right" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif"><a href="accounteam.asp">Need Help</a>?</font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="1" colspan="2" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td height="15" colspan="2"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td width="50%">
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Order #: <%=rsOrder("idOrder")%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Status: <%=trim(rsOrder("chrOrderStatus"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Order Placed: <%=formatdatetime(rsOrder("dtStamp"),1)%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Entered by: <%=trim(rsOrder("chrFirst")) & " " & trim(rsOrder("chrLast"))%></font></td>
                      </tr>
                    </table>
                  </td>
                  <td width="50%">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="15" colspan="2"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td height="1" colspan="2" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td height="25" colspan="2"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	if not rsCart.EOF then
		do until rsCart.EOF
%>
                <tr> 
                  <td colspan="2">
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td nowrap bgcolor="#6699cc"><font color="#ffffff" size="2" face="Arial, Helvetica, sans-serif"><strong>Cart: <a href="viewcart.asp?idCart=<%=rsCart("idCart")%>" class="titlelink"><%=trim(rsCart("chrCart"))%></a></strong></font></td>
                        <td width="100%"><font size="1" face="Arial, Helvetica, sans-serif">&lt;&lt; Click on the Cart Name to view details.</font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
<%
		'get what is in the cart
		set rsInCart = server.CreateObject("adodb.recordset")
		sql = "execute ListOrderedwithDescriptionsbyCart " & rsCart("idCart")
		set rsInCart = dbConnection.Execute(sql)
%>
                <tr> 
                  <td colspan="2">
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr bgcolor="#6699cc"> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="#ffffff">&nbsp;Item #</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif" color="#ffffff">&nbsp;Item</font></td>
                        <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif" color="#ffffff">Qty.&nbsp;</font></td>
                      </tr>
<%
		'reset the bgswitch
		bgswitch = 1
		if rsInCart.EOF then
%>
                      <tr> 
                        <td height="20" colspan="3" align="center"><font size="1" face="Arial, Helvetica, sans-serif">No Assets have been added to this Cart</font></td>
                      </tr>
                      <tr bgcolor="#5b5b5b"> 
                        <td colspan="3"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
		else
			do until rsInCart.EOF
			if bgswitch = 1 then
				bgcolor = "#f5f5f5"
				bgswitch = 0
			else
				bgcolor = "#ffffff"
				bgswitch = 1
			end if
%>
                      <tr bgcolor="<%=bgcolor%>"> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInCart("chrItemNo"))%></font></td>
<%
			if rsInCart("chrType") = "C" then
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInCart("chrItem")) & " - " & trim(rsInCart("chrProcessor"))%><br>
                        &nbsp;<%=trim(rsInCart("chrMemory")) & " - " & trim(rsInCart("chrHDD")) & " - " & trim(rsInCart("chrODrive"))%></font></td>
<%
			else
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInCart("chrItem"))%></font></td>
<%
			end if
%>
                        <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsInCart("intQuantity"))%>&nbsp;</font></td>
                      </tr>
                      <tr bgcolor="#5b5b5b"> 
                        <td height="1" colspan="3"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
			rsInCart.MoveNext
			loop
		'close the connection
		rsInCart.Close
		set rsInCart = nothing
		end if
%>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="25" colspan="2"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
		rsCart.MoveNext
		loop
	end if
%>
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
	rsOrder.Close
	set rsOrder = nothing
	rsCart.Close
	set rsCart = nothing
	dbConnection.Close
	set dbConnection = nothing
%>