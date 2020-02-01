<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "../logoff.asp"
	end if
%>
<!-- #include file="../includes/openconn.asp" -->
<%
	'List what is in the cart
	set rsInCart = server.CreateObject("adodb.recordset")
	sql = "execute ListCheckOutbyCartID " & session("idCart")
	set rsInCart = dbConnection.Execute(sql)
	
	'Find the Cart information
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute FindCartbyID " & session("idCart")
	set rsCart = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="updatecart.asp">
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
    <!-- #include file="includes/top.htm" -->
    <tr> 
      <td bgcolor="#6699cc">
		<table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr bgcolor="#ffffff"> 
            <td valign="top" bgcolor="#ffffff">
              <table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Review &amp; Submit: <%=trim(rsCart("chrCart"))%></strong></font></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">Listed below is a summary of the Assets you have requested. &nbsp; 
                          Please ensure your cart is accurate before submitting it to the Pool Manager.<br>
						  Once you submit your cart, it will be locked and no assets can be added or removed except by the Pool Manager or techIT Solutions Account Manager.<br> 
						  If you decide to exit the shopping cart, all information will be saved until you Review and Submit the cart.<br>
						  If you have any questions please contact your Pool Manager or your techIT Solutions Account Manager.</font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="../images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr bgcolor="#6699cc"> 
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item #</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item Description</font></td>
                        <td height="20" align="center"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Available</font></td>
                        <td height="20" align="center"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Ordered</font></td>
                        <td height="20" align="right"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Options&nbsp;</font></td>
                      </tr>
<%
	if rsInCart.EOF then
%>
                      <tr> 
                        <td height="20" colspan="5" align="center"><font size="1" face="Arial, Helvetica, sans-serif">Your cart is empty. &nbsp;Please click the "Back" button in your browser, add assets, and resubmit your cart.</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="5" bgcolor="#5b5b5b"><img src="../images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsInCart.EOF
%>
                      <tr> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInCart("chrItemNo"))%></font></td>
<%
		if rsInCart("chrType") = "C" then
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInCart("chrItem")) & " - " & trim(rsInCart("chrProcessor"))%><BR>
                        &nbsp;<%=trim(rsInCart("chrMemory")) & " - " & trim(rsInCart("chrODrive"))%></font></td>
<%
		else
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInCart("chrItem"))%></font></td>
<%
		end if
%>
                        <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=rsInCart("intAvailable")%></font></td>
                        <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=rsInCart("intQuantity")%></font></td>
                        <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><a href="addtocart.asp?idDescription=<%=rsInCart("idDescription")%>">Add</a> - <a href="removefromcart.asp?idDescription=<%=rsInCart("idDescription")%>">Remove</a>&nbsp;</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="5" bgcolor="#5b5b5b"><img src="../images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
			subtotal = subtotal + rsInCart("intQuantity")
			rsInCart.MoveNext
		loop
%>
                      <tr> 
                        <td height="30" colspan="3" align="right" bgcolor="#6699cc"><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Total Assets Requested&nbsp;</font></strong></td>
                        <td height="30" align="center" bgcolor="#6699cc"><strong><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><%=subtotal%></font></strong></td>
                        <td height="30" align="center"><input type="submit" name="Submit" value="Submit Cart"></td>
                      </tr>
<%
	end if
%>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="25"><img src="../images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <!-- #include file="includes/bottom.htm" -->
  </table>
</form>
</body>
</html>
<%
	rsInCart.Close
	set rsInCart = nothing
	rsCart.Close
	set rsCart = nothing
	dbConnection.Close
	set dbConnection = nothing
%>
