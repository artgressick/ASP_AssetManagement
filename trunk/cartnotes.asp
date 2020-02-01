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
	'Get a list of the Open Orders
	set rsOrders = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListOpenOrders"
	else
		sql = "execute ListOpenOrdersbyAccess " & session("idUser")
	end if
	set rsOrders = dbConnection.Execute(sql)
	
	'get the list of Carts for the page
	if request("idOrder") = "" then
		nocarts = 0
	else
		nocarts = 1
		idOrder = cint(request("idOrder"))
		'get the records
		set rsCarts = server.CreateObject("adodb.recordset")
		sql = "execute ListCartsbyOrder " & idOrder
		set rsCarts = dbConnection.Execute(sql)
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="cartnotes.asp">
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
            <td width="610"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Retreive Cart</strong></font></td>
                        <td width="50%">&nbsp;</td>
                      </tr>
                      <tr bgcolor="#f5f5f5"> 
                        <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">Please select your order from the drop-down list, then click the Load Carts button.<br>
						The table below will list any carts in this order.</font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	if rsOrders.eof then
%>
                <tr> 
                  <td height="30"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td align="center"><font size="3" face="Arial, Helvetica, sans-serif" color="#ff0000"><b>There are no open orders at this time.</b></font></td>
                </tr>
<%
	else
%>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr> 
                        <td bgcolor="#5b5b5b">
						  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr bgcolor="#f5f5f5"> 
                              <td width="50%" align="right" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif">Open Orders</font></td>
                              <td align="center"> <font size="2" face="Arial, Helvetica, sans-serif"> 
                                <select name="idOrder" size="1" id="idOrder">
<%
		do until rsOrders.EOF
%>
                                  <option value="<%=rsOrders("idOrder")%>" <%if rsOrders("idOrder") = idOrder then%>selected<%end if%>><%=trim(rsOrders("chrOrder"))%></option>
<%
		rsOrders.MoveNext
		loop
%>
                                </select></font></td>
                              <td width="50%"><font size="2" face="Arial, Helvetica, sans-serif"><input type="submit" name="Submit" value="Load Carts"></font></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
		if nocarts = 1 then
%>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr bgcolor="#6699cc"> 
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Cart Name</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Entered By</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Status</font></td>
                        <td height="20" align="right"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Options&nbsp;</font></td>
                      </tr>
<%
			if rsCarts.EOF then
%>
                      <tr align="center"> 
                        <td height="20" colspan="4"><font size="1" face="Arial, Helvetica, sans-serif">There are no Carts for this Order.</font></td>
                      </tr>
                      <tr bgcolor="#5b5b5b"> 
                        <td height="1" colspan="4"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
			else
				do until rsCarts.EOF
%>
                      <tr> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsCarts("chrCart"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsCarts("chrFirst")) & " " & trim(rsCarts("chrLast"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsCarts("chrCartStatus"))%></font></td>
                        <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><a href="retrievenotes.asp?idCart=<%=rsCarts("idCart")%>&idOrder=<%=idOrder%>">Open Notes</A></font></td>
                      </tr>
                      <tr bgcolor="#5b5b5b"> 
                        <td height="1" colspan="4"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
				rsCarts.MoveNext
				loop
			end if
			rsCarts.Close
			set rsCarts = nothing
%>
                    </table>
                  </td>
                </tr>
<%
		end if
	end if
%>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
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
  </form>
</body>
</html>
<%
	rsOrders.Close
	set rsOrders = nothing
	dbConnection.Close
	set dbConnection = nothing
%>