<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on order button
	buttonswitch = 2
	
	'First Line switch for the orders and carts
	firstline = 0
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Get a list of the Customers
	set rsCustomers = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListCustomerNamesandIDs"
	else
		sql = "execute ListCustomerNamesandIDsbyAccess " & session("idUser")
	end if
	set rsCustomers = dbConnection.Execute(sql)
	
	'Get a list of the Orders
	set rsOrderStatus = server.CreateObject("adodb.recordset")
	sql = "execute ListOrderStatus"
	set rsOrderStatus = dbConnection.Execute(sql)
	
	'Get a list of users for the customer
	set rsUsers = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListUsersWhoEnteredanOrder"
	else
		sql = "execute ListUsersWhoEnteredanOrderbyAccess " & session("idUser")
	end if
	set rsUsers = dbConnection.Execute(sql)
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
		<!-- #include file="includes/orders-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="100%">
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="595" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr>
                        <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Orders</strong></font></td>
                        <td width="50%" align="right" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif"><a href="accountteam.asp">Need Help</a>?</font></td>
                      </tr>
                      <tr>
                        <td colspan="2" width="100%" bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">Please enter your search parameters - then click the Find Orders / Carts button.</font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="595" height="1"></td>
                </tr>
                <tr>
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr>
                        <td bgcolor="#5b5b5b">
						  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr bgcolor="#f5f5f5"> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Customers<br>
                                <select name="idCustomer" size="1" id="idCustomer">
                                  <option value="0" selected>All Customers</option>
<%
	if not rsCustomers.EOF then
		do until rsCustomers.EOF
%>
                                  <option value="<%=rsCustomers("idCustomer")%>" <%if cint(request("idCustomer")) = rsCustomers("idCustomer") then%>selected<%end if%>><%=trim(rsCustomers("chrCustomer"))%></option>
<%
		rsCustomers.MoveNext
		loop
	end if
%>
                                </select></font></td>
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Status<br>
                                <select name="idStatus" size="1" id="idStatus">
                                  <option value="0" selected>All</option>
<%
	if not rsOrderStatus.EOF then
		do until rsOrderStatus.EOF
%>
                                  <option value="<%=rsOrderStatus("idStatus")%>" <%if cint(request("idStatus")) = rsOrderStatus("idStatus") then%>selected<%end if%>><%=trim(rsOrderStatus("chrOrderStatus"))%></option>
<%
		rsOrderStatus.MoveNext
		loop
	end if
%>
                                </select></font></td>
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Entered By<br>
                                <select name="idUser" size="1" id="idUser">
                                  <option value="0">All Users</option>
<%
	if not rsUsers.EOF then
		do until rsUsers.EOF
%>
                                  <option value="<%=rsUsers("idUser")%>" <%if cint(request("idUser")) = rsUsers("idUser") then%>selected<%end if%>><%=trim(rsUsers("chrLast")) & ", " & trim(rsUsers("chrFirst"))%></option>
<%
		rsUsers.MoveNext
		loop
	end if
%>
                                </select></font></td>
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Order Types<br>
                                <select name="idType" size="1" id="idType">
                                  <option value="0" <%if cint(request("idType")) = 0 then%>selected<%end if%>>All Types</option>
                                  <option value="1" <%if cint(request("idType")) = 1 then%>selected<%end if%>>Regular</option>
                                  <option value="2" <%if cint(request("idType")) = 2 then%>selected<%end if%>>Internal Use</option>
                                  <option value="3" <%if cint(request("idType")) = 3 then%>selected<%end if%>>Out of System</option>
                                </select></font></td>
                            </tr>
                            <tr bgcolor="#f5f5f5"> 
                              <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">Order / Cart Name<br>
                                <input name="chrSearch" type="text" id="chrSearch" size="45" maxlength="100" value="<%=request("chrSearch")%>"> (Blank means all)</font></td>
                              <td colspan="2" valign="bottom" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><input type="submit" name="Submit" value="Find Orders / Carts"></font></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
<%
	'check the page numbering
	if request("idPage") = "" then
		idPage = 1
		intBegin = 1
		intEnd = 30
	else
		idPage = cint(request("idPage"))
		intEnd = idPage*30
		intBegin = intEnd-29
	end if
%>
                      <tr> 
                        <td height="20" bgcolor="#ffffff" colspan="9"><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Page <%=idPage%></font></td>
                      </tr>
                      <tr> 
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Order #</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Order/Cart Name</font></td>
                        <td bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Arrival</font></td>
                        <td bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Departure</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Customer</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Entered By</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;# Assets</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Status</font></td>
                        <td height="20" bgcolor="#6699CC" align="right"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Options&nbsp;</font></td>
                      </tr>
<%
	if request("idStatus") <> "" then
		'Seperator and changing the text to work
		chrSearch = replace(request("chrSearch"),"'","''")
		chrSearch = replace(chrSearch," ",",")
		idSeperator = ","
		'Get a list of the Orders
		set rsOrders = server.CreateObject("adodb.recordset")
		if session("idAccess") < "O" then
			noedit = 1 'can edit order
			sql = "execute ListOrdersPagewithPageCounter " & intBegin & "," & intEnd & "," & request("idCustomer") & "," & request("idStatus") & "," & request("idUser") & "," & request("idType") & ",'" & chrSearch & "','" & idSeperator & "'"
		else
			noedit = 1 'can edit order
			sql = "execute ListOrdersPagebyAccesswithPageCounter " & intBegin & "," & intEnd & "," & request("idCustomer") & "," & request("idStatus") & "," & request("idUser") & "," & request("idType") & "," & session("idUser") & ",'" & chrSearch & "','" & idSeperator & "'"
		end if
		set rsOrders = dbConnection.Execute(sql)
		if rsOrders.EOF then
%>
                      <tr> 
                        <td height="20" align="center" colspan="9"><font size="1" face="Arial, Helvetica, sans-serif">No orders with this search criteria.</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="9" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
		else
			'figure out how many pages there are
			intTotal = rsOrders("intTotal")
			intPages = round((intTotal/30)+.5,0)
			'----------------------------------------
			do until rsOrders.EOF
			'display the header of the order if differnet
			if rsOrders("idOrder") <> tempOrder then
				tempOrder = rsOrders("idOrder")
				if firstline = 1 then
%>
                      <tr> 
                        <td colspan="9" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
				else
					'past the first line now
					firstline = 1
				end if
%>
                      <tr> 
                        <td height="20" bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;#<%=rsOrders("idOrder")%></font></td>
                        <td height="20" bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<A HREF="vieworder.asp?idOrder=<%=rsOrders("idOrder")%>"><%=trim(rsOrders("chrOrder"))%></A></font></td>
                        <td height="20" bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td height="20" bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td height="20" bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrCustomer"))%></font></td>
                        <td height="20" bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td height="20" bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td height="20" bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrOrderStatus"))%></font></td>
                        <td height="20" bgcolor="#f5f5f5" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%if session("idAccess") = "A" then%><A HREF="editorder.asp?idOrder=<%=rsOrders("idOrder")%>&idCustomer=<%=request("idCustomer")%>&idStatus=<%=request("idStatus")%>&idUser=<%=request("idUser")%>&idType=<%=request("idType")%>&chrSearch=<%=trim(request("chrSearch"))%>">Edit</A> - <A HREF="deleteorder.asp?idOrder=<%=rsOrders("idOrder")%>&idCustomer=<%=request("idCustomer")%>&idStatus=<%=request("idStatus")%>&idUser=<%=request("idUser")%>&idType=<%=request("idType")%>&chrSearch=<%=trim(request("chrSearch"))%>">Remove</A><%end if%>&nbsp;</font></td>
                      </tr>
                      <tr> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<A HREF="viewcart.asp?idCart=<%=rsOrders("idCart")%>"><%=trim(rsOrders("chrCart"))%></A></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsOrders("dtArrival"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsOrders("dtDeparture"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrFirst")) & "&nbsp;" & trim(rsOrders("chrLast"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=rsOrders("intOrdered")%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrCartStatus"))%></font></td>
<%
	'This is for the administrator
	if session("idAccess") < "O" then
%>
                        <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><A HREF="deletecart.asp?idCart=<%=rsOrders("idCart")%>&idOrder=<%=rsOrders("idOrder")%>&idCustomer=<%=request("idCustomer")%>&idStatus=<%=request("idStatus")%>&idUser=<%=request("idUser")%>&idType=<%=request("idType")%>&chrSearch=<%=trim(request("chrSearch"))%>">Remove</A>&nbsp;</font></td>
<%
	'cannot do anything
	else
%>
                        <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
<%
	end if
%>
                      </tr>
<%
			else
			'display another cart
%>
                      <tr> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<A HREF="viewcart.asp?idCart=<%=rsOrders("idCart")%>"><%=trim(rsOrders("chrCart"))%></A></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsOrders("dtArrival"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsOrders("dtDeparture"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrFirst")) & "&nbsp;" & trim(rsOrders("chrLast"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=rsOrders("intOrdered")%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsOrders("chrCartStatus"))%></font></td>
<%
	'This is for the administrator
	if session("idAccess") < "O" then
%>
                        <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><A HREF="deletecart.asp?idCart=<%=rsOrders("idCart")%>&idOrder=<%=rsOrders("idOrder")%>&idCustomer=<%=request("idCustomer")%>&idStatus=<%=request("idStatus")%>&idUser=<%=request("idUser")%>&idType=<%=request("idType")%>&chrSearch=<%=trim(request("chrSearch"))%>">Remove</A>&nbsp;</font></td>
<%
	'cannot do anything
	else
%>
                        <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
<%
	end if
%>
                      </tr>
<%
			end if
			rsOrders.MoveNext
			loop
		end if
		rsOrders.Close
		set rsOrders = nothing
%>
                      <tr> 
                        <td colspan="9" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
                      <tr> 
                        <td colspan="9">
						  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td align="center"><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Page Index:&nbsp;
<%
		for i = 1 to intPages
%>
                              <A HREF="orders.asp?idPage=<%=i%>&idCustomer=<%=request("idCustomer")%>&idStatus=<%=request("idStatus")%>&idUser=<%=request("idUser")%>&idType=<%=request("idType")%>&chrSearch=<%=request("chrSearch")%>"><%=i%></A>&nbsp;
<%
		next
%>
                              </font></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
<%
	end if
%>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table>
            </td>
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
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
	rsCustomers.Close
	set rsCustomers = nothing
	rsUsers.close
	set rsUsers = nothing
	rsOrderStatus.Close
	set rsOrderStatus = nothing
	dbConnection.Close
	set dbConnection = nothing
%>