<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "../logoff.asp"
	end if
%>
<!-- #include file="../includes/openconn.asp" -->
<%
	'Find the Cart information
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute FindCartbyID " & session("idCart")
	set rsCart = dbConnection.Execute(sql)
	
	'Get a list of the Categories
	set rsCategories = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListCategories"
	else
		sql = "execute ListCategoriesbyAccess " & session("idUser")
	end if
	set rsCategories = dbConnection.Execute(sql)
	
	'prime the descriptions
	if request("idCategory") = "" then
		idCategory = rsCategories("idCategory")
	else
		idCategory = request("idCategory")
	end if
	
	'Get the descriptions for the category
	set rsDescriptions = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListDescriptionswithWhatsAvailableCount " & idCategory & "," & session("idCart") & "," & rsCart("idCustomer")
	else
		sql = "execute ListDescriptionswithWhatsAvailableCount3 " & idCategory & "," & session("idCart") & "," & session("idUser")
	end if
	set rsDescriptions = dbConnection.Execute(sql)
	
	'List what is in the cart
	set rsInCart = server.CreateObject("adodb.recordset")
	sql = "execute ListCartwithDescriptionsbyID " & session("idCart")
	set rsInCart = dbConnection.Execute(sql)
	
	'subtotal primer
	subtotal = 0
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="default.asp">
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
    <!-- #include file="includes/top.htm" -->
    <tr> 
      <td bgcolor="#6699cc">
		<table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr bgcolor="#ffffff"> 
            <td width="600" valign="top" bgcolor="#ffffff">
			  <table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Cart Name: <%=trim(rsCart("chrCart"))%></strong></font></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Ship to:</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Ships: <%=formatdatetime(rsCart("dtShip"),1)%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrAddress"))%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Arrives: <%=formatdatetime(rsCart("dtArrival"),1)%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrAddress2"))%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Returns: <%=formatdatetime(rsCart("dtReturn"),1)%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrAddress3"))%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Onsite Contact: <%=trim(rsCart("chrOSPerson"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCity"))%>, <%=trim(rsCart("chrState"))%> <%=trim(rsCart("chrZip"))%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Onsite Number: <%=trim(rsCart("chrOSPhone"))%></font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="../images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr> 
                        <td bgcolor="#5b5b5b">
						  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr bgcolor="#f5f5f5"> 
                              <td width="50%" align="right" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif">Category</font></td>
                              <td><font size="2" face="Arial, Helvetica, sans-serif"> 
                                <select name="idCategory" size="1" id="idCategory">
<%
	if not rsCategories.EOF then
		do until rsCategories.EOF
%>
                                  <option value="<%=rsCategories("idCategory")%>" <%if rsCategories("idCategory") = cint(idCategory) then%>selected<%end if%>><%=trim(rsCategories("chrCategory"))%></option>
<%
		rsCategories.MoveNext
		loop
	end if
%>
                                </select></font></td>
                              <td width="50%" align="left" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><input type="submit" name="Submit" value="Submit"></font></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="20" bgcolor="#6699cc"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item #</font></td>
                        <td height="20" bgcolor="#6699cc"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item Description</font></td>
                        <td height="20" align="center" bgcolor="#6699cc"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Ordered</font></td>
                        <td height="20" align="center" bgcolor="#6699cc"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Available</font></td>
                        <td height="20" align="right" bgcolor="#6699cc"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Options&nbsp;</font></td>
                      </tr>
<%
	if rsDescriptions.EOF then
%>
                      <tr align="center"> 
                        <td height="20" colspan="5"><font size="1" face="Arial, Helvetica, sans-serif">There are no Assets available under this Category</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="5" bgcolor="#5b5b5b"><img src="../images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsDescriptions.EOF
		if bgswitch = 1 then
			bgcolor = "#ffffff"
			bgswitch = 0
		else
			bgcolor = "#f5f5f5"
			bgswitch = 1
		end if
%>
                      <tr> 
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;#<%=rsDescriptions("chrItemNo")%></font></td>
<%
		if rsDescriptions("chrType") = "C" then
%>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsDescriptions("chrItem")) & " - " & trim(rsDescriptions("chrProcessor"))%><br>
                        &nbsp;<%=trim(rsDescriptions("chrMemory")) & " - " & trim(rsDescriptions("chrODrive"))%></font></td>
<%
		else
%>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsDescriptions("chrItem"))%></font></td>
<%
		end if
%>
                        <td height="20" align="center" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif"><%=rsDescriptions("intCart")%></font></td>
                        <td height="20" align="center" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif"><%=rsDescriptions("intAvailable")%></font></td>
                        <td height="20" align="right" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif"><a href="addtocart.asp?idCategory=<%=idCategory%>&idDescription=<%=rsDescriptions("idDescription")%>">Add</a> - <a href="removefromcart.asp?idCategory=<%=idCategory%>&idDescription=<%=rsDescriptions("idDescription")%>">Remove</a>&nbsp;</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="5" bgcolor="#5b5b5b"><img src="../images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
		rsDescriptions.MoveNext
		loop
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
            <td width="200" valign="top" bgcolor="#f5f5f5">
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr align="center" bgcolor="#5b5b5b"> 
                  <td height="35" colspan="2"><A HREF="../orders.asp" class="titlelink"><img src="../images/exitshoppingcart.gif" width="120" height="19" border="0"></a></td>
                </tr>
                <tr align="center" bgcolor="#5b5b5b"> 
                  <td height="35" colspan="2"><a href="checkout.asp"><img src="../images/reviewandsubmit.gif" width="120" height="19" border="0"></a></td>
                </tr>
                <tr> 
                  <td height="25" colspan="2" align="center"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Request Summary</strong></font></td>
                </tr>
                <tr bgcolor="#6699cc"> 
                  <td height="20" align="center" nowrap><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Qty&nbsp;</font></td>
                  <td width="100%" height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item Description</font></td>
                </tr>
<%
	're-prime the bgcolor switch
	bgswitch = 0
	if rsInCart.EOF then
%>
                <tr bgcolor="#ffffff"> 
                  <td height="20" align="center" nowrap colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">Cart is Empty</font></td>
                </tr>
<%
	else
		do until rsInCart.EOF
		if bgswitch = 1 then
			bgswitch = 0
			bgcolor = "#f5f5f5"
		else
			bgswitch = 1
			bgcolor = "#ffffff"
		end if
%>
                <tr bgcolor="<%=bgcolor%>"> 
                  <td height="20" align="center" nowrap><font size="1" face="Arial, Helvetica, sans-serif"><%=rsInCart("intQuantity")%></font></td>
<%
		if rsInCart("chrType") = "C" then
%>
                  <td width="100%" height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<A HREF="default.asp?idCategory=<%=rsInCart("idCategory")%>"><%=trim(rsInCart("chrItem"))%></A><BR>&nbsp;<%=trim(rsInCart("chrProcessor"))%></font></td>
<%
		else
%>
                  <td width="100%" height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<A HREF="default.asp?idCategory=<%=rsInCart("idCategory")%>"><%=trim(rsInCart("chrItem"))%></a></font></td>
<%
		end if
%>
                </tr>
<%
		'add up the quantity
		subtotal = subtotal + rsInCart("intQuantity")
		rsInCart.MoveNext
		loop
	end if
%>
                <tr bgcolor="#6699cc"> 
                  <td height="20" align="center" nowrap><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif"><%=subtotal%></font></td>
                  <td width="100%" height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Total Assets Requested</font></td>
                </tr>
                <tr> 
                  <td height="20" colspan="2"><img src="../images/f5f5f5dot.gif" width="1" height="1"></td>
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
	rsCategories.Close
	set rsCategories = nothing
	rsDescriptions.Close
	set rsDescriptions = nothing
	rsCart.Close
	set rsCart = nothing
	rsInCart.Close
	set rsInCart = nothing
	dbConnection.Close
	set dbConnection = nothing
%>