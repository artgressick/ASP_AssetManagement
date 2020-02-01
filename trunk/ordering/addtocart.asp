<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "../logoff.asp"
	end if
%>
<!-- #include file="../includes/openconn.asp" -->
<%
	'Get a list of the Categories
	set rsAssets = server.CreateObject("adodb.recordset")
	sql = "execute CountAvailableforDescriptionID " & request("idDescription") & "," & session("idCart")
	set rsAssets = dbConnection.Execute(sql)
	
	'List what is in the cart
	set rsInCart = server.CreateObject("adodb.recordset")
	sql = "execute ListCartwithDescriptionsbyID " & session("idCart")
	set rsInCart = dbConnection.Execute(sql)
	
	'Find the Description
	set rsDescription = server.CreateObject("adodb.recordset")
	sql = "execute ViewDescriptionwithCategory " & request("idDescription")
	set rsDescription = dbConnection.Execute(sql)
	
	'subtotal primer
	subtotal = 0
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="uploadcart.asp">
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
    <!-- #include file="includes/top.htm" -->
    <tr> 
      <td bgcolor="#6699cc">
		<table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr bgcolor="#ffffff"> 
            <td width="600" valign="top" bgcolor="#ffffff">
			  <table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Add to Cart</strong></font></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">To add assets to this cart, please select the number of assets and click the Update Cart button.<br>
						Your information will be saved and the cart will be updated.</font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="../images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="5">
                      <tr> 
                        <td width="50%" align="center"><img src="../assetimages/<%=rsDescription("chrLocation")%>"></td>
                        <td width="50%">
						  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr> 
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
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="../images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	if rsAssets("intAvailable") = 0 then
%>
                <tr>
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr>
                        <td bgcolor="#5b5b5b">
						  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr>
                              <td align="center" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>No <%=trim(rsDescription("chrItem"))%>'s are available at this time.</strong></font></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                  </td>
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
                            <tr>
                              <td align="center" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Please add 
                                <select name="intQuantity" size="1" id="intQuantity">
<%
		for i = 1 to rsAssets("intAvailable")
%>
                                  <option value="<%=i%>"><%=i%></option>
<%
		next
%>
                                </select> asset(s) to my cart.</strong>&nbsp;&nbsp;<input type="hidden" name="idDescription" value="<%=request("idDescription")%>">
                                <input type="hidden" value="<%=rsAssets("idCustomer")%>" name="idCustomer"><input type="hidden" value="<%=request("idCategory")%>" name="idCategory">
                                <input type="submit" name="Submit" value="Update Cart"></font></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
<%
	end if
%>
                <tr> 
                  <td height="25"><img src="../images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table>
            </td>
            <td width="200" valign="top" bgcolor="#f5f5f5">
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr align="center" bgcolor="#5b5b5b"> 
                  <td height="35" colspan="2"><%if session("idAccess") = "R" then%><A HREF="../limited/default.asp" class="titlelink"><%else%><A HREF="../orders.asp" class="titlelink"><%end if%><img src="../images/exitshoppingcart.gif" width="120" height="19" border="0"></a></td>
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
	rsAssets.Close
	set rsAssets = nothing
	rsInCart.Close
	set rsInCart = nothing
	rsDescription.Close
	set rsDescription = nothing
	dbConnection.Close
	set dbConnection = nothing
%>