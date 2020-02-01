<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on inventory button
	buttonswitch = 3
	
	'First Line switch
	firstline = 0
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'get a list of open orders. Make sure that the orders come from accounts they can add to.
	set rsCustomers = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListCustomerNamesandIDs"
	else
		sql = "execute ListCustomerNamesandIDsbyAccess " & session("idUser")
	end if
	set rsCustomers = dbConnection.Execute(sql)
	
	'Get a list of the Orders
	set rsInvStatus = server.CreateObject("adodb.recordset")
	sql = "execute ListInvStatus"
	set rsInvStatus = dbConnection.Execute(sql)
	
	'Get a list of the warehouses
	set rsWarehouse = server.CreateObject("adodb.recordset")
	sql = "execute ListWarehouses"
	set rsWarehouse = dbConnection.Execute(sql)
	
	'Find the Description
	set rsDescription = server.CreateObject("adodb.recordset")
	sql = "execute ViewDescriptionwithCategory " & request("idDescription")
	set rsDescription = dbConnection.Execute(sql)
	
	'prime idCustomer, idInventory Status, idWarehouse
	if request("idCustomer") = "" then
		idCustomer = 0
		idInventoryStatus = 0
		idWarehouse = 0
	else
		idCustomer = request("idCustomer")
		idInventoryStatus = request("idInventoryStatus")
		idWarehouse = request("idWarehouse")
	end if
	
	'List the inventory
	set rsInventory = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListInventorybyDescriptionwithFilter " & request("idDescription") & "," & idCustomer & "," & idInventoryStatus & "," & idWarehouse
	else
		sql = "execute ListInventorybyDescriptionwithFilterbyAccess " & request("idDescription") & "," & idCustomer & "," & idInventoryStatus & "," & idWarehouse & "," & session("idUser")
	end if
	set rsInventory = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="viewdescription.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
	<!-- #include file="includes/top.htm" -->
	<!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
		<!-- #include file="includes/inventory-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="100%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="595" height="1"></td>
                </tr>
                <tr> 
                  <td align="center">
					<table width="595" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td align="center"><img src="assetimages/<%=rsDescription("chrLocation")%>"></td>
                        <td valign="top">
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
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr> 
                        <td bgcolor="#5b5b5b"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr bgcolor="#f5f5f5"> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Customer<br>
                                <select name="idCustomer" size="1" id="idCustomer">
                                  <option value="0" <%if cint(request("idCustomer")) = 0 then%>selected<%end if%>>All</option>
<%
	if not rsCustomers.eof then
		do until rsCustomers.eof
%>
                                  <option value="<%=rsCustomers("idCustomer")%>" <%if cint(request("idCustomer")) = rsCustomers("idCustomer") then%>selected<%end if%>><%=trim(rsCustomers("chrCustomer"))%></option>
<%
		rsCustomers.Movenext
		loop
	end if
%>
                                </select>
                                </font></td>
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Status<br>
                                <select name="idInventoryStatus" size="1" id="idInventoryStatus">
                                  <option value="0" <%if cint(request("idStatus")) = 0 then%>selected<%end if%>>All</option>
<%
	if not rsInvStatus.eof then
		do until rsInvStatus.eof
%>
                                  <option value="<%=rsInvStatus("idStatus")%>" <%if cint(request("idStatus")) = rsInvStatus("idStatus") then%>selected<%end if%>><%=trim(rsInvStatus("chrInventoryStatus"))%></option>
<%
		rsInvStatus.movenext
		loop
	end if
%>
                                </select>
                                </font></td>
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Warehouse<br>
                                <select name="idWarehouse" size="1" id="idWarehouse">
                                  <option value="0" <%if cint(request("idWarehouse")) = 0 then%>selected<%end if%>>All</option>
<%
	if not rsWarehouse.eof then
		do until rsWarehouse.eof
%>
                                  <option value="<%=rsWarehouse("idWarehouse")%>" <%if cint(request("idWarehouse")) = rsWarehouse("idWarehouse") then%>selected<%end if%>><%=trim(rsWarehouse("chrWarehouse"))%></option>
<%
		rsWarehouse.movenext
		loop
	end if
%>
                                </select>
                                </font></td>
                              <td valign="bottom"> <font size="1" face="Arial, Helvetica, sans-serif"> 
                                <input type="submit" name="Submit" value="Filter"><input type="hidden" name="idDescription" value="<%=request("idDescription")%>">
                                </font></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr bgcolor="#6699cc"> 
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Asset #</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Serial #</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Warehouse</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Status</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Location</font></td>
                        <td height="20" align="right"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Options&nbsp;</font></td>
                      </tr>
<%
	if rsInventory.EOF then
%>
                      <tr align="center"> 
                        <td height="20" colspan="6"><font size="1" face="Arial, Helvetica, sans-serif">There are no Assets in this Descriptions with this criteria.</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="6" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsInventory.EOF
		if bgswitch = 1 then
			bgcolor = "#ffffff"
			bgswitch = 0
		else
			bgcolor = "#f5f5f5"
			bgswitch = 1
		end if
%>
                      <tr> 
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<a href="viewasset.asp?idInventory=<%=rsInventory("idInventory")%>"><%=rsInventory("chrAssNum")%></a></font></td>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrSerialNum"))%></font></td>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrWarehouse"))%></font></td>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrInventoryStatus"))%></font></td>
<%
	if rsInventory("idCart") = 0 then
%>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Warehouse</font></td>
<%
	else
%>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrCart"))%></font></td>
<%
	end if
%>
                        <td height="20" align="right" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif"><%if session("idAccess") < "O" then%><a HREF="editinventory.asp?idInventory=<%=rsInventory("idInventory")%>">Edit</a><%end if%>&nbsp;</font></td>
                      </tr>
                      <tr bgcolor="#5b5b5b"> 
                        <td height="1" colspan="6"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
		rsInventory.MoveNext
		loop
	end if
%>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
          </tr>
        </table></td>
    </tr>
	<!-- #Begin bottom part -->
    <!-- #include file="includes/bottom.htm" -->
  </table>
  </form>
</body>
</html>
<%
	rsCustomers.close
	set rsCustomers = nothing
	rsInvStatus.Close
	set rsInvStatus = nothing
	rsWarehouse.close
	set rsWarehouse = nothing
	rsDescription.Close
	set rsDescription = nothing
	rsInventory.Close
	set rsInventory = nothing
	dbConnection.Close
	set dbConnection = nothing
%>