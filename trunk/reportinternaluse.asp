<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on inventory button
	buttonswitch = 4
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Get a list of Customers by user.
	set rsCustomers = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListCustomerNamesandIDs"
	else
		sql = "execute ListCustomerNamesandIDsbyAccess " & session("idUser")
	end if
	set rsCustomers = dbConnection.Execute(sql)
	
	'prime idCustomer, idInventory Status, idWarehouse
	if request("idCustomer") = "" then
		idCustomer = 0
		flag = 1 'don't print
	else
		idCustomer = request("idCustomer")
		flag = 0 'ok to print
	end if
	
	'List the inventory
	if flag = 0 then
		set rsInventory = server.CreateObject("adodb.recordset")
		if session("idAccess") < "O" then
			sql = "execute ListInventoryInternalUse " & idCustomer
		else
			sql = "execute ListInventoryInternalUsebyAccess " & idCustomer & "," & session("idUser")
		end if
		set rsInventory = dbConnection.Execute(sql)
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
<script language="JavaScript">
<!--
function newWindow(updateWin) {
  updateWindow = window.open(updateWin,'updateWin','width=300,height=275');
updateWindow.focus()
}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="reportinternaluse.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/reports-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="100%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="595" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Internal Use Report</strong></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr> 
                        <td bgcolor="#c0c0c0"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr bgcolor="#f5f5f5"> 
                              <td align="right" width="50%"><font size="2" face="Arial, Helvetica, sans-serif">Customer</font></td>
                              <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"> 
                                <select name="idCustomer" size="1" id="idCustomer">
                                  <option value="0" <%if cint(request("idCustomer")) = 0 then%>selected<%end if%>>All Customers</option>
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
                                </select>
                                </font></td>
                              <td align="left" width="50%"><font size="2" face="Arial, Helvetica, sans-serif"><input type="submit" name="Submit" value="Find"></font></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	if flag = 0 then
%>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td height="20" bgcolor="#c0c0c0" colspan="6">
                          <table width="100%" border="0" cellspacing="0" cellpadding="5">
                            <tr>
                              <td align="left"><a HREF="javascript:newWindow('excel/excelinternaluse.asp?idCustomer=<%=idCustomer%>')"><img SRC="images/exporttoexcel.gif" border="0" WIDTH="120" HEIGHT="19"></a></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                      <tr> 
                        <td height="20" bgcolor="#6699cc"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Asset Number</font></td>
                        <td height="20" bgcolor="#6699cc"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Serial Number</font></td>
                        <td height="20" bgcolor="#6699cc"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Customer</font></td>
                        <td height="20" bgcolor="#6699cc"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Description</font></td>
                        <td height="20" bgcolor="#6699cc"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Warehouse</font></td>
                        <td height="20" bgcolor="#6699cc"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Order / Location</font></td>
                      </tr>
<%
	if rsInventory.EOF then
%>
                      <tr align="center"> 
                        <td height="20" colspan="6"><font size="1" face="Arial, Helvetica, sans-serif">There are no Assets to display with this criteria</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="6" bgcolor="#c0c0c0"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
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
                      <tr bgcolor="<%=bgcolor%>"> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<A HREF="viewasset.asp?idInventory=<%=rsInventory("idInventory")%>"><%=trim(rsInventory("chrAssNum"))%></A></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrSerialNum"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrCustomer"))%></font></td>
<%
		if rsInventory("chrType") = "C" then
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrItem")) & " - " & trim(rsInventory("chrProcessor"))%></a><br>
                        &nbsp;<%=trim(rsInventory("chrMemory")) & " - " & trim(rsInventory("chrODrive"))%></font></td>
<%
		else
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrItem"))%></font></td>
<%
		end if
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrWarehouse"))%></font></td>
<%
		if rsInventory("idCart") = 0 then
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Not Available</font></td>
<%
		else
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<A HREF="viewcart.asp?idCart=<%=rsInventory("idCart")%>"><%=trim(rsInventory("chrCart"))%></A></font></td>
<%
		end if
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=location%></font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="6"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
                      </tr>
<%
		rsInventory.MoveNext
		loop
	end if
	rsInventory.Close
	set rsInventory = nothing
%>
                    </table></td>
                </tr>
<%
	end if 'from the flag
%>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
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
	rsCustomers.Close
	set rsCustomers = nothing
	dbConnection.Close
	set dbConnection = nothing
%>