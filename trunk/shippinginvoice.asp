<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'find the cart information
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute ViewInvoicebyCart " & request("idCart")
	set rsCart = dbConnection.Execute(sql)
	
	'find the cart information
	set rsAssets = server.CreateObject("adodb.recordset")
	sql = "execute ListInvoiceAssetsbyCart " & request("idCart")
	set rsAssets = dbConnection.Execute(sql)	
	
	'find the cart information
	set rsWarehouse = server.CreateObject("adodb.recordset")
	sql = "execute FindWarehousebyID " & rsCart("idWarehouse")
	set rsWarehouse = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	      <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
	  </tr>
	  <tr> 
	      <td>
		  	<table width="100%" border="0" cellspacing="0" cellpadding="3">
			  <tr> 
			  	<td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Shipping Manifest for: <%=trim(rscart("chrCart"))%></strong></font></td>
          <td align="right" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
        </tr>
      </table>
	</td>
  </tr>            
  <tr> 
    <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
  </tr>
  <tr> 
    <td bgcolor="#f5f5f5">
		<table width="100%" border="0" cellspacing="1" cellpadding="3">
		<tr align="center">                         
          <td><font size="1" face="Arial, Helvetica, sans-serif">Customer</font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif">Cart Number</font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif">Created</font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif">Status</font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif">Billed</font></td>
        </tr>
        <tr align="center" bgcolor="#ffffff"> 
          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCustomer"))%></font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("idCart"))%></font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=formatdatetime(rsCart("dtStamp"),2)%></font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCartStatus"))%></font></td>
<%
	if rsCart("idBilled") = 0 then
		Billed = "No"
	else
		Billed = "Yes"
	end if
%>
          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=Billed%></font></td>
        </tr>
    </table></td>
  </tr>
  <tr> 
    <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
  </tr>
  <tr> 
    <td><table width="100%" border="0" cellspacing="0" cellpadding="1">
        <tr> 
          <td bgcolor="#f5f5f5">
		    <table width="100%" border="0" cellspacing="0" cellpadding="0">
			    <tr> 
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Asset  #</font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item</font></td>
                <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Requested by</font></td>
                <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Requested&nbsp;</font></td>
                <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Shipped&nbsp;</font></td>
                <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Returned&nbsp;</font></td>
              </tr>
<%
	if rsAssets.EOF then
%>
              <tr align="center"> 
                <td height="20" colspan="6" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif">No Assets have been requested at this time.</font></td>
              </tr>
<%
	else
		do until rsAssets.EOF
			if bgswitch = 1 then
				bgcolor = "#f5f5f5"
				bgswitch = 0
			else
				bgcolor = "#ffffff"
				bgswitch = 1
			end if
%>
              <tr bgcolor="<%=bgcolor%>"> 
                <td height="20" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrAssNum"))%></font></td>
<%
			if rsAssets("chrType") = "C" then
%>
                <td height="20" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;&nbsp;<%=trim(rsAssets("chrItem")) & " - " & trim(rsAssets("chrProcessor"))%><br>
                              &nbsp;<%=trim(rsAssets("chrMemory")) & " - " & trim(rsAssets("chrHDD")) & " - " & trim(rsAssets("chrODrive"))%></font></td>
<%
			else
%>
                <td height="20" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrItem"))%></font></td>
<%
			end if
%>
                <td height="20" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrFirst")) & " " & trim(rsAssets("chrLast"))%></font></td>
                <td height="20" align="center" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif"><%=rsAssets("intOrdered")%></font></td>
<%
			subOrdered = subOrdered + rsAssets("intOrdered")
%>
                <td height="20" align="center" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif"><%=rsAssets("intShipped")%></font></td>
<%
			subShipped = subShipped + rsAssets("intShipped")
%>
                <td height="20" align="center" bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif"><%=rsAssets("intReturned")%></font></td>
<%
			subReturned = subReturned + rsAssets("intReturned")
%>
              </tr>
<%
			rsAssets.MoveNext
		loop
%>
			  <tr>
			  	<td height="20" colspan="3" align="right" ><font size="1" face="Arial, Helvetica, sans-serif">Totals&nbsp;</font></td>
                <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=subOrdered%>&nbsp;</font></td>
                <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=subShipped%>&nbsp;</font></td>
                <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=subReturned%>&nbsp;</font></td>
              </tr>
<%
	end if
%>
           </table></td>
        </tr>
     </table></td>
  </tr>
  <tr> 
    <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
  </tr>
  <tr> 
    <td bgcolor="#f5f5f5">
		<table width="100%" border="0" cellspacing="1" cellpadding="3">
        <tr> 
          <td><font size="1" face="Arial, Helvetica, sans-serif">Ship Date</font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif">Carrier</font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif">Account</font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif">End Date</font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif">Return VIA</font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif">D/D Name</font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif">D/D #</font></td>
        </tr>
        <tr bgcolor="#ffffff"> 
          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=formatdatetime(rsCart("dtShip"),2)%></font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCarrier"))%></font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrAccount"))%></font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=formatdatetime(rsCart("dtDeparture"),2)%></font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif">FedEx 2-Day</font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrDDName"))%></font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrDDNumber"))%></font></td>
        </tr>
        <tr valign="top" bgcolor="#ffffff"> 
          <td colspan="3"> <p><font size="2" face="Arial, Helvetica, sans-serif"><strong>To:</strong></font></p>
            <p><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrOSPerson"))%><br>
              <%=trim(rsCart("chrAddress"))%><br>
              <%=trim(rsCart("chrAddress2"))%><br>
              <%=trim(rsCart("chrAddress3"))%><br>
              <%=trim(rsCart("chrAddress4"))%><br>
              <%=trim(rsCart("chrCity")) & ", " & trim(rsCart("chrState")) & " " & trim(rsCart("chrZip"))%><br>
              <%=trim(rsCart("chrOSPhone"))%></font></p></td>
          <td colspan="2"> <p><font size="2" face="Arial, Helvetica, sans-serif"><strong>Return:</strong></font></p>
            <p><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsWarehouse("chrWarehouse"))%> Warehouse<br>
              <%=trim(rsWarehouse("chrAddress"))%><br>
              <%=trim(rsWarehouse("chrAddress2"))%><br>
              <%=trim(rsWarehouse("chrCity")) & ", " & trim(rsWarehouse("chrState")) & " " & trim(rsWarehouse("chrZip"))%><br>
              Phone: <%=trim(rsWarehouse("chrPhone"))%></font></p></td>
          <td colspan="2"> <p><font size="2" face="Arial, Helvetica, sans-serif"><strong>Bill To:</strong></font></p>
            <p><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCustomer"))%><br>
			  <%=trim(rsCart("chrBAddress"))%><br>
			  <%=trim(rsCart("chrBAddress2"))%><br>
			  <%=trim(rsCart("chrBCity")) & ", " & trim(rsCart("chrBState")) & " " & trim(rsCart("chrBZip"))%><br>
			  Phone: <%=trim(rsCart("chrBPhone"))%><br>
			  Fax: <%=trim(rsCart("chrBFax"))%></font></p></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
  </tr>
</table>
</body>
</html>
<%
	rsCart.close
	set rsCart = nothing
	rsAssets.close
	set rsAssets = nothing
	rsWarehouse.Close
	set rsWarehouse = nothing
	dbConnection.Close
	set dbConnection = nothing
%>