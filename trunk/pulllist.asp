<%@ Language=VBScript %>
<!-- #include file="includes/openconn.asp" -->
<%
	'prime the page
	if request("idCart") = "" then
		idCart = session("idLoadOut")
	else
		idCart = request("idCart")
	end if
	
	'find the cart information
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute FindCartbyID " & idCart
	set rsCart = dbConnection.Execute(sql)
	
	'find the cart information
	set rsAssets = server.CreateObject("adodb.recordset")
	sql = "execute ListAssetsOrderedbyCart " & idCart
	set rsAssets = dbConnection.Execute(sql)

	'find the user information
	set rsUser = server.CreateObject("adodb.recordset")
	sql = "execute FindUserbyID " & rsCart("idUser")
	set rsUser = dbConnection.Execute(sql)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Pull List</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Pull List: <%=trim(rsCart("chrCart"))%></strong></font></td>
          <td align="right" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif">Cart Information Last Updated: <%=formatdatetime(rsCart("dtUpdated"),1)%></font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="1" bgcolor="#000000"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
  </tr>
<%
	if rsCart("idExpedite") > 1 then
		select case rsCart("idExpedite")
			case 1
				bgEcolor = "#0000ff"
				text = "This cart is Expedited Status"
			case 2
				bgEcolor = "#ff0000"
				text = "This cart is Rush Status"
		end select
%>
  <tr>
    <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr bgcolor="<%=bgEcolor%>">
          <td align="center"><font size="2" color="#ffffff" face="Arial, Helvetica, sans-serif"><strong><%=text%></strong></font></td>
        </tr>
      </table></td>
  </tr>
<%
	end if
%>
  <tr>
    <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
  </tr>
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="1">
        <tr>
          <td width="25%" valign="top">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>
                  <font size="2" face="Arial, Helvetica, sans-serif"><strong>Shipping Information</strong></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrAddress"))%></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrAddress2"))%></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrAddress3"))%></font><br>
<%
  'RRF - If address4 is blank skip.  Make format look better
  if rsCart("chrAddress4") <> "" Then
%>
                  <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrAddress4"))%></font><br>
<%
  End If
%>
                  <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCity"))%>,&nbsp;<%=trim(rsCart("chrState"))%>&nbsp;<%=trim(rsCart("chrZip"))%></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCountry"))%></font><br>
                  <font size="2" face="Arial, Helvetica, sans-serif"><strong>Cart Number</strong></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("idCart"))%></font>
                </td>
              </tr>
            </table>
          </td>
          <td width="25%" valign="top">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>
                  <font size="2" face="Arial, Helvetica, sans-serif"><strong>On-Site information</strong></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrOSPerson"))%></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrOSPhone"))%></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrOSEmail"))%></font><br><br>
                  <font size="2" face="Arial, Helvetica, sans-serif"><strong>Receiving Invoice</strong></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrIPerson"))%></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrIEmail"))%></font>
                </td>
              </tr>
            </table>
          </td>
          <td width="25%" valign="top">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>
                  <font size="2" face="Arial, Helvetica, sans-serif"><strong>Entered By</strong></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif"><%=rsUser("chrFirst")%>&nbsp;<%=rsUser("chrLast")%></font><br><br>
                  <font size="2" face="Arial, Helvetica, sans-serif"><strong>Department/Division Information</strong></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif">Number: <%=trim(rsCart("chrDDNumber"))%></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif">Name: <%=trim(rsCart("chrDDName"))%></font><br><br>
                  <font size="2" face="Arial, Helvetica, sans-serif"><strong>Manager</strong></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrManager"))%></font>
                </td>
              </tr>
            </table>
          </td>
          <td width="25%" valign="top">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>
                  <font size="2" face="Arial, Helvetica, sans-serif"><strong>Dates</strong></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif">Pull: <%=formatdatetime(rsCart("dtPull"),2)%></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif">Ship: <%=formatdatetime(rsCart("dtShip"),2)%></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif">Setup: <%=formatdatetime(rsCart("dtArrival"),2)%></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif">End: <%=formatdatetime(rsCart("dtDeparture"),2)%></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif">Return: <%=formatdatetime(rsCart("dtReturn"),2)%></font><br>
                  <font size="1" face="Arial, Helvetica, sans-serif">Turn: <%=formatdatetime(rsCart("dtTurn"),2)%></font>
                </td>
              </tr>
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
    <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Shipping</strong></font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="1" bgcolor="#000000"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
  </tr>
  <tr>
    <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td width="50%"><strong><font size="2" face="Arial, Helvetica, sans-serif">Carrier Name</font></strong></td>
          <td width="50%"><strong><font size="2" face="Arial, Helvetica, sans-serif">Carrier Account Number</font></strong></td>
        </tr>
        <tr>
          <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCarrier"))%></font></td>
          <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrAccount"))%></font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Notes</strong></font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td bgcolor="#000000"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
  </tr>
  <tr>
    <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("txtShippingNotes"))%></font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="25">&nbsp;</td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Assets Requested</strong><font size="1"></font></font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td>
		<table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr bgcolor="#6699cc">
          <td align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif">Asset Number</font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif">Serial Number</font></td>
	      <td><font size="1" face="Arial, Helvetica, sans-serif">Item</font></td>
	      <td><font size="1" face="Arial, Helvetica, sans-serif">Added to Cart</font></td>
        </tr>
<%
	if not rsAssets.eof then
		do until rsAssets.eof
		if bgswitch = 1 then
			bgswitch = 0
			bgcolor = "#ffffff"
		else
			bgswitch = 1
			bgcolor = "#f5f5f5"
		end if
		'counter
		counter = counter + 1
%>
        <tr bgcolor="<%=bgcolor%>">
          <td align="center"><font size="1" face="Arial, Helvetica, sans-serif"><input type="checkbox" name="checkbox" value="checkbox"></font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsAssets("chrAssNum"))%></font></td>
          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsAssets("chrSerialNum"))%></font></td>
<%
	if rsAssets("chrType") = "C" then
%>
          <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsAssets("chrItem")) & " - " & trim(rsAssets("chrProcessor"))%><BR>
		  <%=trim(rsAssets("chrMemory")) & " - " & trim(rsAssets("chrHDD")) & " - " & trim(rsAssets("chrODrive"))%></font></td>
<%
	else
%>
	      <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsAssets("chrItem"))%></font></td>
<%
	end if
%>
	      <td><font size="1" face="Arial, Helvetica, sans-serif"><%=formatdatetime(rsAssets("dtStamp"),2)%></font></td>
        </tr>
        <tr bgcolor="<%=bgcolor%>">
          <td align="center"><font size="1" face="Arial, Helvetica, sans-serif"></font></td>
          <td colspan="4"><font size="1" face="Arial, Helvetica, sans-serif" color="#0000ff"><%=trim(rsAssets("txtNotes"))%></font></td>
        </tr>
<%
		rsassets.movenext
		loop
	end if
%>
      </table>
	</td>
  </tr>
  <tr>
    <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Totals</strong></font></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="1" bgcolor="#000000"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
  </tr>
  <tr>
    <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr>
          <td><font size="1" face="Arial, Helvetica, sans-serif">Total Pieces in Cart: <%=counter%></font></td>
        </tr>
      </table></td>
  </tr>
</table>

</body>
</html>
<%
	rsCart.close
	set rsCart = nothing
	rsAssets.Close
	set rsAssets = nothing
	rsUser.Close
	set rsUser = nothing
	dbConnection.Close
	set dbConnection = nothing
%>