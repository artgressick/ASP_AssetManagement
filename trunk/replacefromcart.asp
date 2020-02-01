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
	'find out if there are any assets left in the cart
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute CheckoutStatsbyCartID " & session("idLoadOut")
	set rsCart = dbConnection.Execute(sql)
	
	'get the list of Assets
	set rsAssets = server.CreateObject("adodb.recordset")
	sql = "execute ListAvailable2byDescription " & request("idDescription") & "," & session("idLoadOut") & "," & rsCart("idCustomer")
	set rsAssets = dbConnection.Execute(sql)
	
	'select the first radio button
	selected = 1
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="swapasset.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/orders-nav.htm" -->
        </td>
      <td width="100%" height="100%" valign="top"><table width="625" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15"><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="610"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr>
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Replace Asset: <%=trim(rsCart("chrCart"))%></strong></font></td>
                      </tr>
                      <tr>
                        <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">To replace an asset please select from the list below and press the Submit button. The asset will be replaced with the one you chose.</font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr>
                        <td><font size="2" face="Arial, Helvetica, sans-serif"><strong>Available Assets</strong></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr bgcolor="#6699cc">
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
						<td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Asset #</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item</font></td>
                      </tr>
<%
	if rsAssets.eof then
%>
                      <tr bgcolor="#ffffff">
                        <td height="20" align="center" colspan="3"><font size="1" face="Arial, Helvetica, sans-serif">There are no Assets available to replace.</font></td>
                      </tr>
					  <tr bgcolor="#5b5b5b">
                        <td height="1" colspan="3"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsAssets.eof
		if bgswitch = 1 then
			bgswitch = 0
			bgcolor = "#ffffff"
		else
			bgswitch = 1
			bgcolor = "#f5f5f5"
		end if
%>
					  <tr bgcolor="<%=bgcolor%>">
                        <td height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><input type="radio" name="idNInventory" value="<%=rsAssets("idInventory")%>" <%if selected = 1 then%>checked<%end if%>></font></td>
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrAssNum"))%></font></td>
<%
		if rsAssets("chrType") = "C" then
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrItem")) & " - " & trim(rsAssets("chrProcessor"))%><br>
                          &nbsp;<%=trim(rsAssets("chrMemory")) & " - " & trim(rsAssets("chrHDD")) & " - " & trim(rsAssets("chrODrive"))%></font></td>
<%
		else
%>
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrItem"))%></font></td>
<%
		end if
%>
                      </tr>
                      <tr bgcolor="#5b5b5b">
                        <td height="1" colspan="3"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
		if selected = 1 then
			selected = 0
		end if
		rsAssets.movenext
		loop
	end if
%>
                    </table></td>
                </tr>
<%
	if selected = 0 then
%>
				<tr>
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
				<tr>
                  <td><input type="submit" name="Replace Asset" value="Submit"><input type="hidden" name="idOInventory" value="<%=request("idInventory")%>"></td>
                </tr>
<%
	end if
%>
                <tr>
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
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
	rsCart.Close
	set rsCart = nothing
	rsAssets.Close
	set rsAssets = nothing
	dbConnection.Close
	set dbConnection = nothing
%>