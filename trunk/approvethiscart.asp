<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'Orders button
	buttonswitch = 2
%>
<!-- #include file="includes/openconn.asp" -->
<%	
	'AEG - Start out with no Expedite Charges
	idExpedite = 0 'No Fee
	
	'AEG - Find the Cart information by ID
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute FindCartbyID " & request("idCart")
	set rsCart = dbConnection.Execute(sql)
	
	'AEG - Lets figure out if there is a expidite charge for this order
	'get todays date and add 2 business days to it this is used for comparison
	select case weekday(date)
		case 1 'Sunday
			dtPadding = dateadd("d",3,date) 'Wednesday
		case 2 'Monday
			dtPadding = dateadd("d",2,date) 'Wednesday
		case 3 'Tuesday
			dtPadding = dateadd("d",2,date) 'Thursday
		case 4 'Wednesday
			dtPadding = dateadd("d",2,date) 'Friday
		case 5 'Thursday
			dtPadding = dateadd("d",4,date) 'Monday
		case 6 'Friday
			dtPadding = dateadd("d",4,date) 'Tuesday
		case 7 'Saturday
			dtPadding = dateadd("d",3,date) 'Tuesday
	end select
	'AEG - Check to make sure that the SHIP Date is greater then 3 NO CHARGE Days
	if datediff("d",dtPadding,rsCart("dtShip")) <= 0 then
	'AEG - Let's try 1 day Expedite there will be a charge of x dollars
		idExpedite = 1 'Middle Fee
		select case weekday(date)
			case 1 'Sunday
				dtPadding = dateadd("d",2,date) 'Tuesday
			case 2 'Monday
				dtPadding = dateadd("d",1,date) 'Tuesday
			case 3 'Tuesday
				dtPadding = dateadd("d",1,date) 'Wednesday
			case 4 'Wednesday
				dtPadding = dateadd("d",1,date) 'Thursday
			case 5 'Thursday
				dtPadding = dateadd("d",1,date) 'Friday
			case 6 'Friday
				dtPadding = dateadd("d",3,date) 'Monday
			case 7 'Saturday
				dtPadding = dateadd("d",2,date) 'Monday
		end select
	end if
	'AEG - Check to make sure that the SHIP Date is greater then 2 Expedite Charge Days
	if datediff("d",dtPadding,rsCart("dtShip")) <= 0 then
	'AEG - Must be the higest fee for Shipping
		idExpedite = 2 'Highest Fee
	end if
		
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="updateapproval.asp">
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
                    <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Approve Cart - <%=trim(rsCart("chrCart"))%></strong></font></td>
                  </tr>
                  <tr>
                    <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">This Cart is ready to be approved. Please select the appropriate billing description - then click the Approve Cart button.</font></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
<%
	if idExpedite > 0 then
%>
                  <tr>
                    <td bgcolor="#ff0000"><font size="2" face="Arial, Helvetica, sans-serif" color="#ffff00"><STRONG>This Order will incur Expedite Fees</STRONG></font></td>
                  </tr>
                  <tr>
                    <td height="20"><font size="1"><img src="images/ffffffdot.gif" width="1" height="1"></font></td>
                  </tr>
<%
	end if
%>                  
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif"><input type="radio" name="idBillto" value="0" checked>&nbsp;Charge to my account.</font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif"><input type="radio" name="idBillto" value="1">&nbsp;Charge to the Department / Division: <font color="#000000"><%=trim(rsCart("chrDDName")) & " / " & trim(rsCart("chrDDNumber"))%></font></font></td>
                  </tr>
                  <tr>
                    <td height="20"><font size="1"><img src="images/ffffffdot.gif" width="1" height="1"></font></td>
                  </tr>
                  <tr>
                    <td>
                      <table width="100%" border="0" cellspacing="0" cellpadding="1">
						<tr>
						  <td bgcolor="#6699cc">
							<table width="100%" border="0" cellspacing="0" cellpadding="3">
							  <tr>
								<td><font size="2" face="Arial, Helvetica, sans-serif" color="#ffffff"><strong>Loaner Agreement</strong><br><font size="1">Below is the Person who should receive the Loaner Agreement. Please verify information before submitting.</font></td>
							  </tr>
							  <tr>
							    <td bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif"><input type="checkbox" name="idLoaner" value="1">&nbsp;This cart requires a loaner agreement to be sent before it can be shipped.</font></td>
							  </tr>
<%
	if rsCart("idSendLoaner") = 1 then
		'we need to find the Users email address
		set rsUser = server.CreateObject("adodb.recordset")
		sql = "execute FindUserbyID " & rsCart("idUser")
		set rsUser = dbConnection.Execute(sql)
		'AEG - Load into temp variables
		Sendto = trim(rsUser("chrFirst")) & " " & trim(rsUser("chrLast"))
		Emailto = trim(rsUser("chrEmail"))
		'AEG - Close the user recordset
		rsUser.Close
		set rsUser = nothing
	elseif rsCart("idSendLoaner") = 2 then
		'we need to get the Apple Requestors information
		Sendto = trim(rsCart("chrARName"))
		Emailto = trim(rsCart("chrAREmail"))
	else
		'we need to get the OnSite Information
		Sendto = trim(rsCart("chrOSPerson"))
		Emailto = trim(rsCart("chrOSEmail"))
	end if
%>
							  <tr>
							    <td bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif">Send To<br>
							    <input name="chrOSPerson" type="text" id="chrOSPerson" size="40" value="<%=Sendto%>"></font></td>
							  </tr>
							  <tr>
							    <td bgcolor="#ffffff"><font size="1" face="Arial, Helvetica, sans-serif">Email Address<br>
							    <input name="chrOSEmail" type="text" id="chrOSEmail" size="45" value="<%=Emailto%>"></font></td>
							  </tr>
							</table>
						  </td>
						</tr>
					  </table>
					</td>
                  </tr>
                  <tr>
                    <td height="20"><font size="1"><img src="images/ffffffdot.gif" width="1" height="1"></font></td>
                  </tr>
                  <tr>
                    <td><font size="1"><input type="submit" name="Submit" value="Approve Cart">
                    <input type="hidden" name="idCart" value="<%=request("idCart")%>">
                    <input type="hidden" name="idExpedite" value="<%=idExpedite%>"></font></td>
                  </tr>
                </table></td>
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
	dbConnection.Close
	set dbConnection = nothing
%>