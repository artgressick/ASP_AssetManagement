<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'compare the dates and make sure that the person does not put the end date before the begin date.
	if datediff("d",request("dtArrival"),request("dtDeparture")) > 0 then
		
		'set the temporary variables
		dtArrival = request("dtArrival") 'This is the get there date
		dtDeparture = request("dtDeparture") 'This is the end date
		
		'get todays date and add 1 business day to it this is used for comparison
		select case weekday(date)
			case 1 'Sunday
				dtPadding = dateadd("d",1,date) 'Monday
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
		'Start with 3 days of shipping 
		'Configure the Ship Date to be minus 3 business days from the Arrival Date.
		select case weekday(dtArrival)
			case 1 'Sunday
				dtShip = dateadd("d",-4,dtArrival) 'Wednesday
			case 2 'Monday
				dtShip = dateadd("d",-4,dtArrival) 'Thursday
			case 3 'Tuesday
				dtShip = dateadd("d",-4,dtArrival) 'Friday
			case 4 'Wednesday
				dtShip = dateadd("d",-5,dtArrival) 'Friday
			case 5 'Thursday
				dtShip = dateadd("d",-3,dtArrival) 'Monday
			case 6 'Friday
				dtShip = dateadd("d",-3,dtArrival) 'Tuesday
			case 7 'Saturday
				dtShip = dateadd("d",-3,dtArrival) 'Wednesday
		end select
		'Now that we have a Ship Date we need to see if we can actually ship it in 3 days
		if dtShip <= dtPadding then
		'Start with 2 days of shipping 
		'Configure the Ship Date to be minus 2 business days from the Arrival Date.
			select case weekday(dtArrival)
				case 1 'Sunday
					dtShip = dateadd("d",-3,dtArrival) 'Thursday
				case 2 'Monday
					dtShip = dateadd("d",-4,dtArrival) 'Thursday
				case 3 'Tuesday
					dtShip = dateadd("d",-4,dtArrival) 'Friday
				case 4 'Wednesday
					dtShip = dateadd("d",-2,dtArrival) 'Monday
				case 5 'Thursday
					dtShip = dateadd("d",-2,dtArrival) 'Tuesday
				case 6 'Friday
					dtShip = dateadd("d",-2,dtArrival) 'Wednesday
				case 7 'Saturday
					dtShip = dateadd("d",-2,dtArrival) 'Thursday
			end select
		end if
		'Now that we have a Ship Date we need to see if we can actually ship it in 2 days
		if dtShip <= dtPadding then
		'Start with 1 days of shipping 
		'Configure the Ship Date to be minus 1 business days from the Arrival Date.
			select case weekday(dtArrival)
				case 1 'Sunday
					dtShip = dateadd("d",-2,dtArrival) 'Friday
				case 2 'Monday
					dtShip = dateadd("d",-3,dtArrival) 'Friday
				case 3 'Tuesday
					dtShip = dateadd("d",-1,dtArrival) 'Monday
				case 4 'Wednesday
					dtShip = dateadd("d",-1,dtArrival) 'Tuesday
				case 5 'Thursday
					dtShip = dateadd("d",-1,dtArrival) 'Wednesday
				case 6 'Friday
					dtShip = dateadd("d",-1,dtArrival) 'Thursday
				case 7 'Saturday
					dtShip = dateadd("d",-1,dtArrival) 'Friday
			end select
		end if
		'Now that we have a Ship Date we need to see if we can actually ship it in 1 day
		if not dtShip <= dtPadding then
		'it is ok now we need to insert the information
		'figure the Pull date to minus 1 day from the ship date
			select case weekday(dtShip)
				case 1 'Sunday
					dtPull = dateadd("d",-2,dtShip) 'Friday
				case 2 'Monday
					dtPull = dateadd("d",-3,dtShip) 'Friday
				case 3 'Tuesday
					dtPull = dateadd("d",-1,dtShip) 'Monday
				case 4 'Wednesday
					dtPull = dateadd("d",-1,dtShip) 'Tuesday
				case 5 'Thursday
					dtPull = dateadd("d",-1,dtShip) 'Wednesday
				case 6 'Friday
					dtPull = dateadd("d",-1,dtShip) 'Thursday
				case 7 'Saturday
					dtPull = dateadd("d",-1,dtShip) 'Friday
			end select	
			'figure the Return Date to the next business day
			select case weekday(dtDeparture)
				case 1 'Sunday
					dtReturn = dateadd("d",1,dtDeparture) 'Monday
				case 2 'Monday
					dtReturn = dateadd("d",1,dtDeparture) 'Tuesday
				case 3 'Tuesday
					dtReturn = dateadd("d",1,dtDeparture) 'Wednesday
				case 4 'Wednesday
					dtReturn = dateadd("d",1,dtDeparture) 'Thursday
				case 5 'Thursday
					dtReturn = dateadd("d",1,dtDeparture) 'Friday
				case 6 'Friday
					dtReturn = dateadd("d",3,dtDeparture) 'Monday
				case 7 'Saturday
					dtReturn = dateadd("d",2,dtDeparture) 'Monday
			end select
			'figure the Turn Date to 5 business days
			select case weekday(dtDeparture)
				case 1 'Sunday
					dtTurn = dateadd("d",5,dtDeparture) 'Friday
				case 2 'Monday
					dtTurn = dateadd("d",7,dtDeparture) 'Monday
				case 3 'Tuesday
					dtTurn = dateadd("d",7,dtDeparture) 'Tuesday
				case 4 'Wednesday
					dtTurn = dateadd("d",7,dtDeparture) 'Wednesday
				case 5 'Thursday
					dtTurn = dateadd("d",7,dtDeparture) 'Thursday
				case 6 'Friday
					dtTurn = dateadd("d",7,dtDeparture) 'Friday
				case 7 'Saturday
					dtTurn = dateadd("d",6,dtDeparture) 'Friday
			end select
			
			'AEG - Find the Warehouse to send an email.
			set rsAssets = server.CreateObject("adodb.recordset")
			sql = "execute ChangingCartDates " & request("idCart") & ",'" & dtPull & "','" & dtTurn & "'"
			set rsAssets = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Update Check</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="updatecart.asp">
  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="10" height="10" align="left" valign="top"><img src="images/topleftblue.gif" width="10" height="10"></td>
            <td width="780" height="10"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td><img src="images/eamlogo.gif" width="150" height="45"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                </tr>
              </table></td>
            <td width="10" height="10" align="right" valign="top"><img src="images/toprightblue.gif" width="10" height="10"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr>
            <td bgcolor="#FFFFFF">
			  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td height="35"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td align="center"><font color="#0000FF" size="3" face="Arial, Helvetica, sans-serif"><strong>Assets Affected</strong></font></td>
                </tr>
                <tr> 
                  <td><font size="1" face="Arial, Helvetica, sans-serif">These are all of the assets that will be removed from you cart if you continue with this update.
                  If you cancel this update no dates will be change and the assets listed below will remain in you cart.</font></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
					  <tr bgcolor="#6699cc">
						<td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Asset Number</font></td>
						<td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Serial Number</font></td>
						<td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item</font></td>
					  </tr>
<%
	'set the counter
	counter = 0
	if rsAssets.EOF then
%>
					  <tr>
						<td colspan="3" height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">No Assets are affected by this change</font></td>
					  </tr>
					  <tr>
						<td colspan="3" height="1" align="center" bgcolor="#c0c0c0"><img src="images/c0c0c0dot.gif" WIDTH="1" HEIGHT="1"></td>
					  </tr>
<%
	else
		do until rsAssets.EOF
		if bgswitch = 1 then
			bgswitch = 0
			bgcolor = "#ffffff"
		else
			bgswitch = 1
			bgcolor = "#f5f5f5"
		end if
		counter = counter + 1
%>
					  <tr bgcolor="<%=bgcolor%>">
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrAssNum"))%><input name="<%=counter%>" type="hidden" value="<%=rsAssets("idInventory")%>"></font></td>
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrSerialNum"))%></font></td>
<%
		if rsAssets("chrType") = "C" then
%>
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrItem")) & " " & trim(rsAssets("chrProcessor"))%><br>
						&nbsp;<%=trim(rsAssets("chrMemory")) & " - " & trim(rsAssets("chrODrive")) & " - " & trim(rsAssets("chrHDD"))%></font></td>
<%
		else
%>
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrItem"))%></font></td>
<%
		end if
%>
					  </tr>
					  <tr>
						<td colspan="3" height="1" align="center" bgcolor="#c0c0c0"><img src="images/c0c0c0dot.gif" WIDTH="1" HEIGHT="1"></td>
					  </tr>
<%
		rsAssets.MoveNext
		loop
	end if
%>
					</table>
                  </td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td align="center"><input name="Cancel" type="submit" value="Cancel Update">&nbsp;&nbsp;
                  <input name="Update" type="submit" value="Update Cart">
                  <input name="idSpecial" type="hidden" value="1">
                  <input name="intCounter" type="hidden" value="<%=counter%>">
                  <input name="idCart" type="hidden" value="<%=request("idCart")%>">
                  <input name="idStatus" type="hidden" value="<%=request("idStatus")%>">
                  <input name="idBilled" type="hidden" value="<%=request("idBilled")%>">
                  <input name="idLoaner" type="hidden" value="<%=request("idLoaner")%>">
                  <input name="chrOrder" type="hidden" value="<%=replace(request("chrOrder"),"'","''")%>">
                  <input name="chrDDName" type="hidden" value="<%=replace(request("chrDDName"),"'","''")%>">
                  <input name="chrDDNumber" type="hidden" value="<%=replace(request("chrDDNumber"),"'","''")%>">
                  <input name="chrManager" type="hidden" value="<%=replace(request("chrManager"),"'","''")%>">
                  <input name="chrIPerson" type="hidden" value="<%=replace(request("chrIPerson"),"'","''")%>">
                  <input name="chrIEmail" type="hidden" value="<%=replace(request("chrIEmail"),"'","''")%>">
                  <input name="dtPull" type="hidden" value="<%=dtPull%>">
                  <input name="dtShip" type="hidden" value="<%=dtShip%>">
                  <input name="dtArrival" type="hidden" value="<%=dtArrival%>">
                  <input name="dtArrivalTime" type="hidden" value="<%=request("dtArrivalTime")%>">
                  <input name="dtDeparture" type="hidden" value="<%=dtDeparture%>">
                  <input name="dtDepartureTime" type="hidden" value="<%=request("dtDepartureTime")%>">
                  <input name="dtReturn" type="hidden" value="<%=dtReturn%>">
                  <input name="dtTurn" type="hidden" value="<%=dtTurn%>">
                  <input name="chrAddress" type="hidden" value="<%=replace(request("chrAddress"),"'","''")%>">
                  <input name="chrAddress2" type="hidden" value="<%=replace(request("chrAddress2"),"'","''")%>">
                  <input name="chrAddress3" type="hidden" value="<%=replace(request("chrAddress3"),"'","''")%>">
                  <input name="chrAddress4" type="hidden" value="<%=replace(request("chrAddress4"),"'","''")%>">
                  <input name="chrCity" type="hidden" value="<%=replace(request("chrCity"),"'","''")%>">
                  <input name="chrState" type="hidden" value="<%=replace(request("chrState"),"'","''")%>">
                  <input name="chrZip" type="hidden" value="<%=replace(request("chrZip"),"'","''")%>">
                  <input name="chrCountry" type="hidden" value="<%=replace(request("chrCountry"),"'","''")%>">
                  <input name="chrARName" type="hidden" value="<%=replace(request("chrARName"),"'","''")%>">
                  <input name="chrAREmail" type="hidden" value="<%=replace(request("chrAREmail"),"'","''")%>">
                  <input name="chrOSPerson" type="hidden" value="<%=replace(request("chrOSPerson"),"'","''")%>">
                  <input name="chrOSEmail" type="hidden" value="<%=replace(request("chrOSEmail"),"'","''")%>">
                  <input name="chrOSPhone" type="hidden" value="<%=replace(request("chrOSPhone"),"'","''")%>">
                  <input name="chrOSFax" type="hidden" value="<%=replace(request("chrOSFax"),"'","''")%>">
                  <input name="chrCarrier" type="hidden" value="<%=replace(request("chrCarrier"),"'","''")%>">
                  <input name="chrAccount" type="hidden" value="<%=replace(request("chrAccount"),"'","''")%>">
                  <input name="idPurpose" type="hidden" value="<%=request("idPurpose")%>">
                  <input name="idExpedite" type="hidden" value="<%=request("idExpedite")%>">
                  <input name="txtNotes" type="hidden" value="<%=replace(request("txtNotes"),"'","''")%>">
                  <input name="txtShippingNotes" type="hidden" value="<%=replace(request("txtShippingNotes"),"'","''")%>">
                  <input name="idType" type="hidden" value="<%=request("idType")%>">
                  <input name="idCustomer" type="hidden" value="<%=request("idCustomer")%>">
                  <input name="idUser" type="hidden" value="<%=request("idUser")%>">
                  <input name="idRStatus" type="hidden" value="<%=request("idRStatus")%>">
                  <input name="chrSearch" type="hidden" value="<%=request("chrSearch")%>">
                  </td>
                </tr>
                <tr> 
                  <td height="50"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table>
            </td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="10" height="10" align="left" valign="bottom"><img src="images/bottomleftblue.gif" width="10" height="10"></td>
            <td width="780" height="10"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td align="center"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Copyright © 2004 
                    techIT Solutions LLC. <br>
                    Asset Management Enterprise Portal 4.0 &amp; Corporate Business 
                    Intelligence as products of techIT Solutions. </font></td>
                </tr>
              </table></td>
            <td width="10" height="10" align="right" valign="bottom"><img src="images/bottomrightblue.gif" width="10" height="10"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
			rsAssets.Close
			set rsAssets = nothing
			dbConnection.Close
			set dbConnection = nothing
		'We cannot ship the order becuase of the dates
		else 'this is if we cannot ship the order by the dates provided.
			'then we cannot send the Assets
			'let the user know the possible date to ship 
			select case weekday(date)
				case 1 'Sunday
					dtPossible = dateadd("d",3,date) 'Wednesday
				case 2 'Monday
					dtPossible = dateadd("d",3,date) 'Thursday
				case 3 'Tuesday
					dtPossible = dateadd("d",3,date) 'Friday
				case 4 'Wednesday
					dtPossible = dateadd("d",3,date) 'Saturday
				case 5 'Thursday
					dtPossible = dateadd("d",5,date) 'Tuesday
				case 6 'Friday
					dtPossible = dateadd("d",5,date) 'Wednesday
				case 7 'Saturday
					dtPossible = dateadd("d",4,date) 'Wednesday
			end select
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Entry Error</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="10" height="10" align="left" valign="top"><img src="images/topleftblue.gif" width="10" height="10"></td>
            <td width="780" height="10"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td><img src="images/eamlogo.gif" width="150" height="45"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                </tr>
              </table></td>
            <td width="10" height="10" align="right" valign="top"><img src="images/toprightblue.gif" width="10" height="10"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr>
            <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td height="35"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td align="center"><font color="#FF0000" size="3" face="Arial, Helvetica, sans-serif"><strong>Add Order - Error</strong></font></td>
                </tr>
                <tr> 
                  <td height="45"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr>
                  <td align="center"><p><font size="2" face="Arial, Helvetica, sans-serif">You have entered a Delivery Date that does not allow sufficient time to process your order.</font></p>
                  <p><font size="2" face="Arial, Helvetica, sans-serif">Please click on the "Back" button in your browser and change the <b>Delivery Date</b> to occur no earlier than <b><font color="#FF0000"><%=formatdatetime(dtPossible,2)%></font></b>.</font></p>
                  <p><font size="2" face="Arial, Helvetica, sans-serif">If you have any questions or concerns regarding this matter, please contact your pool manager.</font></p></td>
                </tr> 
                <tr> 
                  <td height="50"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="10" height="10" align="left" valign="bottom"><img src="images/bottomleftblue.gif" width="10" height="10"></td>
            <td width="780" height="10"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td align="center"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Copyright &copy; 2003 
                    techIT Solutions LLC. <br>
                    Asset Management Enterprise Portal 4.0 &amp; Corporate Business 
                    Intelligence as products of techIT Solutions. </font></td>
                </tr>
              </table></td>
            <td width="10" height="10" align="right" valign="bottom"><img src="images/bottomrightblue.gif" width="10" height="10"></td>
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
		end if
	else 'this is if the Departure Date is before the Arrival Date.
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Entry Error</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="10" height="10" align="left" valign="top"><img src="images/topleftblue.gif" width="10" height="10"></td>
            <td width="780" height="10"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td><img src="images/eamlogo.gif" width="150" height="45"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                </tr>
              </table></td>
            <td width="10" height="10" align="right" valign="top"><img src="images/toprightblue.gif" width="10" height="10"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr>
            <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td height="35"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td align="center"><font color="#FF0000" size="3" face="Arial, Helvetica, sans-serif"><strong>Entry 
                    Error</strong></font></td>
                </tr>
                <tr> 
                  <td height="45"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr>
                  <td align="center"><p><font size="3" face="Arial, Helvetica, sans-serif" color="#FF0000"><b>DATE CORRECTION REQUIRED !</b></font></p>
                  <p><font size="2" face="Arial, Helvetica, sans-serif">You have entered an End date that is before the ship date. Please go back and change the date.</font></p></td>
                </tr>
                <tr> 
                  <td height="50"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="10" height="10" align="left" valign="bottom"><img src="images/bottomleftblue.gif" width="10" height="10"></td>
            <td width="780" height="10"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td align="center"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Copyright &copy; 2003 
                    techIT Solutions LLC. <br>
                    Asset Management Enterprise Portal 4.0 &amp; Corporate Business 
                    Intelligence as products of techIT Solutions. </font></td>
                </tr>
              </table></td>
            <td width="10" height="10" align="right" valign="bottom"><img src="images/bottomrightblue.gif" width="10" height="10"></td>
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
	end if
%>