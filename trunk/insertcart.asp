	<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'check the Dates
	if cint(request("idType")) > 1 then
		dtArrival = request("dtArrival") 'this is the get there date
		dtDeparture = dateadd("yyyy",20,request("dtArrival")) 'Add twenty years to the Arrival Date
	else
		dtArrival = request("dtArrival") 'this is the get there date
		dtDeparture = request("dtDeparture") 'this is the end date
	end if
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'compare the dates and make sure that the person does not put the end date before the begin date.
	if datediff("d",dtArrival,dtDeparture) > 0 then
		
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
	
			'get the fields uploaded to the database and send an email. This page also should
			'tell the user that an email has been sent to all of the people and list it.
			
			if request("ckAddress") = 0 then
				'load the Address if they used a stored address.
				set rsAddress = server.CreateObject("adodb.recordset")
				sql = "execute FindAddressbyID " & request("idAddress")
				set rsAddress = dbConnection.Execute(sql)
				
				chrAddress = replace(rsAddress("chrAddress"),"'","''")
				chrAddress2 = replace(rsAddress("chrAddress2"),"'","''")
				chrAddress3 = replace(rsAddress("chrAddress3"),"'","''")
				chrAddress4 = replace(rsAddress("chrAddress4"),"'","''")
				chrCity = replace(rsAddress("chrCity"),"'","''")
				chrState = replace(rsAddress("chrState"),"'","''")
				chrZip = replace(rsAddress("chrZip"),"'","''")
				chrCountry = replace(rsAddress("chrCountry"),"'","''")
				
				'close the connections
				rsAddress.Close
				set rsAddress = nothing
			else
				chrAddress = replace(request("chrAddress"),"'","''")
				chrAddress2 = replace(request("chrAddress2"),"'","''")
				chrAddress3 = replace(request("chrAddress3"),"'","''")
				chrAddress4 = replace(request("chrAddress4"),"'","''")
				chrCity = replace(request("chrCity"),"'","''")
				chrState = replace(request("chrState"),"'","''")
				chrZip = replace(request("chrZip"),"'","''")
				chrCountry = replace(request("chrCountry"),"'","''")
			end if
			'--------------------------------------------------------------------------------------------
			if request("ckCarrier") = 0 then
				'Load the Saved Carrier information if asked to
				set rsCarrier = server.CreateObject("adodb.recordset")
				sql = "execute FindCarrierbyID " & request("idCarrier")
				set rsCarrier = dbConnection.Execute(sql)
				
				chrCarrier = replace(rsCarrier("chrCarrier"),"'","''")
				chrAccount = replace(rsCarrier("chrAccount"),"'","''")
				
				'close the connections
				rsCarrier.Close
				set rsCarrier = nothing
			else
				chrCarrier = replace(request("chrCarrier"),"'","''")
				chrAccount = replace(request("chrAccount"),"'","''")
			end if
			'----------------------------------------------------------------------------------------------
			'prebill the SED, NED and XED
			if cint(request("idCustomer")) = 6 or cint(request("idCustomer")) = 7 or cint(request("idCustomer")) = 9 then
				idBilled = 1
			else
				idBilled = 0
			end if
			'----------------------------------------------------------------------------------------------
			'Send the information string with contents from the form.
			set rsCart = server.CreateObject("adodb.recordset")
			sql = "execute InsertCart3 " & _
				request("idOrder") & "," &_
				request("idType") & "," &_
				request("idCustomer") & ",'" &_
				request("idShow2Show") & "'," &_
				request("idSendLoaner") & "," &_ 
				request("idEmailPrior") & "," &_ 
				request("idEmailLate") & "," &_ 
				idBilled & ",'" &_
				replace(request("chrCart"),"'","''") & "','" &_
				replace(request("chrDDName"),"'","''") & "','" &_
				replace(request("chrDDNumber"),"'","''") & "','" &_
				replace(request("chrManager"),"'","''") & "','" &_
				replace(request("chrIPerson"),"'","''") & "','" &_
				replace(request("chrIEmail"),"'","''") & "','" &_
				dtPull & "','" &_
				dtShip & "','" &_
				dtArrival & "','" &_
				request("dtArrivalTime") & "','" &_
				dtDeparture & "','" &_
				request("dtDepartureTime") & "','" &_
				dtReturn & "','" &_
				dtTurn & "','" &_
				request("idSaveAddress") & "','" &_
				chrAddress & "','" &_
				chrAddress2 & "','" &_
				chrAddress3 & "','" &_
				chrAddress4 & "','" &_
				chrCity & "','" &_
				chrState & "','" &_
				chrZip & "','" &_
				chrCountry & "','" &_
				replace(request("chrSavedAddressName"),"'","''") & "','" &_
				replace(request("chrARName"),"'","''") & "','" &_
				replace(request("chrAREmail"),"'","''") & "','" &_
				replace(request("chrOSPerson"),"'","''") & "','" &_
				replace(request("chrOSEmail"),"'","''") & "','" &_
				replace(request("chrOSPhone"),"'","''") & "','" &_
				replace(request("chrOSFax"),"'","''") & "','" &_
				request("idSaveCarrier") & "','" &_
				replace(request("chrCarrier"),"'","''") & "','" &_
				replace(request("chrAccount"),"'","''") & "','" &_
				request("idPurpose") & "'," &_
				session("idUser") & ",'" &_
				replace(request("txtNotes"),"'","''") & "','" &_
				replace(request("txtShippingNotes"),"'","''") & "'"
		
			'execute and upload the information to SQL Server	
			set rsCart = dbConnection.Execute(sql)
			'This is the unique identifier for the Cart
			idCart = rsCart("idCart")
			
			'Close database connections
			rsCart.Close
			set rsCart = nothing
			'--------------------------------------------------------------------------------------------------
			'AEG - Send the Person putting in the order and email.
			'AEG - Find the user and email information and store it in a temp variable to send an email.
			set rsUser = server.CreateObject("adodb.recordset")
			sql = "execute FindUserbyID " & session("idUser")
			set rsUser = dbConnection.Execute(sql)
			'AEG - Load into temp variables
			chrUserName = trim(rsUser("chrFirst")) & " " & trim(rsUser("chrLast"))
			chrUserEmail = trim(rsUser("chrEmail"))
			'AEG - Close the user recordset
			rsUser.Close
			set rsUser = nothing
			'-------------------------------------------------------------------------
			'AEG - Open the Email Component
			Set Mailer = Server.CreateObject("SoftArtisans.SMTPMail") 'from www.softartisan.com
			Mailer.RemoteHost  = "techit-ex2.techitsolutions.com" 'mail server
			Mailer.FromName    = "administrator"
			Mailer.FromAddress = "administrator@techitsolutions.com"
			Mailer.AddRecipient chrUserName, chrUserEmail
			Mailer.Subject     = "Cart Created - " & replace(request("chrCart"),"'","''")
			Mailer.BodyText    = chrUserName & "," & VbCrLf & VbCrLf &_
			"Your New Cart: " & replace(request("chrOrder"),"'","''") & " has been created and if you haven't already done so, you may now add assets to this cart.  Once you have added all required assets, please complete the check out procedure.  Completeing this process will generate an email to your Pool Manager requesting approval of your Cart." & VbCrLf & VbCrLf &_
			"You will receive an email informing you of your Pool Manager's approval or disapproval." & VbCrLf & VbCrLf &_
			"If you have any challenges with the system, please contact your Pool Manager." & VbCrLf & VbCrLf &_
			"Thank You.." & VbCrLf &_
			"techIT Solutions Asset Management Team"
			'Execute the email
			Mailer.SendMail
			'---------------------------------------------------------------------------------		
			'Close the Database connections
			dbConnection.Close
			set dbConnection = nothing
			'-----------------------------------------------------------------------------------
			'now that the order has been placed we need to tell the user that the order has been placed and ready to go.
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Entry Successful</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
    <!-- #include file="includes/floatingtop.htm" -->
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr>
            <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td height="35"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td align="center"><font color="#0000FF" size="3" face="Arial, Helvetica, sans-serif"><strong>S U C C E S S ! !</strong></font></td>
                </tr>
                <tr> 
                  <td align="center"><p><font size="2" face="Arial, Helvetica, sans-serif">Your new cart has been created.</font></p>
                  <p><font size="2" face="Arial, Helvetica, sans-serif">Click here to <a href="ordering/loadcart.asp?idCart=<%=idCart%>">Add Assets</a> to this cart.</font></p>
                  <p><font size="2" face="Arial, Helvetica, sans-serif">or</font></p>
				  <p><a href="orders.asp">Return to Orders Page</a></font></p></td>
                </tr>
                <tr> 
                  <td height="50"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <!-- #include file="includes/floatingbottom.htm" -->
  </table>
</body>
</html>
<%
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
    <!-- #include file="includes/floatingtop.htm" -->
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
    <!-- #include file="includes/floatingbottom.htm" -->
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
    <!-- #include file="includes/floatingtop.htm" -->
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
    <!-- #include file="includes/floatingbottom.htm" -->
  </table>
</body>
</html>
<%
	end if
%>