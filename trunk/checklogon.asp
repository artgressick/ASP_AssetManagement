<%@ Language=VBScript %>
<!-- #include file="includes/openconn.asp" -->
<%
	'----------------------------------------------------------------------------------------------
	'run the Order automation and email procedures
	
	'run the automation SP
	'sql = "execute AutomationSP"
	'dbConnection.Execute(sql)
		
	'Run the email notification Program
	'set rsEmail = server.CreateObject("adodb.recordset")
	'sql = "execute CheckEmail5DaysPastReturn"
	'set rsEmail = dbConnection.Execute(sql)
	'Send the email notification if not already sent
	'if rsEmail.EOF then
		'start the notification recordset
	'	set rsNotification = server.CreateObject("adodb.recordset")
	'	sql = "execute Orders5DaysPastReturn"
	'	set rsNotification = dbConnection.Execute(sql)
	'	'Only send an email if there are shows outstanding
	'	if not rsNotification.EOF then
	'		'------------------------
	'		'send email techit
	'		Set mailer = Server.CreateObject("SoftArtisans.SMTPMail")
	'		'load the message
	'		chrHeader = "The following Orders are 5 days past their return date. Please check these orders for missing items and billing status." & VbCrLf & VbCrLf
	'		do until rsNotification.EOF
	'			chrBody = chrBody & "Order: " & trim(rsNotification("chrOrder")) & ";Return Date: " & formatdatetime(rsNotification("dtReturn"),short) & VbCrLf &_
	'		rsNotification.MoveNext
	'		loop
	'		'------------------------
	'		rem change this RemoteHost to a valid SMTP address before testing
	'		Mailer.RemoteHost  = "63.236.44.26"
	'		Mailer.FromName    = "Insight Automated Emailer"
	'		Mailer.FromAddress = "administrator@techitsolutions.com"
	'		Mailer.AddRecipient "Arthur Gressick", "arthurg@techitsolutions.com"
	'		Mailer.AddRecipient "Lesa Preston", "lpreston@techitsolutions.com"
	'		Mailer.AddRecipient "Warehouse", "warehouse@techitsolutions.com"
	'		Mailer.AddRecipient "Mark Preston", "mpreston@techitsolutions.com"
	'		Mailer.AddRecipient "Robert Kite", "rkite@techitsolutions.com"
	'		Mailer.Subject     = "Orders 5 Days Past Return"
	'		Mailer.BodyText    = chrHeader & chrBody
	'		Mailer.SendMail
	'		'------------------------
	'		'enter a record for the email that was sent
	'		sql = "insert into tblEmail(" & _
	'		"idType) " & _
	'		"values (" & _
	'		1 & ")"
	'		dbConnection.Execute(sql)
	'	end if
	'end if
	'close the recordset
	'rsEmail.Close
	'set rsEmail = nothing
	
	'----------------------------------------------------------------------------------------------	
	'check the database for the Account information. Open the record set and run the stored
	'procedure.
	set rsLogon = server.CreateObject("adodb.recordset")
	sql = "execute CheckLogon '" & request("chrUsername") & "','" & request("chrPassword") & "'"
	set rsLogon = dbConnection.Execute(sql)
	
	if rsLogon.EOF then
		'no account -------------------------------------------------------------------------------
		linkpage = "logonerror.asp"
	else
		if rsLogon("idStatus") = "0" then
			'their account has been diabled -------------------------------------------------------
			linkpage = "noaccess.asp"
		else		
			'account located load user information ------------------------------------------------
			session("idUser") = rsLogon("idUser")
			session("chrFirst") = trim(rsLogon("chrFirst"))
			session("chrLast") = trim(rsLogon("chrLast"))
			'--------------------------------------------------------------------------------------
			'Goto the Accounts Page
			linkpage =  "chooseaccount.asp"
		end if
	end if
	
	'----------------------------------------------------------------------------------------------
	'close all of the database connections.
	rsLogon.Close
	set rsLogon = nothing
	dbConnection.Close
	set dbConnection = nothing
	
	
	'----------------------------------------------------------------------------------------------
	'Write the cookie to the browser if they asked to remember
	if request("chrRemember") = "Yes" then
		Response.Cookies.Item("chrUsername") = Request("chrUsername")
		Response.Cookies.Item("chrUsername").Expires = #December 31, 2005#
	end if
	
	'----------------------------------------------------------------------------------------------
	'redirect the user to the approiate page
	Response.Redirect linkpage
%>