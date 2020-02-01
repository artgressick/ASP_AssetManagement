<%
  'Send email for staged

	'AEG - Find the Cart Name so that we can send an Email
	'set rsCart = server.CreateObject("adodb.recordset")
	'sql = "execute FindCartNamebyID " & idCart
	'set rsCart = dbConnection.Execute(sql)
	
	'chrCart = trim(rsCart("chrCart"))
	
	'AEG - Close the Recordset
	'rsCart.Close
	'set rsCart = nothing
	
	'--------------------------------------------------------------------------
	'AEG - Open the SMTP Mailer Client
	'Set Mailer = Server.CreateObject("SoftArtisans.SMTPMail") 'from www.softartisan.com
	'Mailer.RemoteHost  = "63.236.44.26" 'mail server
	'Mailer.FromName    = "administrator"
	'Mailer.FromAddress = "administrator@techitsolutions.com"
	'--------------------------------------------------------------------------
	'AEG - Find the User to send an email.
	'set rsUser = server.CreateObject("adodb.recordset")
	'sql = "execute FindUserWhoEnteredCartbyCartID " & idCart
	'set rsUser = dbConnection.Execute(sql)
	
	'AEG - Move the records to temp fields
	'chrUserName = trim(rsUser("chrFirst")) & " " & trim(rsUser("chrLast"))
	'chrUserEmail = trim(rsUser("chrEmail"))
	
	'AEG - Attach the User to the email
	'Mailer.AddRecipient chrUserName, chrUserEmail
	
	'Close the User Recordset
	'rsUser.Close
	'set rsUser = nothing
	'--------------------------------------------------------------------------
	'AEG - Find the Pool Manager(s)
	'set rsPoolManager = server.CreateObject("adodb.recordset")
	'sql = "execute FindPoolManagersbyCart " & idCart
	'set rsPoolManager = dbConnection.Execute(sql)
	
	'AEG - Only do this if we have a record.
	'if not rsPoolManager.EOF then
	'	do until rsPoolManager.EOF
			'Move the information to the Temp Variables
	'		chrPoolManagerName = trim(rsPoolManager("chrFirst")) & " " & trim(rsPoolManager("chrLast"))
	'		chrPoolManagerEmail = trim(rsPoolManager("chrEmail"))
			
			'Attach the information to the email
	'		Mailer.AddRecipient chrPoolManagerName, chrPoolManagerEmail
			
	'	rsPoolManager.MoveNext
	'	loop
	'end if
	
	'Close the Pool Manager Recordset
	'rsPoolManager.Close
	'set rsPoolManager = nothing
	'--------------------------------------------------------------------------
	'AEG - Start the Message information
	'Mailer.Subject     = "Testing - Cart Staged for Shipping - " & chrCart
	'Mailer.BodyText    = chrUserName & "," & VbCrLf & VbCrLf &_
	'"You cart, " & chrCart & " has been staged for shipping." & VbCrLf & VbCrLf &_
	'"When the carrier arrives at the warehouse we will receive an email with the tracking information. An additional email will be sent to the on-site contact when the cart has been picked up by the carrier." & VbCrLf & VbCrLf &_
	'"Thank you again for using the techIT Solutions Asset Management System." & VbCrLf & VbCrLf &_
	'"techIT Solutions Asset Management Team."
	
	'Execute the email
	'Mailer.SendMail	
%>

<%
    'email for shipping.
	'AEG - Find the Cart Name so that we can send an Email
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute FindCartNamebyID " & request("idCart")
	set rsCart = dbConnection.Execute(sql)
	
	chrCart = trim(rsCart("chrCart"))
	
	'AEG - Close the Recordset
	rsCart.Close
	set rsCart = nothing
	
	'--------------------------------------------------------------------------
	'AEG - Open the SMTP Mailer Client
	Set Mailer = Server.CreateObject("SoftArtisans.SMTPMail") 'from www.softartisan.com
	Mailer.RemoteHost  = "63.236.44.26" 'mail server
	Mailer.FromName    = "administrator"
	Mailer.FromAddress = "administrator@techitsolutions.com"
	'--------------------------------------------------------------------------
	'AEG - Find the User to send an email.
	set rsUser = server.CreateObject("adodb.recordset")
	sql = "execute FindUserWhoEnteredCartbyCartID " & request("idCart")
	set rsUser = dbConnection.Execute(sql)
	
	'AEG - Move the records to temp fields
	chrUserName = trim(rsUser("chrFirst")) & " " & trim(rsUser("chrLast"))
	chrUserEmail = trim(rsUser("chrEmail"))
	
	'AEG - Attach the User to the email
	Mailer.AddRecipient chrUserName, chrUserEmail
	
	'Close the User Recordset
	rsUser.Close
	set rsUser = nothing
	'--------------------------------------------------------------------------
	'AEG - Find the OnSite Contact to send an email.
	set rsOnSite = server.CreateObject("adodb.recordset")
	sql = "execute FindOnSiteContactInformationbyCartID " & request("idCart")
	set rsOnSite = dbConnection.Execute(sql)
	
	'AEG - Move the records to temp fields
	chrOnSiteName = trim(rsOnSite("chrOSPerson"))
	chrOnSiteEmail = trim(rsOnSite("chrOSEmail"))
	
	'AEG - Attach the User to the email
	Mailer.AddRecipient chrOnSiteName, chrOnSiteEmail
	
	'Close the User Recordset
	rsOnSite.Close
	set rsOnSite = nothing
	'--------------------------------------------------------------------------
	'AEG - Find the Pool Manager(s)
	set rsPoolManager = server.CreateObject("adodb.recordset")
	sql = "execute FindPoolManagersbyCart " & request("idCart")
	set rsPoolManager = dbConnection.Execute(sql)
	
	'AEG - Only do this if we have a record.
	if not rsPoolManager.EOF then
		do until rsPoolManager.EOF
			'Move the information to the Temp Variables
			chrPoolManagerName = trim(rsPoolManager("chrFirst")) & " " & trim(rsPoolManager("chrLast"))
			chrPoolManagerEmail = trim(rsPoolManager("chrEmail"))
			
			'Attach the information to the email
			Mailer.AddRecipient chrPoolManagerName, chrPoolManagerEmail
			
		rsPoolManager.MoveNext
		loop
	end if
	
	'Close the Pool Manager Recordset
	rsPoolManager.Close
	set rsPoolManager = nothing
	'--------------------------------------------------------------------------
	'AEG - Start the Message information
	Mailer.Subject     = "Testing - Cart Shipped - " & chrCart
	Mailer.BodyText    = chrUserName & "," & VbCrLf & VbCrLf &_
	"You cart, " & chrCart & " has been shipped." & VbCrLf & VbCrLf &_
	"When the carrier arrives at the warehouse we will receive an email with the tracking information. An additional email will be sent to the on-site contact when the cart has been picked up by the carrier." & VbCrLf & VbCrLf &_
	"Thank you again for using the techIT Solutions Asset Management System." & VbCrLf & VbCrLf &_
	"techIT Solutions Asset Management Team."
	
	'Execute the email
	Mailer.SendMail	
%>

<%
	'Email for Pool manager
	'AEG - Find the Pool Manager(s)
	set rsPoolManager = server.CreateObject("adodb.recordset")
	sql = "execute FindPoolManagersbyCart " & request("idCart")
	set rsPoolManager = dbConnection.Execute(sql)
	
	'AEG - Only do this if we have a record.
	if not rsPoolManager.EOF then
		do until rsPoolManager.EOF
			'Move the information to the Temp Variables
			chrPoolManagerName = trim(rsPoolManager("chrFirst")) & " " & trim(rsPoolManager("chrLast"))
			chrPoolManagerEmail = trim(rsPoolManager("chrEmail"))
			
			'Attach the information to the email
			Mailer.AddRecipient chrPoolManagerName, chrPoolManagerEmail
			
		rsPoolManager.MoveNext
		loop
	end if
	
	'Close the Pool Manager Recordset
	rsPoolManager.Close
	set rsPoolManager = nothing

%>