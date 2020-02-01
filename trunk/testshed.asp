<%@ Language=VBScript %>
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
	Set Mailer = Server.CreateObject("SoftArtisans.SMTPMail") 'from www.softartisan.com
	Mailer.RemoteHost  = "63.236.44.26" 'mail server
	Mailer.FromName    = "administrator"
	Mailer.FromAddress = "administrator@techitsolutions.com"
	'--------------------------------------------------------------------------
	'AEG - Find the User to send an email.
	'set rsUser = server.CreateObject("adodb.recordset")
	'sql = "execute FindUserWhoEnteredCartbyCartID " & idCart
	'set rsUser = dbConnection.Execute(sql)
	
	'AEG - Move the records to temp fields
	chrUserName = "Bob Forringer"
	chrUserEmail = "wiztrkii@yahoo.com"
	
	'AEG - Attach the User to the email
	Mailer.AddRecipient chrUserName, chrUserEmail
	
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
	Mailer.Subject     = "Testing Windows Script"
	Mailer.BodyText    = "This is a test of the windows script schedual to run reports everyday"
	
	'Execute the email
	Mailer.SendMail	
%>
