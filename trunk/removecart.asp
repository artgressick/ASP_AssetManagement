<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
%>
<!-- #include file="includes/openconn.asp" -->
<%
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
	Set Mailer = Server.CreateObject("SoftArtisans.SMTPMail") 'from www.softartisans.com
	Mailer.RemoteHost  = "techit-ex2.techitsolutions.com" 'mail server
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
	'AEG - Start the Message information
	Mailer.Subject     = "Cart Disapproved - " & chrCart
	Mailer.BodyText    = chrUserName & "," & VbCrLf & VbCrLf &_
	"You cart, " & chrCart & " has been dissapproved by the Pool Manager." & VbCrLf & VbCrLf &_
	"Reason: " & VbCrLf &_
	replace(request("reason"),"'","''") & VbCrLf & VbCrLf &_
	"Thank you..." & VbCrLf &_
	"techIT Solutions Asset Management Team"
	
	'Execute the email
	Mailer.SendMail
	
	'Send the information string with contents from the form.
	sql = "execute DeleteCart " & _
		request("idCart")
		
	'execute and upload the information to SQL Server	
	dbConnection.Execute(sql)	

	'Close database connections
	dbConnection.Close
	set dbConnection = nothing
	
	Response.Redirect "approvecarts.asp"
%>