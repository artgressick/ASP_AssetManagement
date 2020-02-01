<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'if this is an existing user then
	if request("idType") = 0 then
		'check the accounts table to make sure we don't create duplicate accounts
		set rsAccount = server.CreateObject("adodb.recordset")
		sql = "execute CheckDupAccount " & request("idUser") & "," & request("idCustomer") & "," & request("idEAccess")
		set rsAccount = dbConnection.Execute(sql)
		'if nothing comes back then we can add the account
		if rsAccount.eof then
			sql = "execute InsertAccount " &_
				request("idUser") & "," &_
				request("idCustomer") & "," &_
				request("idEAccess")
			dbConnection.Execute(sql)
		end if
		rsAccount.Close
		set rsAccount = nothing
	else
		'they decided to create a new account
		'check to make sure that the User doesn't already exist. Send it the Email address
		set rsAccount = server.CreateObject("adodb.recordset")
		sql = "execute CheckDupUser '" & replace(request("chrEmail"),"'","''") & "'"
		set rsAccount = dbConnection.Execute(sql)
		'if nothing comes back then we can create the account
		if rsAccount.eof then
			'set the password and status
			idStatus = 1
			chrPassword = "techit"
			sql = "execute InsertUserandAccount " &_
				request("idCustomer") & ",'" &_
				request("idNAccess") & "'," &_
				idStatus & ",'" &_
				replace(request("chrEmail"),"'","''") & "','" &_
				chrPassword & "','" &_
				replace(request("chrFirst"),"'","''") & "','" &_
				replace(request("chrLast"),"'","''") & "','" &_
				replace(request("chrPhone"),"'","''") & "','" &_
				replace(request("chrFax"),"'","''") & "'"
			dbConnection.Execute(sql)
			'---------------------------------------------------------------------------------------
			'combine information
			chrUsername = replace(request("chrFirst"),"'","''") & " " & replace(request("chrLast"),"'","''")
			chrUserEmail = trim(request("chrEmail"))
			'send an email to the user
			Set Mailer = Server.CreateObject("SoftArtisans.SMTPMail") 'from www.softartisan.com
			
			Mailer.RemoteHost  = "techit-ex2.techitsolutions.com" 'mail server
			Mailer.FromName    = "administrator"
			Mailer.FromAddress = "administrator@techitsolutions.com"
			Mailer.AddRecipient chrUserName, chrUserEmail
			Mailer.Subject     = "Asset Management - Account Created"
			Mailer.BodyText    = "Hello," & VbCrLf & VbCrLf &_
			"An account has been created for you. Your username and password is listed below." & VbCrLf & VbCrLf &_
			"Username: " & chrUserEmail & VbCrLf &_
			"Password: techit" & VbCrLf &_
			"Website: http://www.itechit.com" & VbCrLf & VbCrLf &_
			"If you have any questions please contact Customer Service at 1-800-492-2448 or email us at customerservice@techitsolutions.com"
			'Execute the email
			Mailer.SendMail
		else
			if rsAccount("idStatus") = 0 then
				'we need to reactivate the users account instead of creating a new one.
				sql = "execute ActivateUser " & rsAccount("idUser")
				dbConnection.Execute(sql)
				'create the new account
				sql = "execute InsertAccount " &_
					rsAccount("idUser") & "," &_
					request("idCustomer") & "," &_
					request("idNAccess")
				dbConnection.Execute(sql)
			else
				'the use exists and is active we need to add the account. make sure to check for duplicate accounts also.
				set rsAccess = server.CreateObject("adodb.recordset")
				sql = "execute CheckDupAccount " & rsAccount("idUser") & "," & request("idCustomer") & "," & request("idNAccess")
				set rsAccess = dbConnection.Execute(sql)
				if rsAccess.eof then
					'create the new account
					sql = "execute InsertAccount " &_
						rsAccount("idUser") & "," &_
						request("idCustomer") & "," &_
						request("idNAccess")
					dbConnection.Execute(sql)
				end if
				'close the Access recordset
				rsAccess.close
				set rsAccess = nothing
			end if
		end if
		'close the Account recordset
		rsAccount.Close
		set rsAccount = nothing
	end if		
		
	'Close database connections
	dbConnection.Close
	set dbConnection = nothing
	
	Response.Redirect "settings.asp?idCustomer=" & request("idCustomer")
%>