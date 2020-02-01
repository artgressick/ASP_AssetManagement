<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if

	'AEG - Check to see which button was submitted
	'AEG - Submit 1 is cancel and Submit 2 is approval
	if request("submit1") = "" then
%>
<!-- #include file="includes/openconn.asp" -->
<%
		'Send the information string with contents from the form.
		sql = "execute DeleteCustomer " & _
			request("idCustomer")
		
		'execute and upload the information to SQL Server
		dbConnection.Execute(sql)
		
		'Close database connections
		dbConnection.Close
		set dbConnection = nothing
	end if
	
	Response.Redirect "customers.asp"
%>