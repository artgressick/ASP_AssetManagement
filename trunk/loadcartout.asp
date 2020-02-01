<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Cart Status to Pulling Equipment
	idStatus = 4

	'Update the Cart status to Pulling Equipment
	sql = "execute UpdateCartApproval " & _
		request("idCart") & "," &_
		idStatus
		
	'execute and upload the information to SQL Server	
	dbConnection.Execute(sql)
	
	'Close database connections
	dbConnection.Close
	set dbConnection = nothing
	
	'load the cart into a session
	session("idLoadOut") = request("idCart")
	
	
	Response.Redirect "checkout.asp"
%>