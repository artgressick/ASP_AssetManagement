<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Send the information string with contents from the form.
	sql = "execute UpdateOrder " &_
		request("idOrder") & ",'" &_
		replace(request("chrOrder"),"'","''") & "'"
		
	'execute and upload the information to SQL Server	
	dbConnection.Execute(sql)
	
	'Close database connections
	dbConnection.Close
	set dbConnection = nothing
	
	Response.Redirect "orders.asp?idCustomer=" & request("idCustomer") & "&idStatus=" & request("idStatus") & "&idUser=" & request("idUser") & "&idType=" & request("idType") & "&chrSearch=" & request("chrSearch")
%>