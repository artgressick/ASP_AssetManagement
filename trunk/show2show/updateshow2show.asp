<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "../logoff.asp"
	end if
%>
<!-- #include file="../includes/openconn.asp" -->
<%	
	'Find the Cart information
	sql = "execute UpdateShow2Show " & request("idCart") & "," & request("idLinkedCart")
	dbConnection.Execute(sql)
	
	dbConnection.Close
	set dbConnection = nothing
	
	'load the cart into the session variables
	session("idCart") = request("idCart")
	session("idLinkedCart") = request("idLinkedCart")
	
	'send the users to the cart page
	Response.Redirect "default.asp"
	
	
%>