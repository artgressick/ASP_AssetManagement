<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "../logoff.asp"
	end if
	
	'load the cart into the session variables
	session("idCart") = request("idCart")
	
	'send the users to the cart page
	Response.Redirect "default.asp"
%>