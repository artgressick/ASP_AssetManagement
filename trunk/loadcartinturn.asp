<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if

	'load the cart into a session
	session("idLoadInW") = request("idWarehouse")
	
	
	Response.Redirect "checkinturn.asp"
%>