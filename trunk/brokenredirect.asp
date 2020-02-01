<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if

	'AEG - Check to see which button was submitted
	'AEG - Submit 1 is cancel and Submit 2 is approval
	if request("submit1") = "" then
		Response.Redirect "checkinbroken.asp"
	else
		Response.Redirect "checkoutbroken.asp"
	end if
%>