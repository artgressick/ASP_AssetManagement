<%@ Language=VBScript %>
<%
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
%>
<!-- #include file="includes/openconn.asp" -->
<%	
	'Get the Customer Name
	set rsAccount = server.CreateObject("adodb.recordset")
	sql = "execute FindUserAccessbyID " & request("idAccess")
	set rsAccount = dbConnection.Execute(sql)
	
	'load variables --------------------------------------
	session("idAccess") = rsAccount("idAccess")
	session("chrAccess") = rsAccount("chrAccess")
	session("chrAbility") = rsAccount("chrAbility")
	idAccess = rsAccount("idAccess")
	'-----------------------------------------------------
	
	rsAccount.Close
	set rsAccount = nothing
	dbConnection.Close
	set dbConnection = nothing
	
	if idAccess = "R" then
		Response.Redirect "limited/"
	else
		Response.Redirect "default.asp"
	end if
%>