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
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute FindCartbyID " & request("idCart")
	set rsCart = dbConnection.Execute(sql)
	
	if cint(rsCart("idShow2Show")) = True then
		Response.Redirect "../show2show/loadcart.asp?idCart=" & request("idCart")
	else
		session("idCart") = request("idCart")
		Response.Redirect "default.asp"
	end if	
%>