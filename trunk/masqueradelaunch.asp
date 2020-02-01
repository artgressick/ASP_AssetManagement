<%@ Language=VBScript %>
<%
	session("idUser") = ""
	'session.Abandon
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Find your profile
	set rsProfile = server.CreateObject("adodb.recordset")
    sql = "execute FindProfilebyIDUser " & Request("idUser")
	set rsProfile = dbConnection.Execute(sql)
	
	'get the session variable loaded
	session("idUser") = rsProfile("idUser")
	session("chrFirst") = trim(rsProfile("chrFirst"))
	session("chrLast") = trim(rsProfile("chrLast"))
	
	'redirect to the choose account pages.
	Response.Redirect "chooseaccount.asp"
%>