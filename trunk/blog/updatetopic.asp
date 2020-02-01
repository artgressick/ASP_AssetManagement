<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "../logoff.asp"
	end if
%>
<!-- #include file="../includes/openconn.asp" -->
<%
	'Send the information string with contents from the form.
	sql = "execute UpdateTopic " & _
		request("idBlog") & "," &_
		request("idStatus") & "," &_
		request("idType") & "," &_
		request("idPriority") & ",'" &_
		replace(request("chrTitle"),"'","''") & "','" &_
		replace(request("txtMessage"),"'","''") & "'"
		
	'execute and upload the information to SQL Server	
	dbConnection.Execute(sql)
	
	'Close database connections
	dbConnection.Close
	set dbConnection = nothing
	
	Response.Redirect "default.asp"
%>