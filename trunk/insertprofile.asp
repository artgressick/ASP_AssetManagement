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
	sql = "execute  InsertUser " & _
		1 & ",'" &_
		replace(request("chrEmail"),"'","''") & "','" &_
		replace(request("chrPassword"),"'","''") & "','" &_
		replace(request("chrFirst"),"'","''") & "','" &_
		replace(request("chrLast"),"'","''") & "','" &_
		replace(request("chrPhone"),"'","''") & "','" &_
		replace(request("chrFax"),"'","''") & "'"
		

	'execute and upload the information to SQL Server	
	dbConnection.Execute(sql)
	
	'Close database connections
	dbConnection.Close
	set dbConnection = nothing
	
	'RRF - Detemine Redirect by usrSession from Editprofile.asp
	Response.Redirect "profiles.asp"
%>
