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
	sql = "execute InsertLostStatus " & _
		request("idInventory") & "," &_
		request("idCart") & "," &_
		request("idWarehouse") & "," &_
		session("idUser") & ",'" &_
		replace(request("txtExplanation"),"'","''") & "'"
		
	'execute and upload the information to SQL Server	
	dbConnection.Execute(sql)
	
	'Close database connections
	dbConnection.Close
	set dbConnection = nothing
	
	Response.Redirect "lostsplash.asp"
%>