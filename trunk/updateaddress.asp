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
	sql = "execute UpdateSavedAddress " & _
		request("idAddress") & ",'" &_
		replace(request("chrSavedAddressName"),"'","''") & "','" &_
		replace(request("chrAddress"),"'","''") & "','" &_
		replace(request("chrAddress2"),"'","''") & "','" &_
		replace(request("chrAddress3"),"'","''") & "','" &_
		replace(request("chrAddress4"),"'","''") & "','" &_
		replace(request("chrCity"),"'","''") & "','" &_
		replace(request("chrState"),"'","''") & "','" &_
		replace(request("chrZip"),"'","''") & "','" &_
		replace(request("chrCountry"),"'","''") & "'"
		
	'execute and upload the information to SQL Server	
	dbConnection.Execute(sql)
	
	'Close database connections
	dbConnection.Close
	set dbConnection = nothing
	
	Response.Redirect "savedaddresses.asp"
%>