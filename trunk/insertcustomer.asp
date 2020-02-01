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
	sql = "execute InsertCustomer " & _
		replace(request("idCustomer"),"'","''") & ",'" &_
		replace(request("chrCustomer"),"'","''") & "','" &_
		replace(request("chrCAddress"),"'","''") & "','" &_
		replace(request("chrCAddress2"),"'","''") & "','" &_
		replace(request("chrCCity"),"'","''") & "','" &_
		replace(request("chrCState"),"'","''") & "','" &_
		replace(request("chrCZip"),"'","''") & "','" &_
		replace(request("chrCName"),"'","''") & "','" &_
		replace(request("chrCPhone"),"'","''") & "','" &_
		replace(request("chrCFax"),"'","''") & "','" &_
		replace(request("chrCEmail"),"'","''") & "','" &_
		replace(request("chrBAddress"),"'","''") & "','" &_
		replace(request("chrBAddress2"),"'","''") & "','" &_
		replace(request("chrBCity"),"'","''") & "','" &_
		replace(request("chrBState"),"'","''") & "','" &_
		replace(request("chrBZip"),"'","''") & "','" &_
		replace(request("chrBName"),"'","''") & "','" &_
		replace(request("chrBPhone"),"'","''") & "','" &_
		replace(request("chrBFax"),"'","''") & "','" &_
		replace(request("chrBEmail"),"'","''") & "'"
		
	'execute and upload the information to SQL Server	
	dbConnection.Execute(sql)
	
	'Close database connections
	dbConnection.Close
	set dbConnection = nothing
	
	Response.Redirect "customers.asp"
%>