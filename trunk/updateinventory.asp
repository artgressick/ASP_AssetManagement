<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'build a new asset number
	idCustomer = right("000" & request("idCustomer"), 3)
	idDescription = right("000" & request("idDescription"),3)
	idInventory = right("00000" & request("idInventory"),5)
	
	chrAssNum = idCustomer & "-" & idDescription & "-" & idInventory
	
	'Send the information string with contents from the form.
	sql = "execute UpdateInventory " &_
		request("idInventory") & "," &_
		request("idWarehouse") & "," &_
		request("idCustomer") & "," &_
		request("idDescription") & "," &_
		request("idOwner") & ",'" &_
		chrAssNum & "','" &_
		replace(request("chrSerialNum"),"'","''") & "','" & _
		replace(request("txtNotes"),"'","''") & "'"
		
	'execute and upload the information to SQL Server	
	dbConnection.Execute(sql)
	
	'Close database connections
	dbConnection.Close
	set dbConnection = nothing
	
	Response.Redirect "viewdescription.asp?idDescription=" & request("idDescription")
%>