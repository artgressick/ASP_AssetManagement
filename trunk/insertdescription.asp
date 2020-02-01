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
	sql = "execute InsertDescription " & _
		request("idCategory") & "," &_
		request("idPicture") & ",'" &_
		replace(request("chrItemNo"),"'","''") & "','" &_
		replace(request("chrItem"),"'","''") & "','" &_
		replace(request("chrProcessor"),"'","''") & "','" &_
		replace(request("chrMemory"),"'","''") & "','" &_
		replace(request("chrHDD"),"'","''") & "','" &_
		replace(request("chrODrive"),"'","''") & "','" &_
		replace(request("chrRStorage"),"'","''") & "','" &_
		replace(request("chrSCSI"),"'","''") & "','" &_
		replace(request("chrGraphics"),"'","''") & "','" &_
		replace(request("chrWireless"),"'","''") & "','" &_
		replace(request("chrBluetooth"),"'","''") & "','" &_
		replace(request("chrModem"),"'","''") & "','" &_
		replace(request("chrUSB"),"'","''") & "','" &_
		replace(request("chrFireWire"),"'","''") & "','" &_
		replace(request("chrEthernet"),"'","''") & "','" &_
		replace(request("chrOS"),"'","''") & "'"
		
	'execute and upload the information to SQL Server	
	dbConnection.Execute(sql)
	
	'Close database connections
	dbConnection.Close
	set dbConnection = nothing
	
	Response.Redirect "inventory.asp"
%>