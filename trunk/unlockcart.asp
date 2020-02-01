<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'set cart approval
	idStatus = 1
	
	'Send the information string with contents from the form.
	sql = "execute UpdateCartApproval " & _
		request("idCart") & "," &_
		idStatus
		
	'execute and upload the information to SQL Server	
	dbConnection.Execute(sql)
	
	'AEG - Find the Cart Name so that we can send an Email
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute FindCartNamebyID " & request("idCart")
	set rsCart = dbConnection.Execute(sql)
	
	chrCart = trim(rsCart("chrCart"))
	
	'AEG - Close the Recordset
	rsCart.Close
	set rsCart = nothing
	
	'Close database connections
	dbConnection.Close
	set dbConnection = nothing
	
	Response.Redirect "retreivecart.asp?idOrder=" & request("idOrder")
%>