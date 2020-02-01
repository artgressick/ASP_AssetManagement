<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "../logoff.asp"
	end if
%>
<!-- #include file="../includes/openconn.asp" -->
<%
	'go through the checkboxes 1 by 1
	for i = 1 to cint(request("counter"))
		'check to see if the box was checked
		if len(request(i)) <> 0 then
			'check for duplicate records
			set rsChecker = server.CreateObject("adodb.recordset")
			sql = "execute CheckFutureOrders " & request(i) & "," & session("idCart")
			set rsChecker = dbConnection.Execute(sql)
			if rsChecker.eof then
				'insert the record into the Future Orders and Ordering tables
				sql = "execute InsertAssetsintoCart " & _
					request(i) & "," &_
					session("idCart") & "," &_
					session("idUser")
				'execute the upload
				dbConnection.Execute(sql)
			'from the checker
			end if
			rsChecker.Close
			set rsChecker = nothing
		'from the if length/checked
		end if
	next
	
	'close the connection
	dbConnection.Close
	set dbConnection = nothing
	
	'send the use back to the Categories page of the cart.
	Response.Redirect "default.asp?idCategory=" & request("idCategory")
%>
