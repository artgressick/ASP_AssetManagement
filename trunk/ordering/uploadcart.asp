<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "../logoff.asp"
	end if
%>
<!-- #include file="../includes/openconn.asp" -->
<%
	'get the list of Assets
	set rsAssets = server.CreateObject("adodb.recordset")
	sql = "execute ListAvailablebyDescriptionandCustomer " & request("idDescription") & "," & session("idCart") & "," & request("idCustomer")
	set rsAssets = dbConnection.Execute(sql)
	
	'upload the amount requested
	counter = 1
	intQuantity = cint(request("intQuantity"))
	'make sure the Assets didn't disappear and cause an error
	if not rsAssets.EOF then
		do until rsAssets.EOF or counter > intQuantity
			'check for duplicates
			set rsChecker = server.CreateObject("adodb.recordset")
			sql = "execute CheckFutureOrders " & rsAssets("idInventory") & "," & session("idCart")
			set rsChecker = dbConnection.Execute(sql)
			'only insert records that aren't duplicates
			if rsChecker.EOF then
				sql = "execute InsertAssetsintoCart " & _
					rsAssets("idInventory") & "," &_
					session("idCart") & "," &_
					session("idUser")
				'execute the upload
				dbConnection.Execute(sql)
			'from the checker statement
			end if
			'close the checker connection
			rsChecker.Close
			set rsChecker = nothing
			'move to the next record.
			rsAssets.MoveNext
			counter = counter + 1
		loop
	end if
	
	'Close database connections
	rsAssets.Close
	set rsAssets = nothing
	dbConnection.Close
	set dbConnection = nothing
	
	'send the use back to the Categories page of the cart.
	Response.Redirect "default.asp?idCategory=" & request("idCategory")
%>
