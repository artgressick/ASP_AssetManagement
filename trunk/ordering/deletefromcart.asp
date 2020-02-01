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
	sql = "execute ListAssetsinCartbyDescriptionandCart " & request("idDescription") & "," & session("idCart")
	set rsAssets = dbConnection.Execute(sql)
	
	'upload the amount requested
	counter = 1
	intQuantity = cint(request("intQuantity"))
	'make sure the Assets didn't disappear and cause an error
	if not rsAssets.EOF then
		do until rsAssets.EOF or counter > intQuantity
			'remove the assets from the cart and ordered tables
			sql = "execute DeletefromCart " & _
				rsAssets("idInventory") & "," &_
				session("idCart")
			'execute the upload
			dbConnection.Execute(sql)
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
