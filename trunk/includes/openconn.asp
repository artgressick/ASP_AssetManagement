<%
	'open connection
	set dbConnection = server.CreateObject("adodb.Connection")
	dbConnection.Open "AssetManagement", "WebServer","PASSWORD"
%>