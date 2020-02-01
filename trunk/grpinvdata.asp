<!-- #include file="includes/openconn.asp" -->
<%
	'Get the Statistics
  set rsStats = server.CreateObject("adodb.recordset")
  if session("idAccess") < "O" then
 	sql = "execute StatisticsbyAdmin"
  else
	sql = "execute StatisticsbyUser " & session("idUser")
  end if
  set rsStats = dbConnection.Execute(sql)
	
  Response.ContentType="text/xml"
  'We specify the HTTP content type for the response object
  'to be text/xml. The content type tells the browser
  'what type of content to expect
  Dim strXML
  strXML= "<graph bgcolor='ffffff' " &_
  "xaxisname='Cumulative Inventory Report'" &_
  "yaxisname='Assets' " &_
  "caption='" & rsStats("intTotal") & " Total Assets' " &_
  "showgridbg='1'> " &_
  "<set name='Ready' value='" & rsStats("intReady") & "' color='FF0000' link='reportinventory.asp?idCustomer=0&idStatus=1&idWarehouse=0'/> " &_
  "<set name='Out' value='" & rsStats("intOut") & "' color='00FF00' link='reportinventory.asp?idCustomer=0&idStatus=2&idWarehouse=0'/> " &_
  "<set name='Lost' value='" & rsStats("intLost") & "' color='0000FF' link='reportinventory.asp?idCustomer=0&idStatus=9&idWarehouse=0'/> " &_
  "<set name='Damaged' value='" & rsStats("intBroken") & "' color='FFFF00' link='reportinventory.asp?idCustomer=0&idStatus=8&idWarehouse=0'/> " &_
  "<set name='Internal' value='" & rsStats("intInternal") & "' color='00FFFF' link='reportinventory.asp?idCustomer=0&idStatus=7&idWarehouse=0'/> " &_
  "<set name='Turning' value='" & rsStats("intTurning") & "' color='FF00FF' link='reportinventory.asp?idCustomer=0&idStatus=3&idWarehouse=0'/> " &_
  "<set name='Out of Sys' value='" & rsStats("intOutSystem") & "' color='0F0F0F' link='reportinventory.asp?idCustomer=0&idStatus=6&idWarehouse=0'/> " &_
  "</graph>"
  
  Response.Write(strXML)

  rsStats.Close
  set rsStats = nothing
  dbConnection.Close
  set dbConnection = nothing
%>
