<!-- #include file="../includes/openconn.asp" -->
<%
	'List the inventory
	set rsInventory = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute WhatsAvailable " & request("idCustomer") & "," & request("idCategory") & ",'" & dateadd("d",-3,request("dtPull")) & "','" & dateadd("d",5,request("dtTurn")) & "'"
	else
		sql = "execute WhatsAvailablebyUser " & session("idUser") & "," & request("idCustomer") & "," & request("idCategory") & ",'" & dateadd("d",-3,request("dtPull")) & "','" & dateadd("d",5,request("dtTurn")) & "'"
	end if
	set rsInventory = dbConnection.Execute(sql)
	
	'Name for the ouput document
	file_being_created= "WhatsAvailable.xls"
	
	'create a file system object
	set fso = createobject("scripting.filesystemobject")
	
	'create the text file  - true will overwrite any previous files
	'Writes the db output to a .xls file in the same directory
	Set act = fso.CreateTextFile(server.mappath(file_being_created), true)
	
	'All non repetitive html on top goes here
	act.WriteLine("<html><body>")
	act.WriteLine("<table border=""1"">")
	act.WriteLine("<tr>")
	act.WriteLine("<th nowrap>Asset Numebr</th>")
	act.WriteLine("<th nowrap>Serial Number</th>")
	act.WriteLine("<th nowrap>Item</th>")
	act.WriteLine("<th nowrap>Current Location</th>")
	act.WriteLine("</tr>")
	
	'For net loop to create seven word documents from the record set
	'change this to "do while not rs.eof" to output all the records
	'and the corresponsding next should be changed to loop also
	if not rsInventory.EOF then
		do until rsInventory.EOF
			Act.WriteLine("<tr>")
			act.WriteLine("<td align=""right"">" & trim(rsInventory("chrAssNum")) & "</td>" )
			act.WriteLine("<td align=""right"">" & trim(rsInventory("chrSerialNum")) & "</td>" )
			if rsInventory("chrType") = "C" then
				act.WriteLine("<td align=""right"">" & trim(rsInventory("chrItem")) & " - " & trim(rsInventory("chrProcessor")) & ", " & trim(rsInventory("chrMemory")) & " - " & trim(rsInventory("chrODrive")) & " - " & trim(rsInventory("chrHDD")) & "</td>" )
			else
				act.WriteLine("<td align=""right"">" & trim(rsInventory("chrItem")) & "</td>" )
			end if
			if rsInventory("idCart") = 0 then
				act.WriteLine("<td align=""right"">Warehouse</td>" )
			else
				act.WriteLine("<td align=""right"">" & trim(rsInventory("chrCart")) & "</td>" )
			end if
			act.WriteLine("</tr>")
		'move to the next record
		rsInventory.MoveNext
		'return to the top of the for - next loop
		'change this to "loop" to output all the records
		'and the corresponsding for statement above should be changed also
		Loop
	End If
	'All non repetitive html on top goes here
	act.WriteLine("</table></body></html>")
	
	'close the object (excel)
	act.close
	rsInventory.Close
	set rsInventory = Nothing
	dbConnection.Close
	set dbConnection = nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Export to Excel</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#FF0000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="300" height="275" border="0" cellpadding="7" cellspacing="0">
  <tr>
    <td align="center"> <p><font color="#0000FF" size="3" face="Arial, Helvetica, sans-serif"><strong>Export
        to Excel</strong></font></p></td>
  </tr>
  <tr>
    <td><p><font size="2" face="Arial, Helvetica, sans-serif"><strong>Directions</strong>:
        To view the Excel document you must download it to your desktop. To do
        this please follow the directions below.</font></p>
      <p><font size="2" face="Arial, Helvetica, sans-serif"><strong>Using a 2-button
        mouse</strong>: Using the right mouse button right-click on the link below
        and &quot;Download Link to Disk&quot;.</font></p>
      <p><font size="2" face="Arial, Helvetica, sans-serif"><strong>Using a 1-button
        mouse</strong>: Holding down the &quot;Control&quot; key click on the
        link below and &quot;Download Link to Disk&quot;.</font></p></td>
  </tr>
  <tr>
    <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Excel
      File: <a href="WhatsAvailable.xls">What's Available Report</a></strong></font></td>
  </tr>
</table>
</body>
</html>
