<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	if request("Cancel") = "Cancel Update" then
		Response.Redirect "orders.asp?idUser=" & request("idUser") & "&idCustomer=" & request("idCustomer") & "&idStatus=" & request("idRStatus") & "&idType=" & request("idType") & "&chrSearch=" & request("chrSearch")
	end if
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'check to see if we need to delete anything
	if cint(request("idSpecial")) = 1 and cint(request("intCounter")) > 0 then
		'remove the assets
		for i = 1 to cint(request("intCounter"))
			sql = "execute DeletefromCart " & request(i) & "," & request("idCart")
			dbConnection.execute(sql)
		'get the next assets
		next
	end if
	
	'Send the information string with contents from the form.
	sql = "execute UpdateCart2 " & _
		request("idCart") & ",'" &_
		request("idBilled") & "'," &_
		request("idLoaner") & "," &_
		request("idStatus") & ",'" &_
		replace(request("chrOrder"),"'","''") & "','" &_
		replace(request("chrDDName"),"'","''") & "','" &_
		replace(request("chrDDNumber"),"'","''") & "','" &_
		replace(request("chrManager"),"'","''") & "','" &_
		replace(request("chrIPerson"),"'","''") & "','" &_
		replace(request("chrIEmail"),"'","''") & "','" &_
		request("dtPull") & "','" &_
		request("dtShip") & "','" &_
		request("dtArrival") & "','" &_
		request("dtArrivalTime") & "','" &_
		request("dtDeparture") & "','" &_
		request("dtDepartureTime") & "','" &_
		request("dtReturn") & "','" &_
		request("dtTurn") & "','" &_
		replace(request("chrAddress"),"'","''") & "','" &_
		replace(request("chrAddress2"),"'","''") & "','" &_
		replace(request("chrAddress3"),"'","''") & "','" &_
		replace(request("chrAddress4"),"'","''") & "','" &_
		replace(request("chrCity"),"'","''") & "','" &_
		replace(request("chrState"),"'","''") & "','" &_
		replace(request("chrZip"),"'","''") & "','" &_
		replace(request("chrCountry"),"'","''") & "','" &_
		replace(request("chrARName"),"'","''") & "','" &_
		replace(request("chrAREmail"),"'","''") & "','" &_
		replace(request("chrOSPerson"),"'","''") & "','" &_
		replace(request("chrOSEmail"),"'","''") & "','" &_
		replace(request("chrOSPhone"),"'","''") & "','" &_
		replace(request("chrOSFax"),"'","''") & "','" &_
		replace(request("chrCarrier"),"'","''") & "','" &_
		replace(request("chrAccount"),"'","''") & "','" &_
		request("idPurpose") & "'," &_
		request("idExpedite") & ",'" &_
		replace(request("txtNotes"),"'","''") & "','" &_
		replace(request("txtShippingNotes"),"'","''") & "'"
		
		'execute and upload the information to SQL Server	
		dbConnection.Execute(sql)
		'Close the Database connections			
		dbConnection.Close
		set dbConnection = nothing
		'now that the order has been placed we need to tell the user that the order has been placed and ready to go.
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Update Successful</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="10" height="10" align="left" valign="top"><img src="images/topleftblue.gif" width="10" height="10"></td>
            <td width="780" height="10"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td><img src="images/eamlogo.gif" width="150" height="45"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                </tr>
              </table></td>
            <td width="10" height="10" align="right" valign="top"><img src="images/toprightblue.gif" width="10" height="10"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr>
            <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td height="35"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td align="center"><font color="#0000FF" size="3" face="Arial, Helvetica, sans-serif"><strong>S U C C E S S ! !</strong></font></td>
                </tr>
                <tr> 
                  <td align="center"><font size="2" face="Arial, Helvetica, sans-serif">Your Order or Cart has been updated as per your request.<br><BR>
					<a href="viewcart.asp?idCart=<%=request("idCart")%>">Return to Orders Page</a></font></td>
                </tr>
                <tr> 
                  <td height="50"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="10" height="10" align="left" valign="bottom"><img src="images/bottomleftblue.gif" width="10" height="10"></td>
            <td width="780" height="10"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td align="center"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Copyright &copy; 2004 
                    techIT Solutions LLC. <br>
                    Asset Management Enterprise Portal 4.0 &amp; Corporate Business 
                    Intelligence are products of techIT Solutions. </font></td>
                </tr>
              </table></td>
            <td width="10" height="10" align="right" valign="bottom"><img src="images/bottomrightblue.gif" width="10" height="10"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
    </tr>
  </table>
</body>
</html>
