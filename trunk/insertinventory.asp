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
	sql = "execute InsertInventory " &_
		request("idWarehouse") & "," &_
		request("idCustomer") & "," &_
		request("idDescription") & "," &_
		request("idOwner") & ",'" &_
		replace(request("chrSerialNum"),"'","''") & "'"
		
	'execute and upload the information to SQL Server
	dbConnection.Execute(sql)
		
	'Find out what the Asset ID number was.
	set rsInventory = server.CreateObject("adodb.recordset")
	sql = "select @@identity as idInventory"
	set rsInventory = dbConnection.Execute(sql)
	
	idInventory = rsInventory("idInventory")
	
	idCustomer = right("000" & request("idCustomer"), 3)
	idDescription = right("000" & request("idDescription"),3)
	idInventory = right("00000" & rsInventory("idInventory"),5)
	
	chrAssNum = idCustomer & "-" & idDescription & "-" & idInventory
	
	'update the Asset with the correct asset number
	sql = "execute UpdateInventorySpecial " &_
		rsInventory("idInventory") & ",'" &_
		chrAssNum & "'"
		
	'execute the SQL
	dbConnection.Execute(sql)
	
	rsInventory.Close
	set rsInventory = nothing
	
			
	'rrf - Detect which sumit button from edit Inventory, Add Return, or Just Add
	if request.form("submit1") <> "" then
		Response.Redirect "addinventory.asp?idWarehouse=" & Request("idWarehouse") & "&idCustomer=" & Request("idCustomer") & "&idDescription=" & Request("idDescription") & "&idOwner=" & request("idOwner")
	elseif request.form("submit2") <> "" then
		Response.Redirect "inventory.asp"
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Entry Error</title>
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
                  <td align="center"><font color="#0000FF" size="3" face="Arial, Helvetica, sans-serif"><strong>Add Inventory Error</strong></font></td>
                </tr>
                <tr> 
                  <td align="center"><p><font size="2" face="Arial, Helvetica, sans-serif">Duplicate Asset Error</font></p>
                  <p><font size="2" face="Arial, Helvetica, sans-serif">The database has detected that this asset already exists.</font></p>
                  <p><font size="2" face="Arial, Helvetica, sans-serif">idCustomer: <%=request("idCustomer")%> idDescription: <%=request("idDescription")%></font></p></td>
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
                  <td align="center"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Copyright &copy; 2003 
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
<%	
	'from the rsInventory
	'end if
	
	'Close database connections
	dbConnection.Close
	set dbConnection = nothing
%>