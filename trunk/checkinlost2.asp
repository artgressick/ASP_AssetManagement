<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on order button
	buttonswitch = 3
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Get a list of Cart that are open
	set rsInventory = server.CreateObject("adodb.recordset")
	sql = "execute FindInventorybyAssNum '" & request("chrAssNum") & "'"
	set rsInventory = dbConnection.Execute(sql)
	
	if rsInventory.EOF then
		errorflag = 1 'AEG - not an asset
		errormessage = "This asset does not exist. Please go back and try again."
	else
		if rsInventory("idStatus") <> 9 then
			errorflag = 1 'AEG - this asset needs to be checked back in first
			errormessage = "This asset's status is not Lost. Please make sure that you have not already check this asset back in."
		else
			errorflag = 0 'AEG - ok
			set rsProblems = server.CreateObject("adodb.recordset")
			sql = "execute FindLostTicketbyInventory " & rsInventory("idInventory")
			set rsProblems = dbConnection.Execute(sql)
		end if
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="checkinlostfinal.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/inventory-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top"><table width="625" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15"><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="610"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Check In Lost</strong></font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">Listed below are the open tickets of Losses. Please check the system before contining.</font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	if errorflag = 1 then
%>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><strong><%=errormessage%></strong></font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
<%
	else
%>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Warehouse<br>
                        <font size="2"><b><%=trim(rsProblems("chrWarehouse"))%></b></font></font></td>
                      </tr>
<%
		if rsProblems("idCart") = 0 then
			chrCart = "From Warehouse"
		else
			chrCart = trim(rsProblems("chrCart"))
		end if
%>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Cart<br>
                        <font size="2"><b><%=chrCart%></b></font></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Problem with Asset<br>
                        <font size="2"><b><%=trim(rsProblems("txtExplanation"))%></b></font></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr>
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr align="center"> 
                        <td><input name="submit1" type="submit" id="submit1" value="Continue with Check In">
                        <input name="idInventory" type="hidden" value="<%=rsInventory("idInventory")%>">
                        <input name="idLost" type="hidden" value="<%=rsProblems("idLost")%>"></td>
                        <td><input type="submit" name="submit2" id="submit2" value="Cancel Check In"></td>
                      </tr>
                    </table>
                  </td>
                </tr>
<%
		'AEG - Close the database connections
		rsProblems.Close
		set rsProblems = nothing
	end if
%>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <!-- #Begin bottom part -->
    <!-- #include file="includes/bottom.htm" -->
  </table>
  </form>
</body>
</html>
<%
	rsInventory.Close
	set rsInventory = nothing
	dbConnection.Close
	set dbConnection = nothing
%>