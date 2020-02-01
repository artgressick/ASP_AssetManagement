<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on inventory button
	buttonswitch = 4
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Get a list of Customers by user.
	set rsCarts = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListCartsReturning"
	else
		sql = "execute ListCartsReturningbyAccess " & session("idUser")
	end if
	set rsCarts = dbConnection.Execute(sql)
	
	'Get the list of Assets that have been added
	if request("idCart") = "" then
		errorflag = 1
	else
		set rsInventory = server.CreateObject("adodb.recordset")
		if session("idAccess") < "O" then
			sql = "execute ListAssetsReturning " & request("idCart")
		else
			sql = "execute ListAssetsReturningbyAccess " & request("idCart") & "," & session("idUser")
		end if
		set rsInventory = dbConnection.Execute(sql)
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
<script language="JavaScript">
<!--
function newWindow(updateWin) {
  updateWindow = window.open(updateWin,'updateWin','width=300,height=275');
updateWindow.focus()
}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="reportassetsreturning.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/reports-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="100%">
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="595" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Assets Returning Report</strong></font></td>
<%
	if session("idLoadInW") <> "" then
%>
                        <td align="right"><font size="1" face="Arial, Helvetica, sans-serif"><A HREF="checkin.asp">Return to Checkin</A></font></td>
<%
	end if
%>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr>
                  <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr>
                        <td bgcolor="#c0c0c0">
                          <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr bgcolor="#f5f5f5"> 
                              <td align="right" width="50%"><font size="2" face="Arial, Helvetica, sans-serif">Carts</font></td>
                              <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"> 
                                <select name="idCart" size="1" id="idCart">
                                  <option value="0" <%if cint(request("idCart")) = 0 then%>selected<%end if%>>All Carts</option>
<%
	if not rsCarts.EOF then
		do until rsCarts.EOF
%>
                                  <option value="<%=rsCarts("idCart")%>" <%if cint(request("idCart")) = rsCarts("idCart") then%>selected<%end if%>><%=trim(rsCarts("chrCart"))%></option>
<%
		rsCarts.MoveNext
		loop
	end if
%>
                                </select></font></td>
                              <td width="50%" align="left"><font size="1" face="Arial, Helvetica, sans-serif"><input name="Submit" type="submit" value="Find"></font></td>
                            </tr>
                          </table>
                         </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	if errorflag = 0 then
%>
                <tr> 
                  <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td height="20" bgcolor="#c0c0c0" colspan="7">
                          <table width="100%" border="0" cellspacing="0" cellpadding="5">
                            <tr>
                              <td align="left"><a HREF="javascript:newWindow('excel/excelassetreturning.asp?idCart=<%=request("idCart")%>')"><img SRC="images/exporttoexcel.gif" border="0" WIDTH="120" HEIGHT="19"></a></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                      <tr bgcolor="#6699cc"> 
                        <td height="20"><font color="#ffffff" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Asset #</font></td>
                        <td height="20"><font color="#ffffff" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Serial Number</font></td>
                        <td height="20"><font color="#ffffff" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Cart</font></td>
						<td height="20"><font color="#ffffff" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Customer</font></td>
                        <td height="20"><font color="#ffffff" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item</font></td>
                        <td height="20"><font color="#ffffff" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Arrival</font></td>
                        <td height="20"><font color="#ffffff" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Departure</font></td>
                      </tr>
<%
	if rsInventory.EOF then
%>
                      <tr> 
                        <td height="20" colspan="8" align="center"><font size="1" face="Arial, Helvetica, sans-serif">Not Assets to Display.</font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="8"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsInventory.EOF
		if bgswitch = 1 then
			bgswitch = 0
			bgcolor = "#ffffff"
		else
			bgswitch = 1
			bgcolor = "#f5f5f5"
		end if
		'create a counter for billing
		counter = counter+1
%>
                      <tr bgcolor="<%=bgcolor%>"> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrAssNum"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrSerialNum"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrCart"))%></font></td>
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrCustomer"))%></font></td>
<%
	if rsInventory("chrType") = "C" then
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrItem")) & " - " & trim(rsInventory("chrProcessor"))%><BR>
                        &nbsp;<%=trim(rsInventory("chrMemory")) & " - " & trim(rsInventory("chrODrive")) & " - " & trim(rsInventory("chrHDD"))%></font></td>
<%
	else
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrItem"))%></font></td>
<%
	end if
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsInventory("dtArrival"),2)%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsInventory("dtDeparture"),2)%></font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="8"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
                      </tr>
<%
		rsInventory.MoveNext
		loop
	end if
	'close the recordset
	rsInventory.Close
	set rsInventory = nothing
%>
                      <tr bgcolor="#6699cc"> 
                        <td height="20" colspan="8" align="left"><font color="#ffffff" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Total Assets Returning: <%=counter%></font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="8"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	'from the errorflag
	end if
%>
              </table>
            </td>
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
          </tr>
        </table></td>
    </tr>
    <!-- #Begin bottom part -->
    <!-- #include file="includes/bottom.htm" -->
  </table>
  </form>
</body>
</html>
<%
	rsCarts.Close
	set rsCarts = nothing
	dbConnection.Close
	set dbConnection = nothing
%>