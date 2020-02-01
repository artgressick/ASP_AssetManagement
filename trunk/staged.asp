<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on inventory button
	buttonswitch = 2
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'find the Category information
	set rsCarts = server.CreateObject("adodb.recordset")
	sql = "execute ListStagedCarts"
	set rsCarts = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
		<!-- #include file="includes/orders-nav.htm" -->
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
                        <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Ship Staged Carts Process</strong></font></td>
                        <td width="50%" align="right" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr bgcolor="#f5f5f5">
						<td><font size="1" face="Arial, Helvetica, sans-serif">Below is a list of Staged carts that need to be shipped. Please click on the Cart Name to begin the process.</font></td>
					  </tr>
					</table>
                  </td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Cart Name</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Customer</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Ship Type</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Ship Date</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Arrival Date</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Approved</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Last Updated</font></td>
                      </tr>
<%
	if rsCarts.EOF then
%>
                      <tr> 
                        <td height="20" align="center" colspan="7"><font size="1" face="Arial, Helvetica, sans-serif">There are no Carts at this time that are staged.</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="7" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsCarts.EOF
		if bgswitch = 1 then
			bgcolor = "#ffffff"
			bgswitch = 0
		else
			bgcolor = "#f5f5f5"
			bgswitch = 1
		end if
%>
                      <tr> 
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<A HREF="editstaged.asp?idCart=<%=rsCarts("idCart")%>"><%=trim(rsCarts("chrCart"))%></A></font></td>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsCarts("chrCustomer"))%></font></td>
<%
		'AEG - Severity of Cart Shipping
		select case rsCarts("idExpedite")
			case 0
				chrExpedite = "Normal"
			case 1
				chrExpedite = "Expedite"
			case 2
				chrExpedite = "Rush"
			case else
				chrExpedite = "Normal"
		end select
%>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=chrExpedite%></font></td>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsCarts("dtShip"),2)%></font></td>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsCarts("dtArrival"),2)%></font></td>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsCarts("dtApproved"),2)%></font></td>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsCarts("dtUpdated"),2)%></font></td>
                      </tr>
                      <tr> 
                        <td colspan="7" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
		rsCarts.MoveNext
		loop
	end if
%>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
          </tr>
        </table></td>
    </tr>
    <!-- #Begin bottom part -->
    <!-- #include file="includes/bottom.htm" -->
  </table>
</body>
</html>
<%
	rsCarts.Close
	set rsCarts = nothing
	dbConnection.Close
	set dbConnection = nothing
%>