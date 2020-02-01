<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on order button
	buttonswitch = 2
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'use stored procedure => MoveAsset
	'send the following in order: idInventory, idCart1, idCart2, idUser
	'print out to please check to make sure the assets have been moved.
	
	'prime the done counter
	done = 0
	
	'go through the checkboxes 1 by 1
	for i = 1 to cint(request("counter"))
		'check to see if the box was checked
		if len(request(i)) <> 0 then
			'count the Assets that were moved.
			done = done + 1
			'Move the Asset from one cart to the next
			sql = "execute MoveAsset " & _
					request(i) & "," &_
					request("idCart1") & "," &_
					request("idCart2") & "," &_
					session("idUser")
				'execute the upload
				dbConnection.Execute(sql)
		'from the if length/checked
		end if
	next

	'Get the First Cart Name
	set rsCart1 = server.CreateObject("adodb.recordset")
	sql = "execute FindCartNamebyID " & request("idCart1")
	set rsCart1 = dbConnection.Execute(sql)
	
	'Get the First Cart Name
	set rsCart2 = server.CreateObject("adodb.recordset")
	sql = "execute FindCartNamebyID " & request("idCart2")
	set rsCart2 = dbConnection.Execute(sql)
	
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
      <td width="100%" height="100%" valign="top">
	  	<table width="625" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15"><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="610"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Move Assets</strong></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td align="center"><p><font size="2" face="Arial, Helvetica, sans-serif"><strong>You have successfully
                        moved <font color="#0000FF"><%=done%></font> Assets from:</strong></font></p>
                        <p><font size="2" face="Arial, Helvetica, sans-serif"><strong><font color="#0000FF"><%=trim(rsCart1("chrCart"))%></font>
                        to <font color="#0000FF"><%=trim(rsCart2("chrCart"))%></font></strong></font></p>
                        <p><font color="#FF0000" size="2" face="Arial, Helvetica, sans-serif"><strong>Please do not go back or refresh this page.</strong></font></p>
                        <p><font size="2" face="Arial, Helvetica, sans-serif">Click here to return to the <a href="orders.asp">orders</a>
                        page or use the navigation button at top and side.</font></p></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
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
</body>
</html>
<%
	rsCart1.Close
	set rsCart1 = nothing
	rsCart2.Close
	set rsCart2 = nothing
	dbConnection.Close
	set dbConnection = nothing
%>