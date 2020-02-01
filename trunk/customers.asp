<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on Settings button
	buttonswitch = 5
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Find your profile
	set rsCustomers = server.CreateObject("adodb.recordset")
	sql = "execute ListCustomerNamesandIDs"
	set rsCustomers = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="inventory.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/settings-nav.htm" -->
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
                        <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Asset Management Customers</strong></font></td>
                        <td width="50%" align="right" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif"><a href="addcustomer.asp">Add Customer</a></font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr valign="top"> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr bgcolor="#6699cc"> 
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Customer</font></td>
                        <td height="20" align="right"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Options&nbsp;</font></td>
                      </tr>
<%
	if rsCustomers.EOF then
%>
                      <tr> 
                        <td height="20" colspan="2" align="center"><font size="1" face="Arial, Helvetica, sans-serif">There are no Customers to List</font></td>
                      </tr>
                      <tr bgcolor="#5b5b5b"> 
                        <td height="1" colspan="2"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsCustomers.EOF
		if bgswitch = 1 then
			bgcolor = "#ffffff"
			bgswitch = 0
		else
			bgcolor = "#f5f5f5"
			bgswitch = 1
		end if
%>
                      <tr> 
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsCustomers("chrCustomer"))%></font></td>
                        <td height="20" align="right" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif"><a href="editcustomer.asp?idCustomer=<%=rsCustomers("idCustomer")%>">Edit</a> - <a href="deletecustomer.asp?idCustomer=<%=rsCustomers("idCustomer")%>">Remove</a>&nbsp;</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="2" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
		rsCustomers.MoveNext
		loop
	end if
%>
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
  </form>
</body>
</html>
<%
	rsCustomers.Close
	set rsCustomers = nothing
	dbConnection.Close
	set dbConnection = nothing
%>