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
	'Get the Users Saved Addresses
	set rsAddresses = server.CreateObject("adodb.recordset")
	sql = "execute ListSavedAddressesbyUser " & session("idUser")
	set rsAddresses = dbConnection.Execute(sql)
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
            <td width="100%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="595" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>My Saved Addresses</strong></font></td>
                        <td width="50%" align="right" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif"><a href="addaddress.asp">Add Address</a></font></td>
                      </tr>
                      <tr bgcolor="#f5f5f5"> 
                        <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">These 
                          are your saved addresses. &nbsp;Your addresses can be added 
                          from Carts and Orders, or from the link in the right hand corner 
                          of this page.<br> If you would like to edit or remove an address, click on the 
                          appropriate link in the Options column.<br>If you decide to remove an address, 
                          it will only be removed from this list and not from 
                          existing orders.<br></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr bgcolor="#6699cc"> 
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Saved Name</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Address</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;City</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;State</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Zip Code</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Country</font></td>
                        <td height="20"> <font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Options</font></td>
                      </tr>
<%
	if rsAddresses.EOF then
%>
                      <tr> 
                        <td height="20" colspan="7" align="center"><font size="1" face="Arial, Helvetica, sans-serif">You do not have any saved addresses at this time.</font></td>
                      </tr>
                      <tr bgcolor="#5b5b5b"> 
                        <td height="1" colspan="7"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsAddresses.EOF
		if bgswitch = 1 then
			bgcolor = "#ffffff"
			bgswitch = 0
		else
			bgcolor = "#f5f5f5"
			bgswitch = 1
		end if
%>
                      <tr bgcolor="<%=bgcolor%>"> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAddresses("chrSavedAddressName"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAddresses("chrAddress"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAddresses("chrCity"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAddresses("chrState"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAddresses("chrZip"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAddresses("chrCountry"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<A HREF="editaddress.asp?idAddress=<%=rsAddresses("idAddress")%>">Edit</A> - <A HREF="deletesavedaddresses.asp?idAddress=<%=rsAddresses("idAddress")%>">Remove</a></font></td>
                      </tr>
                      <tr bgcolor="#5b5b5b"> 
                        <td height="1" colspan="7"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
		rsAddresses.MoveNext
		loop
	end if
%>
                    </table></td>
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
	rsAddresses.Close
	set rsAddresses = nothing
	dbConnection.Close
	set dbConnection = nothing
%>