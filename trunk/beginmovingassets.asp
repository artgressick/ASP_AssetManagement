<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on order button
	buttonswitch = 2
	
	if request("idCart1") = request("idCart2") then
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
The carts are the same please go back and try choosing differnet carts. 
</body>
</html>
<%
	'this is from the 
	else
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Get a list of Assets that they have in common.
	set rsInventory = server.CreateObject("adodb.recordset")
	sql = "execute ListAssetsinCommon2Carts " & request("idCart1") & "," & request("idCart2")
	set rsInventory = dbConnection.Execute(sql)
	
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
<form name="form1" method="post" action="updatemoveassets.asp">
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
                        <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Move Assets</strong></font></td>
                        <td width="50%">&nbsp;</td>
                      </tr>
                      <tr bgcolor="#f5f5f5"> 
                        <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">Listed below are all of the Assets that the two carts have 
                          in common. Please check the assets that you want to move and then click on the Move Assets button at the bottom </font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr> 
                        <td bgcolor="#5b5b5b">
						  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr bgcolor="#f5f5f5"> 
                              <td align="center" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Moving Assets from <font color="#0000ff"><%=trim(rsCart1("chrCart"))%></font> to <font color="#0000ff"><%=trim(rsCart2("chrCart"))%></font></strong></font></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr bgcolor="#6699cc"> 
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Asset Number</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Serial Number</font></td>
                        <td height="20" align="left"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item</font></td>
                      </tr>
<%
	if rsInventory.EOF then
		flag = 0
%>
                      <tr align="center"> 
                        <td height="20" colspan="4"><font size="1" face="Arial, Helvetica, sans-serif">There are no Assets in common or all assets have been check in to the second Cart.</font></td>
                      </tr>
                      <tr bgcolor="#5b5b5b"> 
                        <td height="1" colspan="4"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		flag = 1
		do until rsInventory.EOF
		counter = counter + 1
%>
                      <tr> 
                        <td height="20" align="center"><input type="checkbox" name="<%=counter%>" value="<%=rsInventory("idInventory")%>"></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrAssNum"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrSerialNum"))%></font></td>
<%
		if rsInventory("chrType") = "C" then
%>
                        <td height="20" align="left"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrItem")) & " - " & trim(rsInventory("chrProcessor"))%><br>
                        &nbsp;<%=trim(rsInventory("chrMemory")) & " - " & trim(rsInventory("chrHDD")) & " - " & trim(rsInventory("chrODrive"))%></font></td>
<%
		else
%>
                          <td height="20" align="left"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrItem"))%></font></td>
<%
		end if
%>
                      </tr>
                      <tr bgcolor="#5b5b5b"> 
                        <td height="1" colspan="4"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
		rsInventory.MoveNext
		loop
	end if
%>
                    </table>
                  </td>
                </tr>
<%
	if flag = 1 then
%>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><input type="submit" name="Submit" value="Move Assets">
                        <input type="hidden" name="counter" value="<%=counter%>">
                        <input type="hidden" name="idCart1" value="<%=request("idCart1")%>">
                        <input type="hidden" name="idCart2" value="<%=request("idCart2")%>"></td>
                      </tr>
                    </table></td>
                </tr>
<%
	end if
%>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
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
	rsCart1.Close
	set rsCart1 = nothing
	rsCart2.Close
	set rsCart2 = nothing
	dbConnection.Close
	set dbConnection = nothing
	
	'this is from the error
	end if
%>