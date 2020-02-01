<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	if session("idLoadInW") = "" then
		Response.Redirect "loadcheckin.asp"
	end if
	
	'turn on order button
	buttonswitch = 2
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Find the Asset by the Asset Number
	set rsAsset = server.CreateObject("adodb.recordset")
	sql = "execute FindInventorybyAssNum '" & request("chrAssNum") & "'"
	set rsAsset = dbConnection.Execute(sql)
	
	'if the asset cannot be found then it was a bad asset number
	if not rsAsset.eof then
	'Find out the Asset status and do the updating if possible
	select case rsAsset("idStatus")
		case 1
			'This is ok and we can do all of the stuff we need to do here for check out
			flag = 1
			message = "This Asset cannot be checked in. It is listed as already being in the warehouse/ready."
		case 2
			'This is checked out
			flag = 0
			message = "Asset has been checked out"
			'it has the correct id so we can check it back in
			'checkin the asset
			sql = "execute CheckInAsset " & rsAsset("idCart") & "," & rsAsset("idInventory") & "," & session("idLoadInW") & "," & session("idUser")
			dbConnection.execute(sql)
		case 3
			'This is checked back in but not been turned
			flag = 1
			message = "This Asset is listed as ready for turning."
		case 4
			'This asset has been noted as out for Repair
			flag = 1
			message = "This Asset has been marked out for repair. Please make sure everything is working and then change the status to ready."
		case 5
			'This asset is listed at Permanent Loan
			flag = 1
			message = "This Asset has been put on Permanent Loan please make sure that it has been properly returned to the system and is not going out permanently."
		case 6
			'This asset is Out of System
			flag = 1
			message = "This Asset is Out of the System. Please check with the Account Manager on this Asset."
		case 7
			'This asset is Internal Use
			flag = 0
			message = "This Asset has been put on Internal Use. Please check with the Account Manager to make sure this is correct."
			'it has the correct id so we can check it back in
			'checkin the asset
			sql = "execute CheckInAsset " & rsAsset("idCart") & "," & rsAsset("idInventory") & "," & session("idLoadInW") & "," & session("idUser")
			dbConnection.execute(sql)
		case 8
			'This asset is Broken/Damaged
			flag = 1
			message = "This Asset has  been marked Broken/Damaged. Please make sure it has been repaired before returning this Asset to the system."
		case 9
			'This asset is Lost/Stolen
			flag = 1
			message = "This Asset was listed as Lost/Stolen. Please notify techIT that it has been found."
	end select
	'from the rsAsset.eof
	else
		flag = 1
		message = "The Asset number did not match a number in the database please try again."
	end if
	
	'if everything was ok then return to the checkout page.
	if flag = 0 then
		Response.Redirect "checkin.asp?chrAssNum=" & request("chrAssNum") & "&idCart=" & rsAsset("idCart")
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="checkoutasset.asp">
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
            <td width="610">
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr>
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Check In Error</strong></font></td>
                        <td align="right" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif"><a href="pulllist.asp">List Here</a></font></td>
                      </tr>
                      <tr bgcolor="#f5f5f5">
                        <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">Please read the error indicated below and follow the recommended corrective actions.<br>
						    The Asset you entered has not been checked in at this time.</font></td>
                      </tr>
                    </table>
				  </td>
                </tr>
                <tr>
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr>
                  <td bgcolor="#5b5b5b">
				    <table width="100%" border="0" cellspacing="1" cellpadding="0">
                      <tr bgcolor="#f5f5f5">
                        <td width="50%" bgcolor="#ffffff">
						  <table width="100%" border="0" cellspacing="0" cellpadding="5">
                            <tr>
                              <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Error:</strong> <font color="#FF0000"><%=message%></font></font></td>
                            </tr>
							<tr>
                              <td align="center"><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                            </tr>
							<tr>
                              <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><a href="checkin.asp">Click here to return to check in.</a></font></td>
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
	rsAsset.Close
	set rsAsset = nothing
	dbConnection.Close
	set dbConnection = nothing
%>