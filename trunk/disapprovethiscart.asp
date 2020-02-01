<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on Settings button
	buttonswitch = 2
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="removecart.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/orders-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top"><table width="625" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15"><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="610"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Disapprove Cart</strong></font></td>
                  </tr>
                  <tr>
                    <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">You are about to disapprove this Cart request. <font color="#FF0000">This action is permanent!!</font> Once you disapprove this Cart, 
					it will be completely deleted from the system and all assets contained within will be made available for other orders.<br>
					Please type a brief reason for disapproving this request and
					click the Disapprove Cart button.</font></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Reason for disapproving this cart. This text will be sent to the person who requested this cart.<br>
                      <textarea name="reason" cols="50" rows="15" wrap="VIRTUAL" id="reason"></textarea></font></td>
                  </tr>
                  <tr>
                    <td height="25"><font size="1"><img src="images/ffffffdot.gif" width="1" height="1"></font></td>
                  </tr>
                  <tr>
                    <td><font size="1">
                      <input type="submit" name="Submit" value="Disapprove Cart"><input type="hidden" name="idCart" value="<%=request("idCart")%>">
                      </font></td>
                  </tr>
                </table></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
    </tr>
    <!-- #Begin bottom part -->
    <!-- #include file="includes/bottom.htm" -->
  </table>
  </form>
</body>
</html>