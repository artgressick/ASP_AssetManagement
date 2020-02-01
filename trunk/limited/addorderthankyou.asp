<%@ Language=VBScript %>
<%
	if session("idUser") = "" then
		Response.Redirect "../logon.asp"
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="../includes/title.htm" -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
    <!-- #include file="includes/top.htm" -->
    <tr> 
      <td width="10" background="images/leftverticalline.gif"><img src="images/leftverticalline.gif" width="10" height="10"></td>
      <td width="780"><table width="780" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr> 
                <td width="50%"><font size="4" face="Arial, Helvetica, sans-serif"><strong>Add Order - Successful</strong></font></td>
                <td width="50%" align="right">&nbsp;</td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="1" bgcolor="#6699cc"><img src="images/6699ccdot.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr> 
                <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">You 
                  have successfully added this order. Please do not go back to make 
                  changes as a duplicate order will be created. If you need to make 
                  any changes please contact techIT or your pool administrator.</font></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
              <tr> 
                <td align="center"><font color="#6699cc" size="3" face="Arial, Helvetica, sans-serif"><strong>Now 
                  that you have successfully added this order you can do the following:</strong></font></td>
              </tr>
              <tr> 
                <td height="30"><img src="images/ffffffdot.gif" width="1" height="1"></td>
              </tr>
              <tr> 
                <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><a href="loadorder.asp?idOrder=<%=request("idOrder")%>">Add assets to this Order</a></font></td>
              </tr>
              <tr> 
                <td align="center"> <p><font size="2" face="Arial, Helvetica, sans-serif">or</font></p></td>
              </tr>
              <tr> 
                <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><a href="default.asp">Return to the homepage and add assets later</a></font></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
        </tr>
      </table></td>
      <td width="10" background="images/rightverticalline.gif"><img src="images/rightverticalline.gif" width="10" height="10"></td>
    </tr>
    <!-- #include file="includes/bottom.htm" -->
  </table>
</body>
</html>
