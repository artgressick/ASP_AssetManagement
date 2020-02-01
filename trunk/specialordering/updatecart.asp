<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "../logoff.asp"
	end if
%>
<!-- #include file="../includes/openconn.asp" -->
<%
	'Update the Cart status
	idStatus = 2
	
	sql = "execute UpdateCartApproval " & session("idCart") & "," & idStatus
	dbConnection.Execute(sql)
	
	'send the email to the people
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
    <!-- #include file="includes/top.htm" -->
    <tr> 
      <td bgcolor="#6699cc">
		<table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr bgcolor="#ffffff"> 
            <td valign="top" bgcolor="#ffffff">
			  <table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Request Received</strong><br></font></td>
                </tr>
                <tr> 
                  <td><font size="2" face="Arial, Helvetica, sans-serif">Thank you for your request.  
                    An email has been sent to the Pool Manager for approval and you have been copied. The pool manager 
                    will review what you have chosen and will Approve or Disapprove this cart.</font></td>
                </tr>
                <tr> 
                  <td height="25"><font size="2" face="Arial, Helvetica, sans-serif"><img src="images/ffffffdot.gif" width="1" height="1"></font></td>
                </tr>
                <tr> 
                  <td><font color="#008040" size="2" face="Arial, Helvetica, sans-serif"><strong><em>If approved by the Pool Manager - </em></strong></font>
                    <font color="#008000" size="2" face="Arial, Helvetica, sans-serif">you 
                      and the warehouse will receive an email to begin processing 
                      this request. You will receive an email when the request 
                      has been shipped as will the Onsite-Contact (if provided.)</font></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><font color="#D70000" size="2" face="Arial, Helvetica, sans-serif"><strong><em>If disapproved by the Pool Manager - </em></strong></font>
                    <font color="#D70000" size="2" face="Arial, Helvetica, sans-serif">You 
                      will receive an email along with the reason this request 
                      was disapproved. This request and all assets will be deleted 
                      and returned back to be loaned to another order.</font></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><font size="2" face="Arial, Helvetica, sans-serif">If you 
                    have any questions please contact you pool manager first, 
                    if your pool manager cannot answer your questions please contact 
                    techIT Solutions at (800) 492-2448.</font></td>
                </tr>
                <tr>
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr>
                  <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><strong><%if session("idAccess") = "R" then%><A HREF="../limited/default.asp"><%else%><A HREF="../orders.asp"><%end if%>Click here to return to the website</A> or close you browser if you are finished.</strong></font></td>
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
    <!-- #include file="includes/bottom.htm" -->
  </table>
</form>
</body>
</html>