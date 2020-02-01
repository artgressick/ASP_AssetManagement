<%@ Language=VBScript %>
<!-- #include file="includes/openconn.asp" -->
<%
	'look up the email address --------------------------------------------------------------------
	set rsAccount = server.CreateObject("adodb.recordset")
	sql = "execute FindUserbyEmail '" & request("chrEmail") & "'"
	set rsAccount = dbConnection.Execute(sql)
	
	'check to see if any records returned ---------------------------------------------------------
	if rsAccount.EOF then
		errorflag = 1 'bad
	else
		errorflag = 0 'good
		chrUserName = trim(rsAccount("chrFirst")) & " " & trim(rsAccount("chrLast"))
		chrUserEmail = trim(rsAccount("chrEmail"))
		'email the user the information -----------------------------------------------------------
		Set Mailer = Server.CreateObject("SoftArtisans.SMTPMail") 'from www.softartisan.com
		
		Mailer.RemoteHost  = "66.35.209.152" 'mail server
		Mailer.FromName    = "administrator"
		Mailer.FromAddress = "administrator@techitsolutions.com"
		Mailer.AddRecipient chrUserName, chrUserEmail
		Mailer.Subject     = "Asset Management - Password Information"
		Mailer.BodyText    = "You or someone has requested the Username and password be sent to this email registered email address. If you did not request this information please log into the Asset Management System and change you password." & VbCrLf & VbCrLf &_
		"Username: " & trim(rsAccount("chrEmail")) & VbCrLf &_
		"Password: " & trim(rsAccount("chrPassword")) & VbCrLf &_
		"Website: http://www.itechit.com" & VbCrLf & VbCrLf &_
		"If you have any questions please contact Customer Service at 1-800-492-2448 or email us at customerservice@techitsolutions.com"
		'Execute the email
		Mailer.SendMail
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title.htm" -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#ffffff" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form action="loadaccount.asp" method="post" name="logon" id="logon">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td align="center" valign="middle"><table width="620" height="350" border="0" cellpadding="0" cellspacing="0">
          <!-- #include file="includes/logon-top.htm" -->
          <tr>
            <td width="10" height="330" background="images/leftblack.gif"><img src="images/leftblack.gif" width="10" height="10"></td>
            <td width="600" valign="top" bgcolor="#f5f5f5"><table width="600" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="25"><img src="images/f5f5f5dot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="600" border="0" cellspacing="0" cellpadding="2">
                      <tr> 
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Lookup password</strong></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="1" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/f5f5f5dot.gif" width="1" height="1"></td>
                </tr>
<%
	if errorflag = 1 then
%>
                <tr> 
                  <td align="center"><p><font size="4" face="Arial, Helvetica, sans-serif"><strong>Error: We could not locate your account</strong></font></p>
                    <p><font size="2" face="Arial, Helvetica, sans-serif">Either you typed in the wrong email address or you do not have an account</font></p>
                    <p><font size="2" face="Arial, Helvetica, sans-serif"><a href="lostinformation.asp">Please click here to try again</a>.</font></p></td>
                </tr>
<%
	else
%>
                <tr> 
                  <td align="center"><p><font size="4" face="Arial, Helvetica, sans-serif"><strong>Password Lookup Successful</strong></font></p>
                    <p><font size="2" face="Arial, Helvetica, sans-serif">Your password should be arriving shortly in the email.</font></p>
                    <p><font size="2" face="Arial, Helvetica, sans-serif"><a href="default.asp">Click here to try logging in again</a>.</font></p></td>
                </tr>
<%
	end if
%>
              </table></td>
            <td width="10" height="330" background="images/rightblack.gif"><img src="images/rightblack.gif" width="10" height="10"></td>
          </tr>
          <!-- #include file="includes/logon-bottom.htm" -->
        </table>
      </td>
    </tr>
  </table>
</form>
</body>
</html>
<%
	rsAccount.Close
	set rsAccount = nothing
	dbConnection.Close
	set dbConnection = nothing
%>