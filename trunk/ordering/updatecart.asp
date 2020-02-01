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
	idStatus = 2 'Ready for Approval
	
	'AEG - Update the Cart Status	
	sql = "execute UpdateCartApproval " & session("idCart") & "," & idStatus
	dbConnection.Execute(sql)
	
	'AEG - Find the Cart Name
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute FindCartNamebyID " & session("idCart")
	set rsCart = dbConnection.Execute(sql)
	
	'AEG - Move the Cart Name to a Temp Variable
	chrCart = trim(rsCart("chrCart"))
	
	'Close the recordset
	rsCart.Close
	set rsCart = nothing
	
	'--------------------------------------------------------------------------
	'AEG - Open the SMTP Mailer Client
	Set Mailer = Server.CreateObject("SoftArtisans.SMTPMail") 'from www.softartisan.com
	Mailer.RemoteHost  = "techit-ex2.techitsolutions.com" 'mail server
	Mailer.FromName    = "administrator"
	Mailer.FromAddress = "administrator@techitsolutions.com"
	'--------------------------------------------------------------------------
	'AEG - Find the User to send an email.
	set rsUser = server.CreateObject("adodb.recordset")
	sql = "execute FindUserbyID " & session("idUser")
	set rsUser = dbConnection.Execute(sql)
	
	'AEG - Move the records to temp fields
	chrUserName = trim(rsUser("chrFirst")) & " " & trim(rsUser("chrLast"))
	chrUserEmail = trim(rsUser("chrEmail"))
	
	'AEG - Attach the User to the email
	Mailer.AddRecipient chrUserName, chrUserEmail
	
	'Close the User Recordset
	rsUser.Close
	set rsUser = nothing
	'--------------------------------------------------------------------------
	'AEG - Find the Pool Manager(s)
	set rsPoolManager = server.CreateObject("adodb.recordset")
	sql = "execute FindPoolManagersbyCart " & session("idCart")
	set rsPoolManager = dbConnection.Execute(sql)
	
	'AEG - Only do this if we have a record.
	if not rsPoolManager.EOF then
		do until rsPoolManager.EOF
			'Move the information to the Temp Variables
			chrPoolManagerName = trim(rsPoolManager("chrFirst")) & " " & trim(rsPoolManager("chrLast"))
			chrPoolManagerEmail = trim(rsPoolManager("chrEmail"))
			
			'Attach the information to the email
			Mailer.AddRecipient chrPoolManagerName, chrPoolManagerEmail
			
		rsPoolManager.MoveNext
		loop
	end if
	
	'Close the Pool Manager Recordset
	rsPoolManager.Close
	set rsPoolManager = nothing
	'--------------------------------------------------------------------------
	'AEG - Start the Message information
	Mailer.Subject     = "Cart Ready for Approval - " & chrCart
	Mailer.BodyText    = chrPoolManagerName & VbCrLf & VbCrLf &_
	chrUserName & " has submitted a cart named, " & chrCart & ", that requires your review." & VbCrLf & VbCrLf &_
	"Thank you..." & VbCrLf &_
	"techIT Solutions Asset Management Team"
	
	'Execute the email
	Mailer.SendMail
	
	'Close the Database Connections
	dbConnection.Close
	set dbConnection = nothing
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
                  <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Request Received</strong></font></td>
                </tr>
                <tr> 
                  <td><font size="2" face="Arial, Helvetica, sans-serif">Thank you for your request.  
                    An email has been sent to the Pool Manager for approval and you have been copied. The pool manager 
                    will review what you have requested, and will Approve or Disapprove this cart.</font></td>
                </tr>
                <tr> 
                  <td height="25"><font size="2" face="Arial, Helvetica, sans-serif"><img src="images/ffffffdot.gif" width="1" height="1"></font></td>
                </tr>
                <tr> 
                  <td><font color="#008040" size="2" face="Arial, Helvetica, sans-serif"><strong><em>If approved by the Pool Manager - </em></strong></font>
                    <font color="#008000" size="2" face="Arial, Helvetica, sans-serif">you 
                      will receive an email notifying you of the approval,
                      and you will also receive an email when the items have been shipped.</font></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><font color="#D70000" size="2" face="Arial, Helvetica, sans-serif"><strong><em>If disapproved by the Pool Manager - </em></strong></font>
                    <font color="#D70000" size="2" face="Arial, Helvetica, sans-serif">You 
                      will receive an email along with the reason this request 
                      was disapproved.</font></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><font size="2" face="Arial, Helvetica, sans-serif">If you 
                    have any questions please contact your Pool Manager or your  
                    techIT Solutions Account Manager.</font></td>
                </tr>
                <tr>
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr>
                  <td align="center"><font size="2" face="Arial, Helvetica, sans-serif"><strong><%if session("idAccess") = "R" then%><A HREF="../limited/default.asp"><%else%><A HREF="../orders.asp"><%end if%>Click here to return to the website</A> or close your browser to end your session.</strong></font></td>
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