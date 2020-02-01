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
<form name="form1" method="post" action="updateemailnotifications.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/settings-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top">
		<table width="625" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15"><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="610">
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15" colspan="2"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Email 
                          Notification Settings</strong></font><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                      </tr>
                      <tr bgcolor="#f5f5f5"> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">These 
                          ae your Email notifications for you company. Please 
                          make sure the examples are followed.</font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15" colspan="2"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr valign="top"> 
                  <td width="50%">
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td bgcolor="#6699cc"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong>When 
                          a New Cart is Created</strong></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox" value="checkbox">
                          Pool Manager</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox2" value="checkbox">
                          Creator </font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox3" value="checkbox">
                          Onsite Contact</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Additional 
                          Contacts (Leave blank if none)<br>
                          <input type="text" name="textfield">
                          </font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#6699cc"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"><strong>Prior 
                          to Cart Expiration Date</strong></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox7" value="checkbox">
                          Pool Manager</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox8" value="checkbox">
                          Creator </font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox9" value="checkbox">
                          Onsite Contact</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Additional 
                          Contacts (Leave blank if none)<br>
                          <input type="text" name="textfield3">
                          </font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#6699cc"><font color="#ffffff" size="2" face="Arial, Helvetica, sans-serif"><strong>When 
                          Cart is Approved or Disapproved</strong></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox15" value="checkbox">
                          Pool Manager</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox14" value="checkbox">
                          Creator</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox13" value="checkbox">
                          Onsite Contact</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Additional 
                          Contacts (Leave blank if none)<br>
                          <input type="text" name="textfield5">
                          </font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#6699cc"><font color="#ffffff" size="2" face="Arial, Helvetica, sans-serif"><strong>One 
                          Day Prior to Return Date</strong></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox19" value="checkbox">
                          Pool Manager</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox20" value="checkbox">
                          Creator</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox21" value="checkbox">
                          Onsite Contact</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Additional 
                          Contacts (Leave blank if none)<br>
                          <input type="text" name="textfield7">
                          </font></td>
                      </tr>
                    </table></td>
                  <td width="50%"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td bgcolor="#6699cc"><font color="#ffffff" size="2" face="Arial, Helvetica, sans-serif"><strong>When 
                          Cart is Ready for Approval</strong></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox4" value="checkbox">
                          Pool Manager</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox5" value="checkbox">
                          Creator</font></td>
                      </tr>
                      <tr> 
                        <td> <font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox6" value="checkbox">
                          Onsite Contact</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Additional 
                          Contacts (Leave blank if none.)<br>
                          <input type="text" name="textfield2">
                          </font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#6699cc"><font color="#ffffff" size="2" face="Arial, Helvetica, sans-serif"><strong>When 
                          Cart Expires</strong></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox10" value="checkbox">
                          Pool Manager</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox11" value="checkbox">
                          Creator </font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox12" value="checkbox">
                          Onsite Contact </font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Additional 
                          Contacts (Leave blank if none)<br>
                          <input type="text" name="textfield4">
                          </font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#6699cc"><font color="#ffffff" size="2" face="Arial, Helvetica, sans-serif"><strong>When 
                          Cart is Shipped</strong></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox16" value="checkbox">
                          Pool Manager</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox17" value="checkbox">
                          Creator</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox18" value="checkbox">
                          Onsite Contact</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Additional 
                          Contacts (Leave blank if none)<br>
                          <input type="text" name="textfield6">
                          </font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#6699cc"><font color="#ffffff" size="2" face="Arial, Helvetica, sans-serif"><strong>When 
                          Cart Returns To Warehouse</strong></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox22" value="checkbox">
                          Pool Manager</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox23" value="checkbox">
                          Creatore</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <input type="checkbox" name="checkbox24" value="checkbox">
                          Onsite Contact</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Additional 
                          Contacts (Leave blank if none)<br>
                          <input type="text" name="textfield8">
                          </font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15" colspan="2"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr>
                  <td colspan="2"><input type="submit" name="Submit" value="Save Settings"></td>
                </tr>
                <tr>
                  <td height="25" colspan="2"><img src="images/ffffffdot.gif" width="1" height="1"></td>
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