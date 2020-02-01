<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "../logoff.asp"
	end if
%>
<!-- #include file="../includes/openconn.asp" -->
<%
	'get a list of orders that will return in 7 days
	set rsTopics = server.CreateObject("adodb.recordset")
	sql = "execute FindTopicbyID " & request("idBlog")
	set rsTopics = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>techIT Solutions Asset Management Blog</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<form name="form1" method="post" action="">
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr> 
            <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Asset Management BLOG</strong></font></td>
            <td width="50%" align="right" valign="bottom">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="20">&nbsp;</td>
    </tr>
    <tr> 
      <td><table width="100%" border="0" cellspacing="0" cellpadding="4">
<%
	'determine the Status
	select case rsTopics("idStatus")
		case 1
			Bstatus = "New"
		case 2
			Bstatus = "Complete"
	end select
			
	'determine the Type
	select case rsTopics("idType")
		case 1
			Btype = "Problem"
		case 2
			Btype = "Upgrade"
	end select
			
	'determine the Type
	select case rsTopics("idPriority")
		case 1
			priority = "High"
		case 2
			priority = "Medium"
		case 3
			priority = "Low"
	end select
%>
          <tr bgcolor="#f5f5f5"> 
            <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Type: <%=Btype%></font></td>
            <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Date Entered: <%=formatdatetime(rsTopics("dtStamp"),1)%></font></td>
          </tr>
          <tr bgcolor="#f5f5f5"> 
            <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Priority: <%=priority%></font></td>
            <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Entered by: <%=trim(rsTopics("chrFirst")) & " " & trim(rsTopics("chrLast"))%></font></td>
          </tr>
          <tr bgcolor="#f5f5f5"> 
            <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">Status: <%=Bstatus%></font></td>
          </tr>
          <tr bgcolor="#f5f5f5"> 
            <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">Title: <%=trim(rsTopics("chrTitle"))%></font></td>
          </tr>
          <tr bgcolor="#f5f5f5"> 
            <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">Message: <%=trim(rsTopics("txtMessage"))%></font></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="20">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
<%
	rsTopics.close
	set rsTopics = nothing
	dbConnection.close
	set dbConnection = nothing
%>