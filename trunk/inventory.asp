<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on inventory button
	buttonswitch = 3
	
	'First Line switch
	firstline = 0
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Get a list of the Categories
	set rsCategories = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListCategories"
	else
		sql = "execute ListCategoriesbyUser " & session("idUser")
	end if
	set rsCategories = dbConnection.Execute(sql)
	
	'prime the descriptions
	if request("idCategory") = "" then
		idCategory = rsCategories("idCategory")
	else
		idCategory = request("idCategory")
	end if
	
	'Get the descriptions for the category need to make counts work also
	set rsDescriptions = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute FindDescriptionbyCategory " & idCategory
	else
		sql = "execute FindDescriptionbyCategoryandUser " & idCategory & "," & session("idUser")
	end if
	set rsDescriptions = dbConnection.Execute(sql)
	
	'find the Category information
	set rsCategory = server.CreateObject("adodb.recordset")
	sql = "execute FindCategorybyID " & idCategory
	set rsCategory = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="inventory.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
		<!-- #include file="includes/inventory-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="100%">
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="595" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr>
                        <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Warehouse Inventory</strong></font></td>
                        <td width="50%" align="right" valign="bottom"><font size="1" face="Arial, Helvetica, sans-serif"><a href="inventorysearch.asp">Power Search</a>&nbsp;|&nbsp;<a href="accountteam.asp">Need Help</a>?</font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr>
                        <td bgcolor="#5b5b5b"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr bgcolor="#f5f5f5"> 
                              <td align="right" width="50%"><font size="2" face="Arial, Helvetica, sans-serif">Categories</font></td>
                              <td><font size="1" face="Arial, Helvetica, sans-serif"><select name="idCategory" size="1" id="idCategory">
<%
	if not rsCategories.EOF then
		do until rsCategories.EOF
%>
                                  <option value="<%=rsCategories("idCategory")%>" <%if rsCategories("idCategory") = cint(idCategory) then%>selected<%end if%>><%=trim(rsCategories("chrCategory"))%></option>
<%
		rsCategories.MoveNext
		loop
	end if
%>
                                </select>
                                </font></td>
                              <td align="left" width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><input type="submit" name="Submit" value="Get Descriptions"></font></td>
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
                <tr> 
                  <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="4"><font size="3" face="Arial, Helvetica, sans-serif"><b><%=trim(rsCategory("chrCategory"))%></b>&nbsp;<%if session("idAccess") < "O" then%><font size="1">&nbsp;<A HREF="editcategory.asp?idCategory=<%=rsCategory("idCategory")%>">Edit this category</A>&nbsp;</font><%end if%></font></td>
                      </tr>
                      <tr> 
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item #</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item Description</font></td>
                        <td height="20" bgcolor="#6699CC"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Quantity</font></td>
                        <td height="20" bgcolor="#6699CC" align="right"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Options&nbsp;</font></td>
                      </tr>
<%
	if rsDescriptions.EOF then
%>
                      <tr> 
                        <td height="20" align="center" colspan="4"><font size="1" face="Arial, Helvetica, sans-serif">There are no descriptions in this category.</font></td>
                      </tr>
                      <tr> 
                        <td height="1" colspan="4" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsDescriptions.EOF
		if bgswitch = 1 then
			bgcolor = "#ffffff"
			bgswitch = 0
		else
			bgcolor = "#f5f5f5"
			bgswitch = 1
		end if
%>
                      <tr> 
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;#<%=rsDescriptions("chrItemNo")%></font></td>
<%
		if rsDescriptions("chrType") = "C" then
%>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<A HREF="viewdescription.asp?idDescription=<%=rsDescriptions("idDescription")%>"><%=trim(rsDescriptions("chrItem")) & " - " & trim(rsDescriptions("chrProcessor"))%></a><br>
                        &nbsp;<%=trim(rsDescriptions("chrMemory")) & " - " & trim(rsDescriptions("chrODrive"))%></font></td>
<%
		else
%>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<A HREF="viewdescription.asp?idDescription=<%=rsDescriptions("idDescription")%>"><%=trim(rsDescriptions("chrItem"))%></A></font></td>
<%
		end if
%>
                        <td height="20" bgcolor="<%=bgcolor%>"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=rsDescriptions("intAssets")%></font></td>
                        <td height="20" bgcolor="<%=bgcolor%>" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%if session("idAccess") < "O" then%><A HREF="editdescription.asp?idDescription=<%=rsDescriptions("idDescription")%>">Edit</A><%end if%>&nbsp;</font></td>
                      </tr>
                      <tr> 
                        <td colspan="4" bgcolor="#5b5b5b"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
		rsDescriptions.MoveNext
		loop
	end if
%>
                    </table>
                  </td>
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
  </form>
</body>
</html>
<%
	rsCategories.Close
	set rsCategories = nothing
	rsDescriptions.Close
	set rsDescriptions = nothing
	rsCategory.Close
	set rsCategory = nothing
	dbConnection.Close
	set dbConnection = nothing
%>