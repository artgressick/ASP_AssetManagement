<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'AEG - Find the Warehouse to send an email.
	set rsAssets = server.CreateObject("adodb.recordset")
	sql = "execute ChangingCartDates " & request("idCart") & ",'" & request("dtPull") & "','" & request("dtTurn") & "'"
	set rsAssets = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Update Check</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="updatecart.asp">
  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="10" height="10" align="left" valign="top"><img src="images/topleftblue.gif" width="10" height="10"></td>
            <td width="780" height="10"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td><img src="images/eamlogo.gif" width="150" height="45"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                </tr>
              </table></td>
            <td width="10" height="10" align="right" valign="top"><img src="images/toprightblue.gif" width="10" height="10"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr>
            <td bgcolor="#FFFFFF">
			  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td height="35"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td align="center"><font color="#0000FF" size="3" face="Arial, Helvetica, sans-serif"><strong>Assets Affected</strong></font></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
					  <tr bgcolor="#6699cc">
						<td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Asset Number</font></td>
						<td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Serial Number</font></td>
						<td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item</font></td>
					  </tr>
<%
	if rsAssets.EOF then
%>
					  <tr>
						<td colspan="3" height="20" align="center"><font size="1" face="Arial, Helvetica, sans-serif">No Assets are affected by this change</font></td>
					  </tr>
					  <tr>
						<td colspan="3" height="1" align="center" bgcolor="#c0c0c0"><img src="images/c0c0c0dot.gif" WIDTH="1" HEIGHT="1"></td>
					  </tr>
<%
	else
		do until rsAssets.EOF
		if bgswitch = 1 then
			bgswitch = 0
			bgcolor = "#ffffff"
		else
			bgswitch = 1
			bgcolor = "#f5f5f5"
		end if
%>
					  <tr bgcolor="<%=bgcolor%>">
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrAssNum"))%></font></td>
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrSerialNum"))%></font></td>
<%
		if rsAssets("chrType") = "C" then
%>
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrItem")) & " " & trim(rsAssets("chrProcessor"))%><br>
						&nbsp;<%=trim(rsAssets("chrMemory")) & " - " & trim(rsAssets("chrODrive")) & " - " & trim(rsAssets("chrHDD"))%></font></td>
<%
		else
%>
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrItem"))%></font></td>
<%
		end if
%>
					  </tr>
					  <tr>
						<td colspan="3" height="1" align="center" bgcolor="#c0c0c0"><img src="images/c0c0c0dot.gif" WIDTH="1" HEIGHT="1"></td>
					  </tr>
<%
		rsAssets.MoveNext
		loop
	end if
%>
					</table>
                  </td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td align="center"><input name="Cancel" type="submit" value="Cancel Update">&nbsp;&nbsp;
                  <input name="Update" type="submit" value="Update Cart">
                  <input name="idSpecial" type="hidden" value="0">
                  <input name="intCounter" type="hidden" value="0">
                  <input name="idCart" type="hidden" value="<%=request("idCart")%>">
                  <input name="idBilled" type="hidden" value="<%=request("idBilled")%>">
                  <input name="idLoaner" type="hidden" value="<%=request("idLoaner")%>">
                  <input name="idStatus" type="hidden" value="<%=request("idStatus")%>">
                  <input name="chrOrder" type="hidden" value="<%=replace(request("chrOrder"),"'","''")%>">
                  <input name="chrDDName" type="hidden" value="<%=replace(request("chrDDName"),"'","''")%>">
                  <input name="chrDDNumber" type="hidden" value="<%=replace(request("chrDDNumber"),"'","''")%>">
                  <input name="chrManager" type="hidden" value="<%=replace(request("chrManager"),"'","''")%>">
                  <input name="chrIPerson" type="hidden" value="<%=replace(request("chrIPerson"),"'","''")%>">
                  <input name="chrIEmail" type="hidden" value="<%=replace(request("chrIEmail"),"'","''")%>">
                  <input name="dtPull" type="hidden" value="<%=request("dtPull")%>">
                  <input name="dtShip" type="hidden" value="<%=request("dtShip")%>">
                  <input name="dtArrival" type="hidden" value="<%=request("dtArrival")%>">
                  <input name="dtArrivalTime" type="hidden" value="<%=request("dtArrivalTime")%>">
                  <input name="dtDeparture" type="hidden" value="<%=request("dtDeparture")%>">
                  <input name="dtDepartureTime" type="hidden" value="<%=request("dtDepartureTime")%>">
                  <input name="dtReturn" type="hidden" value="<%=request("dtReturn")%>">
                  <input name="dtTurn" type="hidden" value="<%=request("dtTurn")%>">
                  <input name="chrAddress" type="hidden" value="<%=replace(request("chrAddress"),"'","''")%>">
                  <input name="chrAddress2" type="hidden" value="<%=replace(request("chrAddress2"),"'","''")%>">
                  <input name="chrAddress3" type="hidden" value="<%=replace(request("chrAddress3"),"'","''")%>">
                  <input name="chrAddress4" type="hidden" value="<%=replace(request("chrAddress4"),"'","''")%>">
                  <input name="chrCity" type="hidden" value="<%=replace(request("chrCity"),"'","''")%>">
                  <input name="chrState" type="hidden" value="<%=replace(request("chrState"),"'","''")%>">
                  <input name="chrZip" type="hidden" value="<%=replace(request("chrZip"),"'","''")%>">
                  <input name="chrCountry" type="hidden" value="<%=replace(request("chrCountry"),"'","''")%>">
                  <input name="chrARName" type="hidden" value="<%=replace(request("chrARName"),"'","''")%>">
                  <input name="chrAREmail" type="hidden" value="<%=replace(request("chrAREmail"),"'","''")%>">
                  <input name="chrOSPerson" type="hidden" value="<%=replace(request("chrOSPerson"),"'","''")%>">
                  <input name="chrOSEmail" type="hidden" value="<%=replace(request("chrOSEmail"),"'","''")%>">
                  <input name="chrOSPhone" type="hidden" value="<%=replace(request("chrOSPhone"),"'","''")%>">
                  <input name="chrOSFax" type="hidden" value="<%=replace(request("chrOSFax"),"'","''")%>">
                  <input name="chrCarrier" type="hidden" value="<%=replace(request("chrCarrier"),"'","''")%>">
                  <input name="chrAccount" type="hidden" value="<%=replace(request("chrAccount"),"'","''")%>">
                  <input name="idPurpose" type="hidden" value="<%=request("idPurpose")%>">
                  <input name="idExpedite" type="hidden" value="<%=request("idExpedite")%>">
                  <input name="txtNotes" type="hidden" value="<%=replace(request("txtNotes"),"'","''")%>">
                  <input name="txtShippingNotes" type="hidden" value="<%=replace(request("txtShippingNotes"),"'","''")%>">
                  <input name="idType" type="hidden" value="<%=request("idType")%>">
                  <input name="idCustomer" type="hidden" value="<%=request("idCustomer")%>">
                  <input name="idUser" type="hidden" value="<%=request("idUser")%>">
                  <input name="idRStatus" type="hidden" value="<%=request("idRStatus")%>">
                  <input name="chrSearch" type="hidden" value="<%=request("chrSearch")%>">
                  </td>
                </tr>
                <tr> 
                  <td height="50"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table>
            </td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="10" height="10" align="left" valign="bottom"><img src="images/bottomleftblue.gif" width="10" height="10"></td>
            <td width="780" height="10"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td align="center"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Copyright © 2004 
                    techIT Solutions LLC. <br>
                    Asset Management Enterprise Portal 4.0 &amp; Corporate Business 
                    Intelligence as products of techIT Solutions. </font></td>
                </tr>
              </table></td>
            <td width="10" height="10" align="right" valign="bottom"><img src="images/bottomrightblue.gif" width="10" height="10"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
    </tr>
  </table>
</form>
</body>
</html>
<%
	rsAssets.Close
	set rsAssets = nothing
	dbConnection.Close
	set dbConnection = nothing
%>