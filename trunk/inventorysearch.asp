<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on inventory button
	buttonswitch = 3
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'prime idCustomer, idInventory Status, idWarehouse
	if request("chrAsset") = "" then
		flag = 1 'don't print
	else
		flag = 0 'ok to print
		set rsInventory = server.CreateObject("adodb.recordset")
		if session("idAccess") < "O" then
			sql = "execute SearchAssets '" & request("chrAsset") & "'"
		else
			sql = "execute SearchAssets '" & request("chrAsset") & "'"
		end if
		set rsInventory = dbConnection.Execute(sql)
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title-meta.htm" -->
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function YY_checkform() { //v4.69
//copyright (c)1998,2002 Yaromat.com
  var a=YY_checkform.arguments,oo=true,v='',s='',err=false,r,o,at,o1,t,i,j,ma,rx,cd,cm,cy,dte,at;
  for (i=1; i<a.length;i=i+4){
    if (a[i+1].charAt(0)=='#'){r=true; a[i+1]=a[i+1].substring(1);}else{r=false}
    o=MM_findObj(a[i].replace(/\[\d+\]/ig,""));
    o1=MM_findObj(a[i+1].replace(/\[\d+\]/ig,""));
    v=o.value;t=a[i+2];
    if (o.type=='text'||o.type=='password'||o.type=='hidden'){
      if (r&&v.length==0){err=true}
      if (v.length>0)
      if (t==1){ //fromto
        ma=a[i+1].split('_');if(isNaN(v)||v<ma[0]/1||v > ma[1]/1){err=true}
      } else if (t==2){
        rx=new RegExp("^[\\w\.=-]+@[\\w\\.-]+\\.[a-z]{2,4}$");if(!rx.test(v))err=true;
      } else if (t==3){ // date
        ma=a[i+1].split("#");at=v.match(ma[0]);
        if(at){
          cd=(at[ma[1]])?at[ma[1]]:1;cm=at[ma[2]]-1;cy=at[ma[3]];
          dte=new Date(cy,cm,cd);
          if(dte.getFullYear()!=cy||dte.getDate()!=cd||dte.getMonth()!=cm){err=true};
        }else{err=true}
      } else if (t==4){ // time
        ma=a[i+1].split("#");at=v.match(ma[0]);if(!at){err=true}
      } else if (t==5){ // check this 2
            if(o1.length)o1=o1[a[i+1].replace(/(.*\[)|(\].*)/ig,"")];
            if(!o1.checked){err=true}
      } else if (t==6){ // the same
            if(v!=MM_findObj(a[i+1]).value){err=true}
      }
    } else
    if (!o.type&&o.length>0&&o[0].type=='radio'){
          at = a[i].match(/(.*)\[(\d+)\].*/i);
          o2=(o.length>1)?o[at[2]]:o;
      if (t==1&&o2&&o2.checked&&o1&&o1.value.length/1==0){err=true}
      if (t==2){
        oo=false;
        for(j=0;j<o.length;j++){oo=oo||o[j].checked}
        if(!oo){s+='* '+a[i+3]+'\n'}
      }
    } else if (o.type=='checkbox'){
      if((t==1&&o.checked==false)||(t==2&&o.checked&&o1&&o1.value.length/1==0)){err=true}
    } else if (o.type=='select-one'||o.type=='select-multiple'){
      if(t==1&&o.selectedIndex/1==0){err=true}
    }else if (o.type=='textarea'){
      if(v.length<a[i+1]){err=true}
    }
    if (err){s+='* '+a[i+3]+'\n'; err=false}
  }
  if (s!=''){alert('The required information is incomplete or contains errors:\t\t\t\t\t\n\n'+s)}
  document.MM_returnValue = (s=='');
}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="inventorysearch.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/reports-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="100%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="595" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Inventory Search</strong></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr>
                        <td bgcolor="#c0c0c0">
                          <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr bgcolor="#f5f5f5"> 
                              <td align="right" width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Asset or Serial Number</font></td>
                              <td><font size="1" face="Arial, Helvetica, sans-serif"><input name="chrAsset" type="text" id="chrAsset" size="20" maxlength="25" value="<%=request("chrAsset")%>"></font></td>
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif"><input name="Submit" type="submit" onClick="YY_checkform('form1','chrAsset','#q','0','Please enter an asset number.');return document.MM_returnValue" value="Search"></font></td>
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
<%
	if flag = 0 then
%>
                <tr> 
                  <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="20" bgcolor="#6699cc"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Asset Number</font></td>
                        <td height="20" bgcolor="#6699cc"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Serial Number</font></td>
                        <td height="20" bgcolor="#6699cc"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item</font></td>
                        <td height="20" bgcolor="#6699cc"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Location</font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="4"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
                      </tr>
<%
		if rsInventory.EOF then
%>
                      <tr> 
                        <td height="20" align="center" colspan="4"><font size="1" face="Arial, Helvetica, sans-serif">There are no Assets with this Asset Number.</font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="4"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
                      </tr>
<%
		else
			do until rsInventory.EOF
			if bgswitch = 1 then
				bgcolor = "#ffffff"
				bgswitch = 0
			else
				bgcolor = "#f5f5f5"
				bgswitch = 1
			end if
%>
                      <tr bgcolor="<%=bgcolor%>"> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<A HREF="viewasset.asp?idInventory=<%=rsInventory("idInventory")%>"><%=trim(rsInventory("chrAssNum"))%></A></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrSerialNum"))%></font></td>
<%
			if rsInventory("chrType") = "C" then
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrItem")) & " - " & trim(rsInventory("chrProcessor"))%><br>
                        &nbsp;<%=trim(rsInventory("chrMemory")) & " - " & trim(rsInventory("chrHDD")) & " - " & trim(rsInventory("chrODrive"))%></font></td>
<%
			else
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrItem"))%></font></td>
<%
			end if
			if rsInventory("idCart") = 0 then
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;Warehouse</font></td>
<%
			else
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrCart"))%></font></td>
<%
			end if
%>                        
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="4"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
                      </tr>
<%
			rsInventory.MoveNext
			loop
		end if
		rsInventory.Close
		set rsInventory = nothing
%>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	end if
%>
              </table></td>
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
          </tr>
        </table></td>
    </tr>
    <!-- #Begin bottom part -->
    <!-- #include file="includes/bottom.htm" -->
  </table>
  <script language="JavaScript">
		document.form1.chrAsset.focus()
  </script>
  </form>
</body>
</html>
<%
	dbConnection.Close
	set dbConnection = nothing
%>