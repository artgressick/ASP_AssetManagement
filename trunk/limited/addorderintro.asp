<%@ Language=VBScript %>
<%
	if session("idUser") = "" then
		Response.Redirect "../logon.asp"
	end if
%>
<!-- #include file="../includes/openconn.asp" -->
<%	
	'Get a list of the Customers
	'we need to pull a list of Customer that they can order from. Super user can order from anyone.
	set rsCustomers = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListCustomerNamesandIDs"
	else
		sql = "execute ListCustomerNamesandIDsbyAccess " & session("idUser")
	end if
	set rsCustomers = dbConnection.Execute(sql)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="../includes/title.htm" -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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
<body>
<form name="form1" method="post" action="addorder.asp">
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
                  <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>New Order</strong></font></td>
                  <td width="50%" align="right" valign="bottom"><font size="2" face="Arial, Helvetica, sans-serif">&lt; 
                    <a href="default.asp">Cancel &amp; Return Home</a></font></td>
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
                  <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">From the drop-down list below, please select the Customer / Pool for whom you would like to create an order. Then click the Continue button.</font></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td><font size="1" face="Arial, Helvetica, sans-serif">Customer / Pool<br>
                    <select name="idCustomer" size="1" id="idCustomer">
                      <option value="0">Please Choose a Customer / Pool</option>
<%
	if not rsCustomers.EOF then
		do until rsCustomers.EOF
%>
                      <option value="<%=rsCustomers("idCustomer")%>"><%=trim(rsCustomers("chrCustomer"))%></option>
<%
		rsCustomers.MoveNext
		loop
	end if
%>
                    </select>
                    </font></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td><font size="1" face="Arial, Helvetica, sans-serif">Do you need tech support at your event?<br>
                    <select name="idSupport" size="1" id="idSupport">
                      <option value="0">Please select tech support option</option>
                      <option value="1">Yes, please provide a quote for tech support.</option>
                      <option value="0">No, this order does not require tech support.</option>
                    </select>
                    </font></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr>
                  <td><input name="Submit" type="submit" onClick="YY_checkform('form1','idCustomer','#q','1','Please choose a customer/pool.','idSupport','#q','1','Please select tech support option.');return document.MM_returnValue" value="Continue">
                  <name="idType" type="hidden" value="1" id="idType"></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td height="100"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
        </table></td>
      <td width="10" background="images/rightverticalline.gif"><img src="images/rightverticalline.gif" width="10" height="10"></td>
    </tr>
    <!-- #include file="includes/bottom.htm" -->
  </table>
</form>
</body>
</html>
<%
	rsCustomers.close
	set rsCustomers = nothing
	dbConnection.Close
	set dbConnection = nothing
%>