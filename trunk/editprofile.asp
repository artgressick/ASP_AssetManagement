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
	set rsProfile = server.CreateObject("adodb.recordset")
    'RRF - Determine if need to find Session User or Requested user
	If Request("idUser") = "" Then
	  sql = "execute FindProfilebyIDUser " & session("idUser")
	Else
	  sql = "execute FindProfilebyIDUser " & Request("idUser")
	End If
	set rsProfile = dbConnection.Execute(sql)	
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
<form name="form1" method="post" action="updateprofile.asp">
<%
	'RRF - If statement used to send Hidden Variable to Updateprofile for redirection.
	If Request("idUser") = "" Then
%>
	<input type="hidden" name="usrSession" value="Set">
	<input type="hidden" name="idUser" value="<%=session("idUser")%>">
<%  
	Else
%>
	<input type="hidden" name="idUser" value="<%=Request("idUser")%>">
<%
	End If
	'RRF - End of Statement
%>
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/settings-nav.htm" -->
        </td>
      <td width="100%" height="100%" valign="top"><table width="625" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15"><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="610"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Edit Profile</strong></font></td>
                    <td width="50%">&nbsp;</td>
                  </tr>
                  <tr bgcolor="#f5f5f5">
                    <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">Please make any applicable changes, then click the Update Profile button.<br>
					If no changes are needed, please click the Back button in your browser.</font></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">First Name<br>
                        <input name="chrFirst" type="text" id="chrFirst" size="25" maxlength="25" value="<%=trim(rsProfile("chrFirst"))%>">
                    </font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Last Name<br>
                        <input name="chrLast" type="text" id="chrLast" size="25" maxlength="25" value="<%=trim(rsProfile("chrLast"))%>">
                    </font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Email<br>
                        <input name="chrEmail" type="text" id="chrEmail" size="35" maxlength="150" value="<%=trim(rsProfile("chrEmail"))%>">
                    </font></td>
                  </tr>
                  <tr>
                    <td bgcolor="#f5f5f5"><font size="1" face="Arial, Helvetica, sans-serif">Password<br>
                        <input name="chrPassword" type="password" id="chrPassword" size="12" maxlength="10" value="<%=trim(rsProfile("chrPassword"))%>">
                    </font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Telephone (example: 408-555-1212)<br>
                        <input name="chrPhone" type="text" id="chrPhone" size="16" maxlength="14" value="<%=trim(rsProfile("chrPhone"))%>">
                    </font></td>
                  </tr>
                  <tr>
                    <td><font size="1" face="Arial, Helvetica, sans-serif">Fax (example: 408-555-1212)<br>
                        <input name="chrFax" type="text" id="chrFax" size="16" maxlength="14" value="<%=trim(rsProfile("chrFax"))%>">
                    </font></td>
                  </tr>
                  <tr>
                    <td height="15"><font size="1"><img src="images/ffffffdot.gif" width="1" height="1"></font></td>
                  </tr>
                  <tr>
                    <td height="27"><font size="1">
                      <input name="Submit" type="submit" onClick="YY_checkform('form1','chrFirst','#q','0','First Name must be filled out.','chrLast','#q','0','Last Name must be filled out.','chrEmail','#q','0','Email address must be filled out.','chrPhone','#q','0','Phone number must be filled out.','chrFax','#q','0','Fax number must be filled out.','chrPassword','#q','0','Password must be filled out.');return document.MM_returnValue" value="Update Profile">
                    </font></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
              </tr>
            </table></td>
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
	rsProfile.Close
	set rsProfile = nothing
	dbConnection.Close
	set dbConnection = nothing
%>