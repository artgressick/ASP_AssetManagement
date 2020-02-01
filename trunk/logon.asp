<%@ Language=VBScript %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!-- #include file="includes/title.htm" -->
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

function YY_checkform() { //v4.66
//copyright (c)1998,2002 Yaromat.com
  var args = YY_checkform.arguments; var myDot=true; var myV=''; var myErr='';var addErr=false;var myReq;
  for (var i=1; i<args.length;i=i+4){
    if (args[i+1].charAt(0)=='#'){myReq=true; args[i+1]=args[i+1].substring(1);}else{myReq=false}
    var myObj = MM_findObj(args[i].replace(/\[\d+\]/ig,""));
    myV=myObj.value;
    if (myObj.type=='text'||myObj.type=='password'||myObj.type=='hidden'){
      if (myReq&&myObj.value.length==0){addErr=true}
      if ((myV.length>0)&&(args[i+2]==1)){ //fromto
        var myMa=args[i+1].split('_');if(isNaN(myV)||myV<myMa[0]/1||myV > myMa[1]/1){addErr=true}
      } else if ((myV.length>0)&&(args[i+2]==2)){
          var rx=new RegExp("^[\\w\.=-]+@[\\w\\.-]+\\.[a-z]{2,4}$");if(!rx.test(myV))addErr=true;
      } else if ((myV.length>0)&&(args[i+2]==3)){ // date
        var myMa=args[i+1].split("#"); var myAt=myV.match(myMa[0]);
        if(myAt){
          var myD=(myAt[myMa[1]])?myAt[myMa[1]]:1; var myM=myAt[myMa[2]]-1; var myY=myAt[myMa[3]];
          var myDate=new Date(myY,myM,myD);
          if(myDate.getFullYear()!=myY||myDate.getDate()!=myD||myDate.getMonth()!=myM){addErr=true};
        }else{addErr=true}
      } else if ((myV.length>0)&&(args[i+2]==4)){ // time
        var myMa=args[i+1].split("#"); var myAt=myV.match(myMa[0]);if(!myAt){addErr=true}
      } else if (myV.length>0&&args[i+2]==5){ // check this 2
            var myObj1 = MM_findObj(args[i+1].replace(/\[\d+\]/ig,""));
            if(myObj1.length)myObj1=myObj1[args[i+1].replace(/(.*\[)|(\].*)/ig,"")];
            if(!myObj1.checked){addErr=true}
      } else if (myV.length>0&&args[i+2]==6){ // the same
            var myObj1 = MM_findObj(args[i+1]);
            if(myV!=myObj1.value){addErr=true}
      }
    } else
    if (!myObj.type&&myObj.length>0&&myObj[0].type=='radio'){
          var myTest = args[i].match(/(.*)\[(\d+)\].*/i);
          var myObj1=(myObj.length>1)?myObj[myTest[2]]:myObj;
      if (args[i+2]==1&&myObj1&&myObj1.checked&&MM_findObj(args[i+1]).value.length/1==0){addErr=true}
      if (args[i+2]==2){
        var myDot=false;
        for(var j=0;j<myObj.length;j++){myDot=myDot||myObj[j].checked}
        if(!myDot){myErr+='* ' +args[i+3]+'\n'}
      }
    } else if (myObj.type=='checkbox'){
      if(args[i+2]==1&&myObj.checked==false){addErr=true}
      if(args[i+2]==2&&myObj.checked&&MM_findObj(args[i+1]).value.length/1==0){addErr=true}
    } else if (myObj.type=='select-one'||myObj.type=='select-multiple'){
      if(args[i+2]==1&&myObj.selectedIndex/1==0){addErr=true}
    }else if (myObj.type=='textarea'){
      if(myV.length<args[i+1]){addErr=true}
    }
    if (addErr){myErr+='* '+args[i+3]+'\n'; addErr=false}
  }
  if (myErr!=''){alert('The required information is incomplete or contains errors:\t\t\t\t\t\n\n'+myErr)}
  document.MM_returnValue = (myErr=='');
}
//-->
</script>
</head>
<body bgcolor="#ffffff" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="document.Logon.chrUsername.focus()">
<form action="checklogon.asp" method="post" name="Logon" id="Logon">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td align="center" valign="middle">
        <table width="620" height="350" border="0" cellpadding="0" cellspacing="0">
          <!-- #include file="includes/logon-top.htm" -->
          <tr>
            <td width="10" height="330" background="images/leftblack.gif"><img src="images/leftblack.gif" width="10" height="10"></td>
            <td width="600" height="330" bgcolor="#f5f5f5"><table width="600" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="150"><p><font size="2" face="Arial, Helvetica, sans-serif"><strong>Account Log On</strong></font></p>
                    <p><font size="1" face="Arial, Helvetica, sans-serif">Username (Email Address)<br>
                      <input name="chrUsername" type="text" id="chrUsername" size="20" maxlength="150" value="<%=Request.Cookies.Item("chrUsername")%>"></font></p>
                    <p><font size="1" face="Arial, Helvetica, sans-serif">Password<br>
                      <input name="chrPassword" type="password" id="chrPassword" size="20" maxlength="10"></font></p>
                    <p><font size="1" face="Arial, Helvetica, sans-serif"><input name="chrRemember" type="checkbox" id="chrRemember" value="Yes">Remember email address</font></p>
                    <p><font size="1" face="Arial, Helvetica, sans-serif"><input name="Submit" type="submit" onClick="YY_checkform('Logon','chrUsername','#q','0','Please enter your Username.','chrPassword','#q','0','Please enter your Password.');return document.MM_returnValue" value="Enter -&gt;"></font></p>
                    <p><font size="1" face="Arial, Helvetica, sans-serif"><a href="lostinformation.asp">Forgot your information?</a></font></p></td>
                  <td width="450"><img src="images/logongraphic.gif" width="450" height="290"></td>
                </tr>
              </table></td>
            <td width="10" height="330" background="images/rightblack.gif"><img src="images/rightblack.gif" width="10" height="10"></td>
          </tr>
          <!-- #include file="includes/logon-bottom.htm" -->
        </table>
      </td>
    </tr>
  </table>
<%
  If request("idUrl") = "" Then
%>
  <input type="hidden" name="idUrl" value="default.asp">
<%
  Else
%>
  <input type="hidden" name="idUrl" value="<%=request("idUrl")%>">
<%
  End if
%>
</form>
</body>
</html>