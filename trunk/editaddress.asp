<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on order button
	buttonswitch = 2
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Find the Address by ID
	set rsAddress = server.CreateObject("adodb.recordset")
	sql = "execute FindAddressbyID " & request("idAddress")
	set rsAddress = dbConnection.Execute(sql)
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
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="updateaddress.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
		<!-- #include file="includes/orders-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top">
	  	<table width="625" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15"><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="610">
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Edit Address</strong></font></td>
                        <td width="50%">&nbsp;</td>
                      </tr>
                      <tr bgcolor="#f5f5f5"> 
                        <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">Please make your changes and then click the Update Address button.</font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Saved Address Name<br>
                          <input name="chrSavedAddressName" type="text" id="chrSavedAddressName" size="45" maxlength="50" value="<%=trim(rsAddress("chrSavedAddressName"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">To (Booth/Receiving Party)<br>
                          <input name="chrAddress" type="text" id="chrAddress" size="45" maxlength="75" value="<%=trim(rsAddress("chrAddress"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Address (Company/Hotel/Venue)<br>
                          <input name="chrAddress2" type="text" id="chrAddress2" size="45" maxlength="75" value="<%=trim(rsAddress("chrAddress2"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Address (PO Box/Street Address) <br>
                          <input name="chrAddress3" type="text" id="chrAddress3" size="45" maxlength="75" value="<%=trim(rsAddress("chrAddress3"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Address (C/O or additional information)<br>
                          <input name="chrAddress4" type="text" id="chrAddress4" size="45" maxlength="75" value="<%=trim(rsAddress("chrAddress4"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">City<br>
                          <input name="chrCity" type="text" id="chrCity" size="40" maxlength="75" value="<%=trim(rsAddress("chrCity"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">State/Providence<br>
                          <input name="chrState" type="text" id="chrState" size="20" maxlength="20" value="<%=trim(rsAddress("chrState"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Zip<br>
                          <input name="chrZip" type="text" id="chrZip" size="14" maxlength="14" value="<%=trim(rsAddress("chrZip"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Country<br>
                          <input name="chrCountry" type="text" id="chrCountry" size="30" maxlength="35" value="<%=trim(rsAddress("chrCountry"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                      </tr>
                      <tr> 
                        <td><input name="Submit" type="submit" onClick="YY_checkform('form1','chrSavedAddressName','#q','0','Please enter a Saved Address Name.','chrAddress','#q','0','Please enter a To Address.','chrAddress2','#q','0','Please enter an Company/Hotel/Venue address.','chrAddress3','#q','0','Please enter a PO Box/Street Address.','chrCity','#q','0','Please enter a City.','chrState','#q','0','Please enter a State/Providence.','chrZip','#q','0','Please enter a Zip Code.','chrCountry','#q','0','Please enter a Country.');return document.MM_returnValue" value="Update Address">
                        <input type="hidden" name="idAddress" value="<%=request("idAddress")%>"></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
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
  <script language="JavaScript">
		document.form1.chrSavedAddressName.focus()
  </script>
  </form>
</body>
</html>
<%
	rsAddress.Close
	set rsAddress = nothing
	dbConnection.Close
	set dbConnection = nothing
%>