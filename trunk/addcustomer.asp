<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on Settings button
	buttonswitch = 5
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
<form name="form1" method="post" action="insertcustomer.asp">
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
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Add Customer </strong></font></td>
                        <td width="50%">&nbsp;</td>
                      </tr>
                      <tr bgcolor="#f5f5f5"> 
                        <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">To add a customer, please enter the information below and click the Add Customer button.</font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr> 
                        <td><font size="2" face="Arial, Helvetica, sans-serif"><strong>Customer Information</strong></font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#5b5b5b">
						  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr bgcolor="#ffffff"> 
                              <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">Customer ID (example: 10 or 1 or 99)<br>
                                <input name="idCustomer" type="text" id="idCustomer" size="4" maxlength="3"> 
                              </font></td>
                            </tr>
                            <tr bgcolor="#ffffff">
                              <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">Company Name (example: Apple Corporate Events or Adobe Corporate Events)<br>
                                  <input name="chrCustomer" type="text" id="chrCustomer" size="50" maxlength="50">
                              </font></td>
                            </tr>
                            <tr bgcolor="#ffffff"> 
                              <td width="50%" bgcolor="#f5f5f5"><strong><font size="2" face="Arial, Helvetica, sans-serif">Contact Information</font></strong></td>
                              <td width="50%" bgcolor="#f5f5f5"><strong><font size="2" face="Arial, Helvetica, sans-serif">Billing Information</font></strong></td>
                            </tr>
                            <tr bgcolor="#ffffff"> 
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Person's Name<br>
                                <input name="chrCName" type="text" id="chrCName" size="35" maxlength="75"></font></td>
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Person's Name<br>
                                <input name="chrBName" type="text" id="chrBName" size="35" maxlength="75"></font></td>
                            </tr>
                            <tr bgcolor="#ffffff"> 
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Email<br>
                                <input name="chrCEmail" type="text" id="chrCEmail" size="35" maxlength="150"></font></td>
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Email<br>
                                <input name="chrBEmail" type="text" id="chrBEmail" size="35" maxlength="150"></font></td>
                            </tr>
                            <tr bgcolor="#ffffff"> 
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Address<br>
                                <input name="chrCAddress" type="text" id="chrCAddress" size="35" maxlength="100"></font></td>
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Address<br>
                                <input name="chrBAddress" type="text" id="chrBAddress" size="35" maxlength="100"></font></td>
                            </tr>
                            <tr bgcolor="#ffffff"> 
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Address<br>
                                <input name="chrCAddress2" type="text" id="chrCAddress2" size="35" maxlength="100"></font></td>
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Address<br>
                                <input name="chrBAddress2" type="text" id="chrBAddress2" size="35" maxlength="100"></font></td>
                            </tr>
                            <tr bgcolor="#ffffff"> 
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">City<br>
                                <input name="chrCCity" type="text" id="chrCCity" size="35" maxlength="75"></font></td>
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">City<br>
                                <input name="chrBCity" type="text" id="chrBCity" size="35" maxlength="75"></font></td>
                            </tr>
                            <tr bgcolor="#ffffff"> 
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">State<br>
                                <input name="chrCState" type="text" id="chrCState" size="4" maxlength="2"></font></td>
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">State<br>
                                <input name="chrBState" type="text" id="chrBState" size="4" maxlength="2"></font></td>
                            </tr>
                            <tr bgcolor="#ffffff"> 
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Zip<br>
                                <input name="chrCZip" type="text" id="chrCZip" size="12" maxlength="10"></font></td>
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Zip<br>
                                <input name="chrBZip" type="text" id="chrBZip" size="12" maxlength="10"></font></td>
                            </tr>
                            <tr bgcolor="#ffffff"> 
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Phone (example: 408-555-1212)<br>
                                <input name="chrCPhone" type="text" id="chrCPhone" size="16" maxlength="14"></font></td>
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Phone (example: 408-555-1212)<br>
                                <input name="chrBPhone" type="text" id="chrBPhone" size="16" maxlength="14"></font></td>
                            </tr>
                            <tr bgcolor="#ffffff"> 
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Fax (example: 408-555-1313)<br>
                                <input name="chrCFax" type="text" id="chrCFax" size="16" maxlength="14"></font></td>
                              <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Fax (example: 408-555-1212)<br>
                                <input name="chrBFax" type="text" id="chrBFax" size="16" maxlength="14"></font></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr>
                  <td><input name="Submit" type="submit" onClick="YY_checkform('form1','idCustomer','#1_999','1','Please enter a Customer ID.','chrCustomer','#q','0','Please enter the Customer Name.','chrCName','#q','0','Please enter the Contact Name.','chrBName','#q','0','Please enter the Billing Name.','chrCEmail','#q','0','Please enter the Contact Email Address.','chrBEmail','#q','0','Please enter the Billing Email Address.','chrCAddress','#q','0','Please enter the Contact Address.','chrBAddress','#q','0','Please enter the Billing Address.','chrCCity','#q','0','Please enter the Contact City.','chrBCity','#q','0','Please enter the Billing City.','chrCState','#q','0','Please enter the Contact State.','chrBState','#q','0','Please enter the Billing State.','chrCZip','#q','0','Please enter the Contact Zip Code.','chrBZip','#q','0','Please enter the Billing Zip Code.','chrCPhone','#q','0','Please enter the Contact Phone.','chrBPhone','#q','0','Please enter the Billing Phone.','chrCFax','#q','0','Please enter the Contact Fax.','chrBFax','#q','0','Please enter the Billing Fax.');return document.MM_returnValue" value="Add Customer"></td>
                </tr>
                <tr>
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
          </tr>
        </table>
	  </td>
    </tr>
    <!-- #Begin bottom part -->
    <!-- #include file="includes/bottom.htm" -->
  </table>
  <script language="JavaScript">
		document.form1.chrCustomer.focus()
  </script>
  </form>
</body>
</html>