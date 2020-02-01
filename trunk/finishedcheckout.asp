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
	'find out if there are any assets left in the cart
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute CheckoutStatsbyCartID " & session("idLoadOut")
	set rsCart = dbConnection.Execute(sql)
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
<form name="form1" method="post" action="updatecheckout.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/orders-nav.htm" -->
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
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Finalize Checkout: <%=trim(rsCart("chrCart"))%></strong></font></td>
						<td align="right"><font size="1" face="Arial, Helvetica, sans-serif"><a href="checkout.asp">Return to Checkout</a></font></td>
                      </tr>
                      <tr>
                        <td bgcolor="#f5f5f5" colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">This screen indicates if you have checked out all assets that were approved.<br>
						If all assets have been checked out, then no assets will be listed.<br>
						Please print the Shipping Manifest and enter the appropriate tracking information.<br>
						If any assets are listed below, they will need to be removed or replaced with another asset.</font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	if rsCart("intOrdered") <> rsCart("intScanned") then
	'get a list of Asset that have not been scanned out
	set rsAssets = server.CreateObject("adodb.recordset")
	sql = "execute ListAssetsNotScannedbyCart " & session("idLoadOut")
	set rsAssets = dbConnection.Execute(sql)
%>
                <tr>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr>
                        <td><font size="2" face="Arial, Helvetica, sans-serif"><strong>Order Incomplete: Missing Assets</strong></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr bgcolor="#6699cc">
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Asset #</font></td>
                        <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item Description</font></td>
                        <td height="20" align="right"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Options&nbsp;</font></td>
                      </tr>
<%
	if not rsAssets.eof then
		do until rsAssets.eof
		if bgswitch = 1 then
			bgcolor = "#ffffff"
			bgswitch = 0
		else
			bgcolor = "#f5f5f5"
			bgswitch = 1
		end if
%>
                      <tr bgcolor="<%=bgcolor%>">
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrAssNum"))%></font></td>
<%
		if rsAssets("chrType") = "C" then
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrItem")) & " - " & trim(rsAssets("chrProcessor"))%><br>
                          &nbsp;<%=trim(rsAssets("chrMemory")) & " - " & trim(rsAssets("chrHDD")) & " - " & trim(rsAssets("chrODrive"))%></font></td>
<%
		else
%>
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsAssets("chrItem"))%></font></td>
<%
		end if
%>
                        <td height="20" align="right"><font size="1" face="Arial, Helvetica, sans-serif"><a href="deletefromcart.asp?idInventory=<%=rsAssets("idInventory")%>&amp;idCart=<%=session("idLoadOut")%>">Remove</a> - <a href="replacefromcart.asp?idInventory=<%=rsAssets("idInventory")%>&amp;idDescription=<%=rsAssets("idDescription")%>">Replace</a>&nbsp;</font></td>
                      </tr>
                      <tr bgcolor="#5b5b5b">
                        <td height="1" colspan="3"><img src="images/5b5b5bdot.gif" width="1" height="1"></td>
                      </tr>
<%
		rsAssets.movenext
		loop
	end if
%>
                    </table></td>
                </tr>
                <tr>
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	else
		'AEG - Find the Cart information
		set rsCartInfo = server.CreateObject("adodb.recordset")
		sql = "execute FindCartbyID " & session("idLoadOut")
		set rsCartInfo = dbConnection.Execute(sql)
%>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"># Reconfigured System<br>
                          <input name="intReconfigured" type="text" id="intReconfigured" size="6" maxlength="4"> (These are systems that you add RAM, Drives, etc.)</font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Special Load<br>
                          <select name="chrSpecialLoad" size="1" id="chrSpecialLoad">
							<option value="0">Please Choose</option>
							<option value="Yes">Yes</option>
							<option value="No">No</option>
						  </select></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Carrier (This will edit the Cart Carrier.)<br>
                          <input name="chrCarrier" type="text" id="chrCarrier" size="35" maxlength="30" value="<%=trim(rsCartInfo("chrCarrier"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Account Number (This will edit the Cart Account Number.)<br>
                          <input name="chrAccount" type="text" id="chrAccount" size="35" maxlength="25" value="<%=trim(rsCartInfo("chrAccount"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Tracking Number (Can be added later in the Staged Carts area.)<br>
                          <input name="chrTracking" type="text" id="chrTracking" size="35" maxlength="25" value="<%=trim(rsCartInfo("chrTracking"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                      </tr>
                      <tr> 
                        <td><input name="Submit" type="submit" onClick="YY_checkform('form1','intReconfigured','#q','0','Please enter the amount of system you reconfigured.','chrSpecialLoad','#q','1','Please select where you installed custom loads.');return document.MM_returnValue" value="Finish Checkout"></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr>
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
		rsCartInfo.Close
		set rsCartInfo = nothing
	end if
%>
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
	rsCart.Close
	set rsCart = nothing
	dbConnection.Close
	set dbConnection = nothing
%>