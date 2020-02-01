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
	'AEG - Find the Cart information by ID
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute FindCartbyID " & request("idCart")
	set rsCart = dbConnection.Execute(sql)
	
	'AEG - List the Cart Statuses
	set rsStatus = server.CreateObject("adodb.recordset")
	sql = "execute ListCartStatus"
	set rsStatus = dbConnection.Execute(sql)
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
<form name="form1" method="post" action="precheckupdatecart.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #include file="includes/top.htm" -->
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
                        <td><strong><font size="3" face="Arial, Helvetica, sans-serif">Edit Cart - <%=trim(rsCart("chrCart"))%></font></strong></td>
                      </tr>
                      <tr bgcolor="#f5f5f5"> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Please modify the information requested below to add a new cart. All fields with <font color="#FF0000" size="2">*</font> are required, and all information must be entered as per the examples given.</font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="#5b5b5b"><table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Cart Information</strong> <font size="1"></font></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="50%"><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Cart Status</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <select name="idStatus" size="1" id="idStatus">
<%
	if not rsStatus.EOF then
		do until rsStatus.EOF
%>
							<option value="<%=rsStatus("idStatus")%>" <%if rsStatus("idStatus") = rsCart("idStatus") then%>selected<%end if%>><%=trim(rsStatus("chrCartStatus"))%></option>
<%
		rsStatus.MoveNext
		loop
	end if
%>
						  </select></font></td>
						<td width="50%"><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Billed?</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <select name="idBilled" size="1" id="idBilled">
							<option value="0" <%if rsCart("idBilled") = false then%>selected<%end if%>>No</option>
							<option value="1" <%if rsCart("idBilled") = true then%>selected<%end if%>>Yes</option>
						  </select></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Cart Name (Event Name or Company)</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrOrder" type="text" id="chrOrder" value="<%=trim(rsCart("chrCart"))%>" size="40" maxlength="75" tabindex="1"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Your Manager<br>
                          <input name="chrManager" type="text" id="chrManager" size="30" maxlength="50" tabindex="4" value="<%=trim(rsCart("chrManager"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Your Division / Department <U>Name</U></font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrDDName" type="text" id="chrDDName" size="35" maxlength="50" tabindex="2" value="<%=trim(rsCart("chrDDName"))%>"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Person to Receive the Invoice (if differnet then on Account)<br>
                          <input name="chrIPerson" type="text" id="chrIPerson" size="30" maxlength="50" tabindex="5" value="<%=trim(rsCart("chrIPerson"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Your Division / Department <U>Number</U></font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrDDNumber" type="text" id="chrDDNumber" size="25" maxlength="25" tabindex="3" value="<%=trim(rsCart("chrDDNumber"))%>"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Email of Person Receiving Invoice<br>
                          <input name="chrIEmail" type="text" id="chrIEmail" size="35" maxlength="100" tabindex="6" value="<%=trim(rsCart("chrIEmail"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td width="100%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Loaner Agreement</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <select name="idLoaner" size="1" id="idLoaner">
							<option value="0" <%if rsCart("idLoaner") = 0 then%>selected<%end if%>>Not Required</option>
							<option value="1" <%if rsCart("idLoaner") = 1 then%>selected<%end if%>>Sent</option>
							<option value="2" <%if rsCart("idLoaner") = 2 then%>selected<%end if%>>Signed and Returned</option></select></font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="#5b5b5b"><table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Date Information</strong><BR>
                        <font size="1">Be very careful when changing this information. Changing the dates may cause the assets to be overbooked.</font></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Pull Date (exmaple: 01/01/2003)</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="dtPull" type="text" id="dtPull" size="12" maxlength="10" value="<%=formatdatetime(rsCart("dtPull"),2)%>"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">End / Departure Date (example: 01/01/2003)</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="dtDeparture" type="text" id="dtDeparture" size="12" maxlength="10" value="<%=formatdatetime(rsCart("dtDeparture"),2)%>"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Ship Date (example: 01/01/2003)<br>
                          <input name="dtShip" type="text" id="dtShip" size="12" maxlength="10" value="<%=formatdatetime(rsCart("dtShip"),2)%>"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">End / Departure Time (if applicable) (example: 9:00 AM)<br>
                          <input name="dtDepartureTime" type="text" id="dtDepartureTime" size="10" maxlength="10" value="<%=formatdatetime(rsCart("dtDepartureTime"),4)%>"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Arrival Date (exmaple: 01/01/2003)</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="dtArrival" type="text" id="dtArrival" size="12" maxlength="10" value="<%=formatdatetime(rsCart("dtArrival"),2)%>"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Return Date (example: 01/01/2003)</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="dtReturn" type="text" id="dtReturn" size="12" maxlength="10" value="<%=formatdatetime(rsCart("dtReturn"),2)%>"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Arrival Time (if applicable) (example: 12:01 PM)<br>
                          <input name="dtArrivalTime" type="text" id="dtArrivalTime" size="14" maxlength="11" value="<%=formatdatetime(rsCart("dtArrivalTime"),4)%>"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Turn Date (example: 01/01/2003)<br>
                          <input name="dtTurn" type="text" id="dtTurn" size="14" maxlength="11" value="<%=formatdatetime(rsCart("dtTurn"),2)%>"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Expedite Level<br>
                          <select name="idExpedite" size="1" id="idExpedite">
                            <option value="0" <%if rsCart("idExpedite") = 0 then%>selected<%end if%>>Regular Shipping</option>
                            <option value="1" <%if rsCart("idExpedite") = 1 then%>selected<%end if%>>Expedite</option>
                            <option value="2" <%if rsCart("idExpedite") = 2 then%>selected<%end if%>>Rush</option>
                          </select></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="#5b5b5b">
					<table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Ship To Information</strong></font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="100%">
						  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">To (Booth / Receiving Party)<br>
                                <input name="chrAddress" type="text" id="chrAddress" size="55" maxlength="75" tabindex="13" value="<%=trim(rsCart("chrAddress"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Address (Company / Hotel/Venue)<br>
                                <input name="chrAddress2" type="text" id="chrAddress2" size="55" maxlength="75" tabindex="14" value="<%=trim(rsCart("chrAddress2"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Address (PO Box / Street Address)<br>
                                <input name="chrAddress3" type="text" id="chrAddress3" size="55" maxlength="75" tabindex="15" value="<%=trim(rsCart("chrAddress3"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Address (c/o or additional information)<br>
                                <input name="chrAddress4" type="text" id="chrAddress4" size="55" maxlength="75" tabindex="16" value="<%=trim(rsCart("chrAddress4"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">City<br>
                                <input name="chrCity" type="text" id="chrCity" size="45" maxlength="75" tabindex="17" value="<%=trim(rsCart("chrCity"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">State / Providence<br>
                                <input name="chrState" type="text" id="chrState" size="22" maxlength="20" tabindex="18" value="<%=trim(rsCart("chrState"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Zip (example: 12345-1234)<br>
                                <input name="chrZip" type="text" id="chrZip" size="20" maxlength="14" tabindex="19" value="<%=trim(rsCart("chrZip"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Country (United States, Japan, France)<br>
                                <input name="chrCountry" type="text" id="chrCountry" size="35" maxlength="35" tabindex="20" value="<%=trim(rsCart("chrCountry"))%>"></font></td>
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
                  <td bgcolor="#5b5b5b"><table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>On-Site Contact Information</strong> <font size="1"></font></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="50%"><font color="#0000ff" size="1" face="Arial, Helvetica, sans-serif">Apple Requestor Name</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrARName" type="text" id="chrARName" size="40" maxlength="50" value="<%=trim(rsCart("chrARName"))%>"></font></td>
                        <td width="50%"><font color="#0000ff" size="1" face="Arial, Helvetica, sans-serif">Apple Requestor Email</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrAREmail" type="text" id="chrAREmail" size="40" maxlength="100" value="<%=trim(rsCart("chrAREmail"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Contact Name</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrOSPerson" type="text" id="chrOSPerson" size="40" maxlength="50" tabindex="23" value="<%=trim(rsCart("chrOSPerson"))%>"></font></td>
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Contact Cell / Business Phone (example: (408) 555-1212)</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrOSPhone" type="text" id="chrOSPhone" size="20" maxlength="14" tabindex="25" value="<%=trim(rsCart("chrOSPhone"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Contact Email</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrOSEmail" type="text" id="chrOSEmail" size="40" maxlength="100" tabindex="24" value="<%=trim(rsCart("chrOSEmail"))%>"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Contact Fax Number (example: (408) 555-1212)<br>
                          <input name="chrOSFax" type="text" id="chrOSFax" size="20" maxlength="14" tabindex="26" value="<%=trim(rsCart("chrOSFax"))%>"></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="#5b5b5b"><table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Purpose of Equipment Loan</strong> <font size="1">(Please be as specific as possible.)</font></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
<%
	'This is for Bobs pools
	if rsCart("idCustomer") = 6 or rsCart("idCustomer") = 7 or rsCart("idCustomer") = 9 then
%>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Purpose of Loan<br>
                          <select name="idPurpose" size="1" id="idPurpose">
						  	<option value="1" <%if rsCart("idPurpose") = "1" Then%>Selected<%end If%>>Demo</option>
							<option value="2" <%if rsCart("idPurpose") = "2" Then%>Selected<%end If%>>Event</option>
							<option value="3" <%if rsCart("idPurpose") = "3" Then%>Selected<%end If%>>Seed</option>
						  </select></font></td>
                      </tr>
<%
	else
%>
						<input type="hidden" name="idPurpose" value="<%=rsCart("idPurpose")%>">
<%	
	end if
%>
					  <tr> 
						<td><font size="1" face="Arial, Helvetica, sans-serif">What will the equipment be used for?&nbsp;&nbsp;How will it be displayed?
                  &nbsp;&nbsp;What properties are being used?<BR>Who from Apple will be on-site?&nbsp;&nbsp;What is the benefit to Apple?&nbsp;&nbsp;What is the ROI?<br>
                          <textarea name="txtNotes" cols="60" rows="5" wrap="VIRTUAL" id="txtNotes" tabindex="27"><%=trim(rsCart("txtNotes"))%></textarea>
                          </font> </td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="#5b5b5b"><table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Shipping Method</strong><font size="1"><br>
						Saving this information will make it available as a Saved Carrier for future Orders and Carts.</font></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="100%"> <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td width="50%"><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Carrier Name<br>
                                <input name="chrCarrier" type="text" id="chrCarrier" size="30" maxlength="30" tabindex="31" value="<%=trim(rsCart("chrCarrier"))%>"></font></td>
                              <td width="50%"><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Carrier Account Number<br>
                                <input name="chrAccount" type="text" id="chrAccount" size="30" maxlength="25" tabindex="32" value="<%=trim(rsCart("chrAccount"))%>"></font></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="#5b5b5b"><table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Setup Notes / Software Loads / Additional Information</strong>&nbsp;<font size="1">(Please be as specific as possible.)</font></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
					  <tr> 
						<td><font size="1" face="Arial, Helvetica, sans-serif">Please put any special needs, setup notes or anything that you need the warehouse to know. You can add more notes to this order at anytime.<br>
                          <textarea name="txtShippingNotes" cols="60" rows="5" wrap="VIRTUAL" id="txtShippingNotes" tabindex="34"><%=trim(rsCart("txtShippingNotes"))%></textarea>
                          </font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	'change the validation java on the webpage
	addressjavastring = "'chrAddress','#q','0','Please enter an Address.','chrAddress2','#q','0','Please enter an Address on line 2.','chrCity','#q','0','Please enter the City.','chrState','#q','0','Please enter the State.','chrZip','#q','0','Please enter the Zip Code.',"
	carrierjavastring = "'chrCarrier','#q','0','Please enter the Carrier Name.','chrAccount','#q','0','Please enter the Account Number.',"
%>
                <tr> 
                  <td><input name="Submit" type="submit" onClick="YY_checkform('form1',<%=addressjavastring%><%=carrierjavastring%>'chrOrder','#q','0','Please enter the Order Name.','chrDDName','#q','0','Please enter the Division/Department Name.','chrDDNumber','#q','0','Please enter the Division/Department Number.','chrOSPerson','#q','0','Please enter an On-Site Contact Name.','chrOSPhone','#q','0','Please enter an On-Site Phone number.','chrOSEmail','#q','0','Please enter an On-Site Email address.');return document.MM_returnValue" value="Update Cart">
				  <input name="idCart" type="hidden" value="<%=request("idCart")%>">
				  <input name="idRStatus" type="hidden" value="<%=request("idStatus")%>">
				  <input name="idUser" type="hidden" value="<%=request("idUser")%>">
				  <input name="idCustomer" type="hidden" value="<%=request("idCustomer")%>">
				  <input name="idType" type="hidden" value="<%=request("idType")%>">
				  <input name="chrSearch" type="hidden" value="<%=request("chrSearch")%>"></td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <!-- #include file="includes/bottom.htm" -->
  </table>
  <script language="JavaScript">
		document.form1.chrOrder.focus()
  </script>
  </form>
</body>
</html>
<%
	rsCart.Close
	set rsCart = nothing
	rsStatus.Close
	set rsStatus = nothing
	dbConnection.Close
	set dbConnection = nothing
%>