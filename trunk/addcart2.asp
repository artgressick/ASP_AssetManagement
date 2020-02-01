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
	'Error Flags
	addressflag = 1 'turn on the flag
	carrierflag = 1 'turn on the flag
	
	'Get a list of the Addresses
	set rsAddresses = server.CreateObject("adodb.recordset")
	sql = "execute ListSavedAddressesbyUser " & session("idUser")
	set rsAddresses = dbConnection.Execute(sql)
	
	'Get a list of the Carriers
	set rsCarriers = server.CreateObject("adodb.recordset")
	sql = "execute ListSavedCarriersbyUser " & session("idUser")
	set rsCarriers = dbConnection.Execute(sql)
	
	'Find the order so that we can populate it.
	set rsOrder = server.CreateObject("adodb.recordset")
	sql = "execute FindOrderbyID " & request("idOrder")
	set rsOrder = dbConnection.Execute(sql)
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

function YYcalclose(YYwhat){//v4.0
  if (YYwhat>=0){
    var yyTag = YYwhat - yyW + 2;
    if ((yyTag > 0)&&(yyTag <= dom[yyDiv.m])){
      var d=yyTag;
      if (YYLang=='de'){YYstrdatum = eval(yyDiv.m+1)+'/'+d+'/'+yyDiv.year;} //4.12.1968
      if (YYLang=='com'){YYstrdatum = YYstrm[yyDiv.m].substring(0,3) +' '+d+', '+yyDiv.year;}
      if (YYLang=='av'){YYstrdatum = d+'/'+YYstrm[yyDiv.m].substring(0,3)+'/'+yyDiv.year;}
      yyDatevar.value=YYstrdatum;
    }
  }
  if (document.layers){yyDiv.visibility = "hide";}
  if (document.all||document.getElementById){yyDiv.style.visibility = "hidden";}
}

function YYgoYear(YY){//v4.0
  var newYear = eval(yyDiv.year) + YY;
  yyDiv.year = newYear.toString(10);
  if (YY==0){} else {YYsetMonth(yyDiv.m,yyDiv.year)}
  setTimeout('YYcaldraw(yyDiv.d,yyDiv.m,yyDiv.year)',(document.layers)?'300':'1');
}

function YYsetMonth(YYm, YYy){//v4.0
  var startDate = new Date();
  startDate.setMonth(YYm);   startDate.setYear(YYy);   startDate.setDate(1);
  yyW = startDate.getDay();
  if (yyW==0){yyW=7}
  var daSchalt = yyDiv.year % 4;
  if (daSchalt==0){dom[1]=29}else {dom[1]=28}
}

function YYgoMonth(YY){//v4.0
   yyDiv.m=yyDiv.m+YY;
   if (yyDiv.m<0){yyDiv.m+=12;YYgoYear(-1)}
     else {if (yyDiv.m>11){yyDiv.m=yyDiv.m-12;YYgoYear(1)}
       else{setTimeout('YYcaldraw(yyDiv.d,yyDiv.m,yyDiv.year)',(document.layers)?'300':'1')}
     }
   YYsetMonth(yyDiv.m,yyDiv.year);
}

function YYsetDate(){//v4.0
   var myDate = new Date();
   yyDiv.year=myDate.getYear();
   if ((myDate.getYear() > 86)&&(myDate.getYear() <= 99)) { yyDiv.year= '19' + myDate.getYear() }
   if ((myDate.getYear() > 99)&&(myDate.getYear() < 1900)) { yyDiv.year= (1900 + myDate.getYear())+''; }
   if (myDate.getYear() <= 86){ yyDiv.year= '20' + myDate.getYear() }//2000!!
   yyDiv.m =  myDate.getMonth();
   yyDiv.d = myDate.getDate();
   var w = myDate.getDay();
   YYsetMonth(yyDiv.m,yyDiv.year);
   YYgoYear(0);
}

function YYcaldraw(ycd,ycm,ycy){//v4.0
  // writing the calendar table
  var yyfnt="<font size=1 color='"+yyDiv.yyTextcolor+"' face=\'Arial, sans-serif\'>";
  var myTR = "<tr bgcolor=\'"+yyDiv.yyBgcolor+"\'>";
  var yyatag="<a href='#' style=\"color: "+yyDiv.yyTextcolor+"; text-decoration: none\" onClick=";
  if (document.layers||document.all||document.getElementById){
   var myMonth = YYstrm[ycm];
   var mytxt="<table bgcolor=\'#000000\' border=\'0\' cellspacing=\'1\' cellpadding=\'3\' width=\'210\'>";
   mytxt+=myTR+"<td colspan='7' align='center'>"+yyfnt+yyatag+"'YYgoMonth(-1)'>&lt;&lt;&nbsp;</a>&nbsp;&nbsp;";
   mytxt+=myMonth;
   mytxt+="&nbsp;&nbsp;"+ycy+"&nbsp;&nbsp;";
   mytxt+=yyatag+"'YYgoMonth(1)'>&nbsp;&gt;&gt;</a></font></td></tr>"+myTR;
   mytxt+="<td>"+yyfnt+"MO</font></td><td>"+yyfnt+"TU</font></td><td>"+yyfnt+"WE</font></td><td>"+yyfnt+"TH</font></td>";
   mytxt+="<td>"+yyfnt+"FR</font></td><td>"+yyfnt+"SA</font></td><td>"+yyfnt+"SU</font></td></tr>"+myTR;
   for (var i=0;i<=41;i++){
     myStr=((i > (dom[ycm]+yyW-2))||(i < yyW-1))?"&nbsp;":i-yyW+2;
     mytxt+="<td>"+yyfnt+yyatag+"\'YYcalclose("+i+")\' title='"+myMonth+" "+myStr+", "+ycy+"'>"+ myStr + "</a></font></td>";
     if ((i==6) || (i==13) || (i==20) || (i==27) || (i==34) || (i==41)) { mytxt+=myTR }
   }
   mytxt+=myTR+"<td colspan='7' align='center'>"+yyfnt+"";
   mytxt+=yyatag+"'YYcalclose()' title='close'>Cancel and Close</a></font></td></tr>";
   mytxt+="</table>";
 }
 if (document.layers){
   with (yyDiv.document){
     open('text/html');
     write(mytxt);
     close();
   }
 }  // end of ns4
 else if (document.all||document.getElementById){
   yyDiv.innerHTML=mytxt;
 } // end of ie4x / dom
}

function YY_Calendar(YYwhat,YYleft, YYtop,YYformat, YYtextcolor, YYbgcolor){//v4.0
  yyDiv= MM_findObj('Calendar1');
  yyDiv.yyTextcolor = YYtextcolor;
  yyDiv.yyBgcolor = YYbgcolor;
  YYsetDate();
  if (document.layers){
    yyDiv.left = YYleft;
    yyDiv.top = YYtop;
    yyDiv.visibility ="show";
  }
  if (document.all){
    yyDiv.style.pixelLeft = YYleft;
    yyDiv.style.pixelTop = YYtop;
    yyDiv.style.visibility = "visible";
  }else
  if (document.getElementById){
    yyDiv.style.left = YYleft;
    yyDiv.style.top = YYtop;
    yyDiv.style.visibility = "visible";
  }
  yyDatevar = MM_findObj(YYwhat);
  YYLang=YYformat;


}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-- #BeginBehavior YY_Calendar1 -->
<script LANGUAGE="JavaScript">
<!--
   var yyDatevar ='YYnull';
   var yyDiv=null;var YYLang='de';
   var dom= new Array(12);
   dom[0]=31;dom[1]=28;dom[2]=31;dom[3]=30;dom[4]=31;dom[5]=30;dom[6]=31;dom[7]=31;dom[8]=30;dom[9]=31;dom[10]=30;dom[11]=31;
   var YYstrm= new Array(12);
   YYstrm[0]='January';YYstrm[1]='February';YYstrm[2]='March';YYstrm[3]='April';YYstrm[4]='May';YYstrm[5]='June';YYstrm[6]='July';
   YYstrm[7]='August'; YYstrm[8]='September';YYstrm[9]='October';YYstrm[10]='November';YYstrm[11]='December';
   
//-->
</script>
<div id="Calendar1" style="position:absolute; left:1px; top:1px; width:200px; height:115px; z-index:20; visibility: hidden"></div>
<!-- #EndBehavior YY_Calendar1 -->
<form name="form1" method="post" action="insertcart.asp">
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
                        <td><strong><font size="3" face="Arial, Helvetica, sans-serif">Add Cart to <%=trim(rsOrder("chrOrder"))%></font></strong></td>
                      </tr>
                      <tr bgcolor="#f5f5f5"> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Please enter the information requested below to add a new cart. All fields with <font color="#FF0000" size="2">*</font> are required, and all information must be entered as per the examples given.</font></td>
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
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Cart Name (Event Name or Company)</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrOrder" type="text" id="chrOrder" size="40" maxlength="75" tabindex="1"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Your Manager<br>
                          <input name="chrManager" type="text" id="chrManager" size="30" maxlength="50" tabindex="4" value="<%=trim(rsOrder("chrManager"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Your Division / Department <U>Name</U></font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrDDName" type="text" id="chrDDName" size="35" maxlength="50" tabindex="2" value="<%=trim(rsOrder("chrDDName"))%>"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Person to Receive the Invoice (if differnet then on Account)<br>
                          <input name="chrIPerson" type="text" id="chrIPerson" size="30" maxlength="50" tabindex="5" value="<%=trim(rsOrder("chrIPerson"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Your Division / Department <U>Number</U></font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrDDNumber" type="text" id="chrDDNumber" size="25" maxlength="25" tabindex="3" value="<%=trim(rsOrder("chrDDNumber"))%>"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Email of Person Receiving Invoice<br>
                          <input name="chrIEmail" type="text" id="chrIEmail" size="35" maxlength="100" tabindex="6" value="<%=trim(rsOrder("chrIEmail"))%>"></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="#5b5b5b"><table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Date Information</strong> <font size="1"></font></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Arrival Date (exmaple: 01/01/2003)</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="dtArrival" type="text" id="dtArrival" size="12" maxlength="10" tabindex="7" value="<%=trim(rsOrder("dtArrival"))%>">&nbsp;&nbsp;<a href="#" onClick="YY_Calendar('dtArrival',335,330,'de','#000000','#f5f5f5','YY_calendar1')"><img src="images/calendar.gif" width="34" height="21" border="0"></a></font></td>
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">End / Departure Date (example: 01/01/2003)</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="dtDeparture" type="text" id="dtDeparture" size="12" maxlength="10" tabindex="9" value="<%=trim(rsOrder("dtDeparture"))%>">&nbsp;&nbsp;<a href="#" onClick="YY_Calendar('dtDeparture',387,330,'de','#000000','#f5f5f5','YY_calendar1')"><img src="images/calendar.gif" width="34" height="21" border="0"></a></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Arrival Time (if applicable) (example: 12:01 PM)<br>
                          <input name="dtArrivalTime" type="text" id="dtArrivalTime" size="10" maxlength="8" tabindex="8"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">End / Departure Time (if applicable) (example: 9:00 AM)<br>
                          <input name="dtDepartureTime" type="text" id="dtDepartureTime" size="10" maxlength="8" tabindex="10"></font></td>
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
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Ship To Information</strong> <font size="1"><br>
						Saving this information will make it available as a Saved Shipping Address for future Orders and Carts.</font></font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
<%
	if not rsAddresses.EOF then
		addressflag = 0
%>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><input type="radio" name="ckAddress" value="0" <%if addressflag = 0 then%>checked<%end if%>></td>
                        <td width="100%"><font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif"><strong>Your saved Shipping Addresses:</strong></font></td>
                      </tr>
                      <tr> 
                        <td><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"> 
                          <select name="idAddress" size="1" id="idAddress" tabindex="12">
<%
		do until rsAddresses.EOF
%>
                            <option value="<%=rsAddresses("idAddress")%>"><%=trim(rsAddresses("chrSavedAddressName"))%>: <%=trim(rsAddresses("chrAddress"))%>, <%=trim(rsAddresses("chrCity"))%>, <%=trim(rsAddresses("chrState"))%></option>
<%
		rsAddresses.MoveNext
		loop
%>
                          </select></font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	end if
%>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><input name="ckAddress" type="radio" value="1" tabindex="11" <%if addressflag = 1 then%>checked<%end if%>></td>
                        <td width="100%"><font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif"><strong>New Shipping Address:</strong></font></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td width="100%" bgcolor="#f5f5f5"> <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">To (Booth / Receiving Party)<br>
                                <input name="chrAddress" type="text" id="chrAddress" size="55" maxlength="75" tabindex="13" value="<%=trim(rsOrder("chrAddress"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Address (Company / Hotel/Venue)<br>
                                <input name="chrAddress2" type="text" id="chrAddress2" size="55" maxlength="75" tabindex="14" value="<%=trim(rsOrder("chrAddress2"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Address (PO Box / Street Address)<br>
                                <input name="chrAddress3" type="text" id="chrAddress3" size="55" maxlength="75" tabindex="15" value="<%=trim(rsOrder("chrAddress3"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Address (c/o or additional information)<br>
                                <input name="chrAddress4" type="text" id="chrAddress4" size="55" maxlength="75" tabindex="16" value="<%=trim(rsOrder("chrAddress4"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">City<br>
                                <input name="chrCity" type="text" id="chrCity" size="45" maxlength="75" tabindex="17" value="<%=trim(rsOrder("chrCity"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">State / Providence<br>
                                <input name="chrState" type="text" id="chrState" size="22" maxlength="20" tabindex="18" value="<%=trim(rsOrder("chrState"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Zip (example: 12345-1234)<br>
                                <input name="chrZip" type="text" id="chrZip" size="20" maxlength="14" tabindex="19" value="<%=trim(rsOrder("chrZip"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Country (United States, Japan, France)<br>
                                <input name="chrCountry" type="text" id="chrCountry" size="35" maxlength="35" tabindex="20" value="<%=trim(rsOrder("chrCountry"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif"> 
                                <input name="idSaveAddress" type="checkbox" id="idSaveAddress" value="YES" tabindex="21">&nbsp;Save this address. &nbsp;&nbsp;&nbsp;Saved Name&nbsp; 
                                <input name="chrSavedAddressName" type="text" id="chrSavedAddressName" size="25" maxlength="50" tabindex="22">&nbsp;(example: My Office or The Javits.)</font></td>
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
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>On-Site Contact Information</strong> <font size="1"></font></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Contact Name</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrOSPerson" type="text" id="chrOSPerson" size="40" maxlength="50" tabindex="23" value="<%=trim(rsOrder("chrOSPerson"))%>"></font></td>
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Contact Cell / Business Phone (example: (408) 555-1212)</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrOSPhone" type="text" id="chrOSPhone" size="20" maxlength="14" tabindex="25" value="<%=trim(rsOrder("chrOSPhone"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Contact Email</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrOSEmail" type="text" id="chrOSEmail" size="40" maxlength="100" tabindex="24" value="<%=trim(rsOrder("chrOSEmail"))%>"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Contact Fax Number (example: (408) 555-1212)<br>
                          <input name="chrOSFax" type="text" id="chrOSFax" size="20" maxlength="14" tabindex="26" value="<%=trim(rsOrder("chrOSFax"))%>"></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="#5b5b5b"><table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Purpose of Equipment Loan</strong>&nbsp;<font size="1">(Please be as specific as possible.)</font></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
<%
	'This is for bob's pools
	if request("idCustomer") = 6 or request("idCustomer") = 7 or request("idCustomer") = 9 then
%>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Purpose of Loan<br>
                          <select name="idPurpose" size="1" id="idPurpose">
						  	<option value="1">Demo</option>
							<option value="2">Event</option>
							<option value="3">Seed</option>
						  </select></font></td>
                      </tr>
<%
	end if
%>
					  <tr> 
						<td><font size="1" face="Arial, Helvetica, sans-serif">What will the equipment be used for?&nbsp;&nbsp;How will it be displayed?
                  &nbsp;&nbsp;What properties are being used?<BR>Who from Apple will be on-site?&nbsp;&nbsp;What is the benefit to Apple?&nbsp;&nbsp;What is the ROI?<br>
                          <textarea name="txtNotes" cols="60" rows="5" wrap="VIRTUAL" id="txtNotes"><%=trim(rsOrder("txtNotes"))%></textarea>
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
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Shipping Method </strong> <font size="1"><br>
						Saving this information will make it available as a Saved Carrier for future Orders and Carts.</font></font></td>
                      </tr>
                    </table></td>
                </tr>
<%
	if not rsCarriers.EOF then
		carrierflag = 0
%>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><input type="radio" name="ckCarrier" value="0" <%if carrierflag = 0 then%>checked<%end if%>></td>
                        <td width="100%"><font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif">Your Saved Carriers:</font></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td width="100%" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"> 
                          <select name="idCarrier" size="1" id="idCarrier" tabindex="30">
<%
		do until rsCarriers.EOF
%>
                            <option value="<%=rsCarriers("idCarrier")%>"><%=trim(rsCarriers("chrCarrier"))%> - <%=trim(rsCarriers("chrAccount"))%></option>
<%
		rsCarriers.MoveNext
		loop
%>
                          </select></font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	end if
%>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><input name="ckCarrier" type="radio" value="1" <%if carrierflag = 1 then%>checked<%end if%>></td>
                        <td width="100%"><font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif">New Carrier:</font></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td width="100%" bgcolor="#f5f5f5"> <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td width="50%"><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Carrier Name<br>
                                <input name="chrCarrier" type="text" id="chrCarrier" size="30" maxlength="30" tabindex="31" value="<%=trim(rsOrder("chrCarrier"))%>"></font></td>
                              <td width="50%"><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Carrier Account Number<br>
                                <input name="chrAccount" type="text" id="chrAccount" size="30" maxlength="25" tabindex="32" value="<%=trim(rsOrder("chrAccount"))%>"></font></td>
                            </tr>
                            <tr> 
                              <td colspan="2" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif"> 
                                <input name="idSaveCarrier" type="checkbox" id="idSaveCarrier" value="YES" tabindex="33">&nbsp;Save this Carrier.</font></td>
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
                          <textarea name="txtShippingNotes" cols="60" rows="5" wrap="VIRTUAL" id="txtShippingNotes" tabindex="34"><%=trim(rsOrder("txtShippingNotes"))%></textarea>
                          </font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	'change the validation java on the webpage
	if addressflag = 1 then
		addressjavastring = "'chrAddress','#q','0','Please enter an Address.','chrAddress2','#q','0','Please enter an Address on line 2.','chrAddress3','#q','0','Please enter an Address on line 3.','chrCity','#q','0','Please enter the City.','chrState','#q','0','Please enter the State.','chrZip','#q','0','Please enter the Zip Code.',"
	else
		addressjavastring = "'ckAddress[1]','chrAddress','1','Please enter an Address.','ckAddress[1]','chrAddress2','1','Please enter an Address on line 2.','ckAddress[1]','chrAddress3','1','Please enter an Address on line 3.','ckAddress[1]','chrCity','1','Please enter the City.','ckAddress[1]','chrState','1','Please enter the State.','ckAddress[1]','chrZip','1','Please enter the Zip Code.',"
	end if
	
	if carrierflag = 1 then
		carrierjavastring = "'chrCarrier','#q','0','Please enter the Carrier Name.','chrAccount','#q','0','Please enter the Account Number.',"
	else
		carrierjavastring = "'ckCarrier[1]','chrCarrier','1','Please enter the Carrier Name.','ckCarrier[1]','chrAccount','1','Please enter the Account Number.',"
	end if
%>
                <tr> 
                  <td><input name="Submit" type="submit" onClick="YY_checkform('form1','idSaveAddress','chrSavedAddressName','2','Please enter a Name for the Saved Address.',<%=addressjavastring%><%=carrierjavastring%>'chrOrder','#q','0','Please enter the Order Name.','chrDDName','#q','0','Please enter the Division/Department Name.','chrDDNumber','#q','0','Please enter the Division/Department Number.','chrOSPerson','#q','0','Please enter an On-Site Contact Name.','chrOSPhone','#q','0','Please enter an On-Site Phone number.','chrOSEmail','#q','0','Please enter an On-Site Email address.','dtArrival','#q','0','Please enter an Arrival Time','dtDeparture','#q','0','Please enter an End/Departure Time');return document.MM_returnValue" value="Add Cart">
				  <input name="idCustomer" type="hidden" value="<%=request("idCustomer")%>"><input name="idOrder" type="hidden" value="<%=request("idOrder")%>"></td>
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
	rsAddresses.Close
	set rsAddresses = nothing
	rsOrder.Close
	set rsOrder = nothing
	rsCarriers.Close
	set rsCarriers = nothing
	dbConnection.Close
	set dbConnection = nothing
%>