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
	
	'Get a list of the Addresses
	set rsCustomer = server.CreateObject("adodb.recordset")
	sql = "execute FindCustomerbyID " & request("idCustomer")
	set rsCustomer = dbConnection.Execute(sql)
	
	'Get a list of the Carriers
	set rsCarriers = server.CreateObject("adodb.recordset")
	sql = "execute ListSavedCarriersbyUser " & session("idUser")
	set rsCarriers = dbConnection.Execute(sql)
	
	'tabindex primer
	tabindex = 0
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
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="document.form1.chrOrder.focus()">
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
<form name="form1" method="post" action="insertorder.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #include file="includes/top.htm" -->
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
                        <td><strong><font size="3" face="Arial, Helvetica, sans-serif">Add Order to <%=trim(rsCustomer("chrCustomer"))%></font></strong></td>
                      </tr>
                      <tr bgcolor="#f5f5f5"> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Please enter the requested information below to add a new order. &nbsp;All fields with <font color="#FF0000" size="2">*</font> are required and all information must be entered as per the examples given.</font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="#5b5b5b">
					<table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Order/Cart Information</strong><br>
                        <font size="1">Please enter the information about the Order Name, Cart Name and Billing Information.</font></font></td>
                      </tr>
                    </table>
				  </td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="100%" colspan="2"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Order Name (Event Name or Company)</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrOrder" type="text" id="chrOrder" size="40" maxlength="75" tabindex="1"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Cart Name (Booth or Building Number)</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrCart" type="text" id="chrCart" size="40" maxlength="75" tabindex="2"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Your Manager<br>
                          <input name="chrManager" type="text" id="chrManager" size="30" maxlength="50" tabindex="5"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Your Division / Department <U>Name</U></font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrDDName" type="text" id="chrDDName" size="35" maxlength="50" tabindex="3"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Person to Receive Invoice (if different than on account)<br>
                          <input name="chrIPerson" type="text" id="chrIPerson" size="30" maxlength="50" tabindex="6"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Your Division / Department <U>Number</U></font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrDDNumber" type="text" id="chrDDNumber" size="25" maxlength="25" tabindex="4"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Email of Person Receiving Invoice<br>
                          <input name="chrIEmail" type="text" id="chrIEmail" size="35" maxlength="100" tabindex="7"></font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
<%
	'ordering java string
	ordercartjavastring = "'chrOrder','#q','0','Please enter an Order Name.','chrCart','#q','0','Please enter a Cart Name.','chrDDName','#q','0','Please enter the Division/Department Name.','chrDDNumber','#q','0','Please enter the Division/Department Number.'"
	'reset the tabindex
	tabindex = tabindex+7
%>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="#5b5b5b">
					<table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Date Information</strong><br>
                        <font size="1">This is the date information. If you are adding a Standard Order then you will need to provide a beging date and end date for your assets.
                        If you are able to add Internal Use or Out of System Order then the system will calculate the end date.</font></font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
<%
	if request("idType") = 1 then
%>
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Arrival Date (example: 01/01/2003)</font><br>
                          <font size="1" face="Arial, Helvetica, sans-serif"><input name="dtArrival" type="text" id="dtArrival" size="12" maxlength="10" tabindex="8">&nbsp;&nbsp;<a href="#" onClick="YY_Calendar('dtArrival',335,330,'de','#000000','#f5f5f5','YY_calendar1')"><img src="images/calendar.gif" width="34" height="21" border="0"></a></font></td>
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">End / Departure Date (example: 01/01/2003)</font><br>
                          <font size="1" face="Arial, Helvetica, sans-serif"><input name="dtDeparture" type="text" id="dtDeparture" size="12" maxlength="10" tabindex="10">&nbsp;&nbsp;<a href="#" onClick="YY_Calendar('dtDeparture',387,330,'de','#000000','#f5f5f5','YY_calendar1')"><img src="images/calendar.gif" width="34" height="21" border="0"></a></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Arrival Time (if applicable) (example: 12:01 PM)</font><br>
                          <font size="1" face="Arial, Helvetica, sans-serif"><input name="dtArrivalTime" type="text" id="dtArrivalTime" size="10" maxlength="8" tabindex="9"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">End / Departure Time (if applicable) (example: 9:00 AM)</font><br>
                          <font size="1" face="Arial, Helvetica, sans-serif"><input name="dtDepartureTime" type="text" id="dtDepartureTime" size="10" maxlength="8" tabindex="11"></font></td>
                      </tr>
<%
	if session("idAccess") < "P" then
%>
                      <tr> 
                        <td width="100%" colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">Since you will be creating your first cart, does this cart need to have <font color="#0000ff">assets that ship directly from another show</font>?</font><br>
                          <font size="1" face="Arial, Helvetica, sans-serif"><select name="idShow2Show" size="1" id="idShow2Show" tabindex="12">
                            <option value="0" selected>No, this cart will NOT contain any assets from another show.</option>
                            <option value="1">Yes, this cart will contain ONLY assets from another show.</option>
                          </select></font></td>
                      </tr>
<%
	end if
	'date java string
	datejavastring = "'dtArrival','#q','0','Please enter an Arrival Time','dtDeparture','#q','0','Please enter an End/Departure Time'"
	'reset the tabindex
	tabindex = 12
	else
%>
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">Arrival Date (example: 01/01/2003)</font><br>
                          <font size="1" face="Arial, Helvetica, sans-serif"><input name="dtArrival" type="text" id="dtArrival" size="12" maxlength="10" tabindex="8">&nbsp;&nbsp;<a href="#" onClick="YY_Calendar('dtArrival',335,330,'de','#000000','#f5f5f5','YY_calendar1')"><img src="images/calendar.gif" width="34" height="21" border="0"></a></font></td>
                        <td width="50%"><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">End / Departure Date (example: 01/01/2003)</font><br>
                          <font size="2" face="Arial, Helvetica, sans-serif" color="#0000ff"><strong>Undefined</strong></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">Arrival Time (if applicable) (example: 12:01 PM)</font><br>
                          <font size="1" face="Arial, Helvetica, sans-serif"><input name="dtArrivalTime" type="text" id="dtArrivalTime" size="10" maxlength="8" tabindex="9"><input name="idShow2Show" type="hidden" id="idShow2Show" value="0"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">End / Departure Time (if applicable) (example: 9:00 AM)</font><br>
                          <font size="2" face="Arial, Helvetica, sans-serif" color="#0000ff"><strong>Undefined</strong></font></td>
                      </tr>
<%
	'date javastring
	datejavastring = "'dtArrival','#q','0','Please enter an Arrival Time'"
	'reset the tabindex
	tabindex = 9
	end if
%>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="#5b5b5b">
					<table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Ship to Information</strong><br>
						<font size="1">Saving this information will make it available as a Saved Shipping Address for future orders and carts.</font></font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
<%
	if not rsAddresses.EOF then
%>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><input type="radio" tabindex="<%=tabindex+1%>" name="ckAddress" value="0"></td>
                        <td width="100%"><font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif"><strong>Your saved Shipping Addresses:</strong></font></td>
                      </tr>
                      <tr> 
                        <td><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td width="100%" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"> 
                          <select name="idAddress" size="1" id="idAddress" tabindex="<%=tabindex+2%>">
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
	'reset the tabindex
	tabindex = tabindex+2
	addressflag = 0
	end if
%>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><input name="ckAddress" tabindex="<%=tabindex+1%>" type="radio" value="1" checked></td>
                        <td width="100%"><font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif"><strong>New Shipping Address:</strong></font></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td width="100%" bgcolor="#f5f5f5"> <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">To (Booth / Receiving Party)<br>
                                <input name="chrAddress" type="text" id="chrAddress" size="55" maxlength="75" tabindex="<%=tabindex+2%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Address (Company / Hotel / Venue)<br>
                                <input name="chrAddress2" type="text" id="chrAddress2" size="55" maxlength="75" tabindex="<%=tabindex+3%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Address (PO Box / Street Address)<br>
                                <input name="chrAddress3" type="text" id="chrAddress3" size="55" maxlength="75" tabindex="<%=tabindex+4%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font size="1" face="Arial, Helvetica, sans-serif">Address (c/o or additional information)<br>
                                <input name="chrAddress4" type="text" id="chrAddress4" size="55" maxlength="75" tabindex="<%=tabindex+5%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">City<br>
                                <input name="chrCity" type="text" id="chrCity" size="45" maxlength="75" tabindex="<%=tabindex+6%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">State / Province<br>
                                <input name="chrState" type="text" id="chrState" size="22" maxlength="20" tabindex="<%=tabindex+7%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Zip Code<br>
                                <input name="chrZip" type="text" id="chrZip" size="20" maxlength="14" tabindex="<%=tabindex+8%>"></font></td>
                            </tr>
                            <tr> 
                              <td><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Country<br>
                                <input name="chrCountry" type="text" id="chrCountry" value="United States" size="35" maxlength="35" tabindex="<%=tabindex+9%>"></font></td>
                            </tr>
                            <tr> 
                              <td bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif"> 
                                <input name="idSaveAddress" type="checkbox" id="idSaveAddress" value="YES" tabindex="<%=tabindex+10%>">&nbsp;Save this address. &nbsp;&nbsp;&nbsp;Saved Name&nbsp; 
                                <input name="chrSavedAddressName" type="text" id="chrSavedAddressName" size="25" maxlength="50" tabindex="<%=tabindex+11%>">&nbsp;(example: My Office or The Javits.)</font></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
<%
	'address java string
	if addressflag = 1 then
		addressjavastring = "'chrAddress','#q','0','Please enter an Address.','chrAddress2','#q','0','Please enter an Address on line 2.','chrCity','#q','0','Please enter the City.','chrState','#q','0','Please enter the State.','chrZip','#q','0','Please enter the Zip Code.'"
	else
		addressjavastring = "'ckAddress[1]','chrAddress','1','Please enter an Address.','ckAddress[1]','chrAddress2','1','Please enter an Address on line 2.','ckAddress[1]','chrCity','1','Please enter the City.','ckAddress[1]','chrState','1','Please enter the State.','ckAddress[1]','chrZip','1','Please enter the Zip Code.'"
	end if
	'reset the tabindex
	tabindex = tabindex+11
%>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="#5b5b5b"><table width="100%" border="0" cellspacing="1" cellpadding="3">
                      <tr> 
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Contact Information</strong><br>
						<font size="1">This information includes people who will receive Loaner Agreements and emails for shipments.</font></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="100%" colspan="2"><font color="#0000ff" size="1" face="Arial, Helvetica, sans-serif">If needed by the Pool Manager who should receive a loaner agreement?</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <select name="idSendLoaner" size="1" tabindex="<%=tabindex+1%>" id="idSendLoaner">
						  	<option value="0">Please Choose</option>
							<option value="1">Myself</option>
							<option value="2">Apple Requestor</option>
							<option value="3">Onsite Contact</option>
						  </select></font></td>
                      </tr>
					  <tr> 
                        <td width="50%"><font color="#0000ff" size="1" face="Arial, Helvetica, sans-serif">Apple Requestor Name</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrARName" type="text" id="chrARName" size="40" maxlength="50" tabindex="<%=tabindex+2%>"></font></td>
                        <td width="50%"><font color="#0000ff" size="1" face="Arial, Helvetica, sans-serif">Apple Requestor Email</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrAREmail" type="text" id="chrAREmail" size="40" maxlength="100" tabindex="<%=tabindex+6%>"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">On-Site Contact Name</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrOSPerson" type="text" id="chrOSPerson" size="40" maxlength="50" tabindex="<%=tabindex+3%>"></font></td>
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">On-Site Contact Phone (example: 408 555-1212)</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrOSPhone" type="text" id="chrOSPhone" size="20" maxlength="14" tabindex="<%=tabindex+7%>"></font></td>
                      </tr>
                      <tr> 
                        <td width="50%"><font color="#FF0000" size="2">*</font><font color="#000000" size="1" face="Arial, Helvetica, sans-serif">On-Site Contact Email</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <input name="chrOSEmail" type="text" id="chrOSEmail" size="40" maxlength="100" tabindex="<%=tabindex+4%>"></font></td>
                        <td width="50%"><font size="1" face="Arial, Helvetica, sans-serif">On-Site Contact Fax Number (example: 408 555-1212)<br>
                          <input name="chrOSFax" type="text" id="chrOSFax" size="20" maxlength="14" tabindex="<%=tabindex+8%>"></font></td>
                      </tr>
					  <tr> 
                        <td width="50%"><font color="#0000ff" size="1" face="Arial, Helvetica, sans-serif">Send email to Onsite Contact prior to returning</font><font size="1" face="Arial, Helvetica, sans-serif"><br>
                          <select name="idEmailPrior" size="1" tabindex="<%=tabindex+5%>" id="idEmailPrior">
						  	<option value="1">Yes</option>
							<option value="2">No</option>
						  </select></font></td>
                        <td width="50%"><font color="#0000ff" size="1" face="Arial, Helvetica, sans-serif">Send email to Onsite Contact when 10 days late.<br>
                          <select name="idEmailLate" size="1" tabindex="<%=tabindex+9%>" id="idEmailLate">
						  	<option value="1">Yes</option>
							<option value="2">No</option>
						  </select></font></td>
                      </tr>
                    </table></td>
                </tr>
<%
	contactjavastring = "'idSendLoaner','#q','1','Please choose a person to receive a loaner agreement.','chrOSPerson','#q','0','Please enter an On-Site Contact Name.','chrOSPhone','#q','0','Please enter an On-Site Phone number.','chrOSEmail','#q','0','Please enter an On-Site Email address.'"
	'reset the tabindex
	tabindex = tabindex+9
%>
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
	if request("idCustomer") = 6 or request("idCustomer") = 7 or request("idCustomer") = 9 then
%>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Purpose of Loan<br>
                          <select name="idPurpose" size="1" tabindex="<%=tabindex+1%>" id="idPurpose">
						  	<option value="1">Demo</option>
							<option value="2">Event</option>
							<option value="3">Seed</option>
						  </select></font></td>
                      </tr>
<%
	'reset the tabindex
	tabindex = tabindex+1
	end if
%>
					  <tr> 
						<td><font size="1" face="Arial, Helvetica, sans-serif">What will the equipment be used for? &nbsp;How will it be displayed? &nbsp;What properties are being used? &nbsp;Who from Apple will<br> 
                          be on-site? &nbsp;What is the benefit to Apple? &nbsp;What is the ROI? &nbsp;What other equipment will be needed?<br>
                          <textarea name="txtNotes" cols="60" rows="5" wrap="VIRTUAL" id="txtNotes" tabindex="<%=tabindex+1%>"></textarea>
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
                        <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Shipping Method</strong><br>
						<font size="1">Saving this information will make it available as a saved carrier for future orders and carts.</font></font></td>
                      </tr>
                    </table></td>
                </tr>
<%
	'reset the tabindex
	tabindex = tabindex+1
	if not rsCarriers.EOF then
		carrierflag = 0
%>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><input type="radio" tabindex="<%=tabindex+1%>" name="ckCarrier" value="0"></td>
                        <td width="100%"><font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif"><STRONG>Your Saved Carriers:</STRONG></font></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td width="100%" bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"> 
                          <select name="idCarrier" size="1" id="idCarrier" tabindex="<%=tabindex+2%>">
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
	'reset the tabindex
	tabindex = tabindex+2
	end if
%>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><input name="ckCarrier" type="radio" tabindex="<%=tabindex+1%>" value="1" checked></td>
                        <td width="100%"><font color="#0000FF" size="2" face="Arial, Helvetica, sans-serif"><STRONG>New Carrier:</STRONG></font></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td width="100%" bgcolor="#f5f5f5"> <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr> 
                              <td width="50%"><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Carrier Name<br>
                                <input name="chrCarrier" type="text" id="chrCarrier" size="30" maxlength="30" tabindex="<%=tabindex+2%>"></font></td>
                              <td width="50%"><font color="#FF0000" size="2">*</font><font size="1" face="Arial, Helvetica, sans-serif" color="#000000">Carrier Account Number<br>
                                <input name="chrAccount" type="text" id="chrAccount" size="30" maxlength="25" tabindex="<%=tabindex+3%>"></font></td>
                            </tr>
                            <tr> 
                              <td colspan="2" bgcolor="#c0c0c0"><font size="1" face="Arial, Helvetica, sans-serif"> 
                                <input name="idSaveCarrier" type="checkbox" id="idSaveCarrier" value="YES" tabindex="<%=tabindex+4%>">&nbsp;Save this carrier</font></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
<%
	'carrier java string
	if carrierflag = 1 then
		carrierjavastring = "'chrCarrier','#q','0','Please enter the Carrier Name.','chrAccount','#q','0','Please enter the Account Number.'"
	else
		carrierjavastring = "'ckCarrier[1]','chrCarrier','1','Please enter the Carrier Name.','ckCarrier[1]','chrAccount','1','Please enter the Account Number.'"
	end if
	'reset the tabindex
	tabindex = tabindex+4
%>
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
						<td><font size="1" face="Arial, Helvetica, sans-serif">Please put any special needs, setup notes or anything that you need the warehouse to know in the space below. You can add more notes to this order at any time.<br>
                          <textarea name="txtShippingNotes" cols="60" rows="5" wrap="VIRTUAL" id="txtShippingNotes" tabindex="<%=tabindex+1%>"></textarea>
                          </font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><input name="Submit" type="submit" tabindex="<%=tabindex+2%>" onClick="YY_checkform('form1',<%=ordercartjavastring%>,<%=datejavastring%>,<%=addressjavastring%>,'idSaveAddress','chrSavedAddressName','2','Please enter a Name for the Saved Address.',<%=contactjavastring%>,<%=carrierjavastring%>);return document.MM_returnValue" value="Create Order & Cart">
				  <input name="idCustomer" type="hidden" value="<%=request("idCustomer")%>">
				  <input name="idType" type="hidden" value="<%=request("idType")%>">
				  <input name="idSupport" type="hidden" value="<%=request("idSupport")%>"></td>
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
  </form>
</body>
</html>
<%
	rsAddresses.Close
	set rsAddresses = nothing
	rsCarriers.Close
	set rsCarriers = nothing
	rsCustomer.Close
	set rsCustomer = nothing
	dbConnection.Close
	set dbConnection = nothing
%>