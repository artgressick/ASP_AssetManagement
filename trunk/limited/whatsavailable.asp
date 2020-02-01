<%@ Language=VBScript %>
<%
	if session("idUser") = "" then
		Response.Redirect "../logon.asp"
	end if
%>
<!-- #include file="../includes/openconn.asp" -->
<%	
	'what's available
	if request("dtPull") = "" or request("dtTurn") = "" then
		errorflag = 1
	else
		errorflag = 0
		set rsInventory = server.CreateObject("adodb.recordset")
		sql = "execute WhatsAvailablebyUser " & session("idUser") & "," & request("idCustomer") & "," & request("idCategory") & ",'" & dateadd("d",-2,request("dtPull")) & "','" & dateadd("d",5,request("dtTurn")) & "'"
		set rsInventory = dbConnection.Execute(sql)
	end if
	
	'List the categories
	set rsCategories = server.CreateObject("adodb.recordset")
	sql = "execute ListCategoriesbyUser " & session("idUser")
	set rsCategories = dbConnection.Execute(sql)
	
	'List the customers
	set rsCustomers = server.CreateObject("adodb.recordset")
	sql = "execute ListCustomerNamesandIDsbyAccess " & session("idUser")
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
   mytxt+="<td>"+yyfnt+"FR</font></td><td>"+yyfnt+"SA</font></td><td>"+yyfnt+"<font color=red>SU</font></font></td></tr>"+myTR;
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
<body>
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
<form name="form1" method="post" action="whatsavailable.asp">
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
    <!-- #include file="includes/top.htm" -->
    <tr> 
      <td width="10" background="images/leftverticalline.gif"><img src="images/leftverticalline.gif" width="10" height="10"></td>
      <td width="780">
		<table width="780" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td>
			  <table width="100%" border="0" cellspacing="0" cellpadding="3">
                <tr> 
                  <td width="50%"><font size="4" face="Arial, Helvetica, sans-serif"><strong>What's Available</strong></font></td>
                  <td width="50%" align="right"><font size="2" face="Arial, Helvetica, sans-serif"><a href="default.asp">Return Home</a> </font></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="1" bgcolor="#6699cc">
			  <table width="100%" border="0" cellspacing="1" cellpadding="3">
                <tr> 
                  <td bgcolor="#f5f5f5"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Begin Date</strong></font><br>
 				    <font size="1" face="Arial, Helvetica, sans-serif">Please enter the date the equipment needs to Arrive at your location minus 3 days to
					allow for prep and shipping. We recommend subtracting 4 to 5 days from the date on which you need the assets.</font> 
                    <p><font size="2" face="Arial, Helvetica, sans-serif"><strong>End Date</strong></font><br>
					<font size="1" face="Arial, Helvetica, sans-serif">Please enter the date on which your order will begin returning plus 5 days for shipping and cleaning.
					We recommend adding 7 days to the date on which you will ship back the assets.</font></p></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
          <tr> 
            <td>
			  <table width="100%" border="0" cellspacing="0" cellpadding="1">
                <tr> 
                  <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr bgcolor="#f5f5f5"> 
                        <td align="right"><font size="1" face="Arial, Helvetica, sans-serif">Begin Date</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><input name="dtPull" type="text" id="dtPull" size="10" maxlength="10" value="<%=request("dtPull")%>">&nbsp;&nbsp;<a href="#" onClick="YY_Calendar('dtPull',150,265,'de','#000000','#f5f5f5','YY_calendar1')"><img src="../images/calendar.gif" width="34" height="21" border="0"></a></font></td>
                        <td align="right"><font size="1" face="Arial, Helvetica, sans-serif">End Date</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><input name="dtTurn" type="text" id="dtTurn" size="10" maxlength="10" value="<%=request("dtPull")%>">&nbsp;&nbsp;<a href="#" onClick="YY_Calendar('dtTurn',350,265,'de','#000000','#f5f5f5','YY_calendar1')"><img src="../images/calendar.gif" width="34" height="21" border="0"></a></font></td>
						<td><font size="1" face="Arial, Helvetica, sans-serif"><select name="idCustomer" size="1" id="idCustomer">
                            <option value="0" <%if cint(request("idCustomer")) = 0 then%>selected<%end if%>>All Customers</option>
<%
	if not rsCustomers.EOF then
		do until rsCustomers.EOF
%>
                            <option value="<%=rsCustomers("idCustomer")%>" <%if cint(request("idCustomer")) = rsCustomers("idCustomer") then%>selected<%end if%>><%=trim(rsCustomers("chrCustomer"))%></option>
<%
		rsCustomers.MoveNext
		loop
	end if
%>
                          </select></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><select name="idCategory" size="1" id="idCategory">
                            <option value="0" <%if cint(request("idCategory")) = 0 then%>selected<%end if%>>All Categories</option>
<%
	if not rsCategories.EOF then
		do until rsCategories.EOF
%>
                            <option value="<%=rsCategories("idCategory")%>" <%if cint(request("idCategory")) = rsCategories("idCategory") then%>selected<%end if%>><%=trim(rsCategories("chrCategory"))%></option>
<%
		rsCategories.MoveNext
		loop
	end if
%>
                          </select></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><input name="Submit" type="submit" onClick="YY_checkform('form1','dtPull','#q','0','Please enter a Begin Date.','dtTurn','#q','0','Please enter an End Date.');return document.MM_returnValue" value="Find"></font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
<%
	if errorflag = 0 then
%>
          <tr> 
            <td>
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr bgcolor="#6699cc"> 
                  <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Asset #</font></td>
                  <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Serial #</font></td>
                  <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item Description</font></td>
                  <td height="20"><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Current Location</font></td>
                </tr>
<%
	if rsInventory.EOF then
%>
                <tr align="center"> 
                  <td height="20" colspan="4"><font size="1" face="Arial, Helvetica, sans-serif">Your search criteria has retuned no Assets.</font></td>
                </tr>
                <tr bgcolor="#6699cc"> 
                  <td height="1" colspan="4"><img src="images/6699ccdot.gif" width="1" height="1"></td>
                </tr>
<%
	else
		do until rsInventory.EOF
		counter = counter + 1
		if bgswitch = 1 then
			bgcolor = "#ffffff"
			bgswitch = 0
		else
			bgcolor = "#f5f5f5"
			bgswitch = 1
		end if
%>
                <tr bgcolor="<%=bgcolor%>"> 
                  <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=rsInventory("chrAssNum")%> - <%=counter%></font></td>
                  <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrSerialNum"))%></font></td>
<%
		if rsInventory("chrType") = "C" then
%>
                  <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrItem")) & " - " & trim(rsInventory("chrProcessor"))%><br>
                  &nbsp;<%=trim(rsInventory("chrMemory")) & " - " & trim(rsInventory("chrODrive")) & " - " & trim(rsInventory("chrHDD"))%></font></td>
<%
		else
%>
                  <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrItem"))%></font></td>
<%
		end if
		if rsInventory("idCart") = "0" then
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
                <tr> 
                  <td height="1" colspan="4" bgcolor="#6699cc"><img src="images/6699ccdot.gif" width="1" height="1"></td>
                </tr>
<%
			rsInventory.MoveNext
		loop
		rsInventory.Close
		set rsInventory = nothing
	end if
%>
              </table>
            </td>
          </tr>
<%
	end if
%>
          <tr> 
            <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
          </tr>
        </table>
      </td>
      <td width="10" background="images/rightverticalline.gif"><img src="images/rightverticalline.gif" width="10" height="10"></td>
    </tr>
    <!-- #include file="includes/bottom.htm" -->
  </table>
</form>
</body>
</html>
<%
	rsCategories.Close
	set rsCategories = nothing
	dbConnection.Close
	set dbConnection = nothing
%>