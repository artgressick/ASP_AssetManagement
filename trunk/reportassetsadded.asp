<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on inventory button
	buttonswitch = 4
%>
<!-- #include file="includes/openconn.asp" -->
<%
	'Get a list of Customers by user.
	set rsCustomers = server.CreateObject("adodb.recordset")
	if session("idAccess") < "O" then
		sql = "execute ListCustomerNamesandIDs"
	else
		sql = "execute ListCustomerNamesandIDsbyAccess " & session("idUser")
	end if
	set rsCustomers = dbConnection.Execute(sql)
	
	if request("idCustomer") = "" then
		idCustomer = 0
	else
		idCustomer = request("idCustomer")
	end if
	
	'Get the list of Assets that have been added
	if request("dtPull") = "" or request("dtTurn") = "" then
		errorflag = 1
	else
		set rsInventory = server.CreateObject("adodb.recordset")
		if session("idAccess") < "O" then
			sql = "execute ListAssetsAddedbyDates " & idCustomer & ",'" & request("dtPull") & "','" & request("dtTurn") & "'"
		else
			sql = "execute ListAssetsAddedbyDatesbyAccess " & idCustomer & "," & session("idUser") & ",'" & request("dtPull") & "','" & request("dtTurn") & "'"
		end if
		set rsInventory = dbConnection.Execute(sql)
	end if
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
function newWindow(updateWin) {
  updateWindow = window.open(updateWin,'updateWin','width=300,height=275');
updateWindow.focus()
}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
<form name="form1" method="post" action="reportassetsadded.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/reports-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="100%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="595" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Assets Added Report</strong></font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="1">
                      <tr>
                        <td bgcolor="#c0c0c0">
                          <table width="100%" border="0" cellspacing="0" cellpadding="3">
                            <tr bgcolor="#f5f5f5"> 
                              <td align="right"><font size="2" face="Arial, Helvetica, sans-serif">Begin Date</font></td>
                              <td><font size="1" face="Arial, Helvetica, sans-serif"><input name="dtPull" type="text" id="dtPull" size="10" maxlength="10" value="<%=request("dtPull")%>">&nbsp;&nbsp;<a href="#" onClick="YY_Calendar('dtPull',250,230,'de','#000000','#f5f5f5','YY_calendar1')"><img src="images/calendar.gif" width="34" height="21" border="0"></a></font></td>
                              <td align="right"><font size="2" face="Arial, Helvetica, sans-serif">End Date</font></td>
                              <td><font size="1" face="Arial, Helvetica, sans-serif"><input name="dtTurn" type="text" id="dtTurn" size="10" maxlength="10" value="<%=request("dtTurn")%>">&nbsp;&nbsp;<a href="#" onClick="YY_Calendar('dtTurn',475,230,'de','#000000','#f5f5f5','YY_calendar1')"><img src="images/calendar.gif" width="34" height="21" border="0"></a></font></td>
                              <td align="right"><font size="2" face="Arial, Helvetica, sans-serif">Customer</font></td>
                              <td align="left"><font size="2" face="Arial, Helvetica, sans-serif"> 
                                <select name="idCustomer" size="1" id="idCustomer">
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
                              <td><font size="1" face="Arial, Helvetica, sans-serif"><input name="Submit" type="submit" onClick="YY_checkform('form1','dtPull','#q','0','Please enter a Begin Date.','dtTurn','#q','0','Please enter an End Date.');return document.MM_returnValue" value="Find"></font></td>
                            </tr>
                          </table>
                         </td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	if errorflag = 0 then
%>
                <tr> 
                  <td>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td height="20" bgcolor="#c0c0c0" colspan="6">
						  <table width="100%" border="0" cellspacing="0" cellpadding="5">
							<tr>
							  <td align="left"><a HREF="javascript:newWindow('excel/exceladded.asp?idCustomer=<%=idCustomer%>&amp;dtPull=<%=request("dtPull")%>&amp;dtTurn=<%=request("dtTurn")%>')"><img SRC="images/exporttoexcel.gif" border="0" WIDTH="120" HEIGHT="19"></a></td>
							</tr>
						  </table>
                        </td>
                      </tr>
                      <tr bgcolor="#6699cc"> 
                        <td height="20"><font color="#ffffff" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Asset #</font></td>
                        <td height="20"><font color="#ffffff" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Serial Number</font></td>
						<td height="20"><font color="#ffffff" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Customer</font></td>
                        <td height="20"><font color="#ffffff" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Item</font></td>
                        <td height="20"><font color="#ffffff" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Date Added</font></td>
                      </tr>
<%
	if rsInventory.EOF then
%>
                      <tr> 
                        <td height="20" colspan="5" align="center"><font size="1" face="Arial, Helvetica, sans-serif">Not Assets to Display.</font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="5"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
                      </tr>
<%
	else
		do until rsInventory.EOF
		if bgswitch = 1 then
			bgswitch = 0
			bgcolor = "#ffffff"
		else
			bgswitch = 1
			bgcolor = "#f5f5f5"
		end if
		'create a counter for billing
		counter = counter+1
%>
                      <tr bgcolor="<%=bgcolor%>"> 
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrAssNum"))%></font></td>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrSerialNum"))%></font></td>
						<td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrCustomer"))%></font></td>
<%
	if rsInventory("chrType") = "C" then
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrItem")) & " - " & trim(rsInventory("chrProcessor"))%><BR>
                        &nbsp;<%=trim(rsInventory("chrMemory")) & " - " & trim(rsInventory("chrODrive")) & " - " & trim(rsInventory("chrHDD"))%></font></td>
<%
	else
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=trim(rsInventory("chrItem"))%></font></td>
<%
	end if
%>
                        <td height="20"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;<%=formatdatetime(rsInventory("dtIn"),2)%></font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="5"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
                      </tr>
<%
		rsInventory.MoveNext
		loop
	end if
	'close the recordset
	rsInventory.Close
	set rsInventory = nothing
%>
                      <tr bgcolor="#6699cc"> 
                        <td height="20" colspan="5" align="left"><font color="#ffffff" size="1" face="Arial, Helvetica, sans-serif">&nbsp;Total Assets Added: <%=counter%></font></td>
                      </tr>
                      <tr bgcolor="#c0c0c0"> 
                        <td height="1" colspan="5"><img src="images/c0c0c0dot.gif" width="1" height="1"></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="20"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
<%
	'from the errorflag
	end if
%>
              </table>
            </td>
            <td><img src="images/ffffffdot.gif" width="15" height="1"></td>
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
	rsCustomers.Close
	set rsCustomers = nothing
	dbConnection.Close
	set dbConnection = nothing
%>