<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "../logoff.asp"
	end if
%>
<!-- #include file="../includes/openconn.asp" -->
<%	
	'Find the Cart information
	set rsCart = server.CreateObject("adodb.recordset")
	sql = "execute FindCartbyID " & request("idCart")
	set rsCart = dbConnection.Execute(sql)
	
	'check to see if the Linked cart has been filled out
	if rsCart("idLinkedCart") > 0 then
		session("idLinkedCart") = rsCart("idLinkedCart")
		session("idCart") = rsCart("idCart")
		Response.Redirect "default.asp"
	else
		'Find the Cart information
		set rsCarts = server.CreateObject("adodb.recordset")
		sql = "execute ListShow2ShowCartsAllAccess " & request("idCart") & "," & session("idUser") & ",'" & session("idAccess") & "'"
		set rsCarts = dbConnection.Execute(sql)
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

function YY_checkform() { //v4.65
//copyright (c)1998,2002 Yaromat.com
  var args = YY_checkform.arguments; var myDot=true; var myV=''; var myErr='';var addErr=false;var myReq;
  for (var i=1; i<args.length;i=i+4){
    if (args[i+1].charAt(0)=='#'){myReq=true; args[i+1]=args[i+1].substring(1);}else{myReq=false}
    var myObj = MM_findObj(args[i].replace(/\[\d+\]/ig,""));
    myV=myObj.value;
    if (myObj.type=='text'||myObj.type=='password'||myObj.type=='hidden'){
      if (myReq&&myObj.value.length==0){addErr=true}
      if ((myV.length>0)&&(args[i+2]==1)){ //fromto
        var myMa=args[i+1].split('_');if(isNaN(parseInt(myV))||myV<myMa[0]/1||myV > myMa[1]/1){addErr=true}
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
<form name="form1" method="post" action="updateshow2show.asp">
  <table width="700" border="0" align="center" cellpadding="0" cellspacing="0">
    <!-- #include file="includes/top.htm" -->
    <tr> 
      <td bgcolor="#6699cc"><table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr> 
            <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr> 
                  <td><font size="3" face="Arial, Helvetica, sans-serif"><strong>Cart Name: <%=trim(rsCart("chrCart"))%></strong></font></td>
                </tr>
                <tr> 
                  <td>
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Ship to:</font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Ships: <%=formatdatetime(rsCart("dtShip"),1)%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrAddress"))%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Arrives: <%=formatdatetime(rsCart("dtArrival"),1)%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrAddress2"))%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Returns: <%=formatdatetime(rsCart("dtReturn"),1)%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrAddress3"))%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Onsite Contact: <%=trim(rsCart("chrOSPerson"))%></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif"><%=trim(rsCart("chrCity"))%>, <%=trim(rsCart("chrState"))%> <%=trim(rsCart("chrZip"))%></font></td>
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Onsite Number: <%=trim(rsCart("chrOSPhone"))%></font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
                <tr> 
                  <td height="25"><img src="../images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><font size="2" face="Arial, Helvetica, sans-serif">Please choose a Cart below that you want to move assets from. You will only be able to choose carts that end before 
                  this cart Arrives. Once you have linked these carts together you will only be able to pull assets from this cart. You will have to create another Cart for Adding assets from the warehouse
                  or another Cart.</font></td>
                </tr>
                <tr> 
                  <td><select name="idLinkedCart" size="1" id="idLinkedCart">
                      <option value="0" selected>Please choose a Cart</option>
<%
	if not rsCarts.EOF then
		do until rsCarts.EOF
%>
                      <option value="<%=rsCarts("idCart")%>"><%=formatdatetime(rsCarts("dtDeparture"),2) & " - " & trim(rsCarts("chrCart"))%></option>
<%
		rsCarts.MoveNext
		loop
	end if
%>
                    </select></td>
                </tr>
                <tr> 
                  <td height="15"><img src="../images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><input name="Submit" type="submit" onClick="YY_checkform('form1','idLinkedCart','#q','1','Please choose a Cart to pull assets from.');return document.MM_returnValue" value="Start Adding Assets">
                  <input type="hidden" name="idCart" value="<%=request("idCart")%>"></td>
                </tr>
                <tr> 
                  <td height="25"><img src="../images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
    <!-- #include file="includes/bottom.htm" -->
  </table>
</form>
</body>
<%
		rsCarts.Close
		set rsCarts = nothing
	end if
	'close the connections
	rsCart.Close
	set rsCart = nothing
	dbConnection.Close
	set dbConnection = nothing
%>