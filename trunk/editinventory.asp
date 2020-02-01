<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "logoff.asp"
	end if
	
	'turn on inventory button
	buttonswitch = 3
%>
<!-- #include file="includes/openconn.asp" -->
<%	
	'Get a list of the Warehouses
	set rsWarehouses = server.CreateObject("adodb.recordset")
	sql = "execute ListWarehouses"
	set rsWarehouses = dbConnection.Execute(sql)
	
	'Get a list of the Customers
	set rsCustomers = server.CreateObject("adodb.recordset")
	sql = "execute ListCustomerNamesandIDs"
	set rsCustomers = dbConnection.Execute(sql)
	
	'Get a list of the Descriptions
	set rsDescriptions = server.CreateObject("adodb.recordset")
	sql = "execute ListDescriptionDropDown"
	set rsDescriptions = dbConnection.Execute(sql)
	
	'Find the inventory item
	set rsInventory = server.CreateObject("adodb.recordset")
	sql = "execute FindInventorybyID " & request("idInventory")
	set rsInventory = dbConnection.Execute(sql)
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
<form name="form1" method="post" action="updateinventory.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
    <!-- #include file="includes/top.htm" -->
    <!-- #Middle top part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
		<!-- #include file="includes/inventory-nav.htm" -->
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
                        <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Edit Inventory</strong></font></td>
                        <td width="50%">&nbsp;</td>
                      </tr>
                      <tr bgcolor="#f5f5f5"> 
                        <td colspan="2"><strong><font color="#FF0000" size="2" face="Arial, Helvetica, sans-serif">Warning you are about to change the asset number. You will be responsible for priting new label in the system and placing them on the boxes and assets themselves. </font></strong></td>
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
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Warehouse<br>
                          <select name="idWarehouse" size="1" id="idWarehouse">
<%
	if not rsWarehouses.EOF then
		do until rsWarehouses.EOF
%>
							<option value="<%=rsWarehouses("idWarehouse")%>" <%if rsInventory("idWarehouse") = rsWarehouses("idWarehouse") then%>selected<%end if%>><%=trim(rsWarehouses("chrWarehouse"))%></option>
<%
		rsWarehouses.MoveNext
		loop
	end if
%>
						  </select></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Customer<br>
                          <select name="idCustomer" size="1" id="idCustomer">
<%
	if not rsCustomers.EOF then
		do until rsCustomers.EOF
%>
							<option value="<%=rsCustomers("idCustomer")%>" <%if rsInventory("idCustomer") = rsCustomers("idCustomer") then%>selected<%end if%>><%=trim(rsCustomers("chrCustomer"))%></option>
<%
		rsCustomers.MoveNext
		loop
	end if
%>
						  </select></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Asset Owner<br>
                          <select name="idOwner" size="1" id="idOwner">
							<option value="1" <%if rsInventory("idOwner") = 1 then%>selected<%end if%>>Pool</option>
							<option value="2" <%if rsInventory("idOwner") = 2 then%>selected<%end if%>>Product Marketing</option>
							<option value="3" <%if rsInventory("idOwner") = 3 then%>selected<%end if%>>Third Party</option>
							<option value="4" <%if rsInventory("idOwner") = 4 then%>selected<%end if%>>Education Marketing</option>
						  </select></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Asset Description<br>
                          <select name="idDescription" size="1" id="idDescription">
<%
	if not rsDescriptions.EOF then
		do until rsDescriptions.EOF
			if rsDescriptions("chrType") = "C" then
%>
							<option value="<%=rsDescriptions("idDescription")%>" <%if rsInventory("idDescription") = rsDescriptions("idDescription") then%>selected<%end if%>><%=trim(rsDescriptions("chrItem")) & " - " & trim(rsDescriptions("chrProcessor")) & " - " & trim(rsDescriptions("chrMemory")) & " - " & trim(rsDescriptions("chrODrive"))%></option>
<%
			else
%>
							<option value="<%=rsDescriptions("idDescription")%>" <%if rsInventory("idDescription") = rsDescriptions("idDescription") then%>selected<%end if%>><%=trim(rsDescriptions("chrItem"))%></option>
<%
			end if
		rsDescriptions.MoveNext
		loop
	end if
%>
						  </select></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Asset Number</font><br>
                          <font size="2" face="Arial, Helvetica, sans-serif" color="#FF0000"><strong>Previous Number: <%=trim(rsInventory("chrAssNum"))%></strong></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Serial Number<br>
                          <input name="chrSerialNum" type="text" id="chrSerialNum" size="27" maxlength="25" value="<%=trim(rsInventory("chrSerialNum"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Notes<br>
                          <textarea name="txtNotes" cols="50" rows="5" wrap="VIRTUAL" id="txtNotes"><%=trim(rsInventory("txtNotes"))%></textarea></font></td>
                      </tr>
                      <tr> 
                        <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                      </tr>
                      <tr> 
                        <td><input name="Submit" type="submit" onClick="YY_checkform('form1','chrAssNum','#q','0','Please enter an Asset Number.','chrSerialNum','#q','0','Please enter a Serial Number.');return document.MM_returnValue" value="Update Inventory">
                        <input type="hidden" name="idInventory" value="<%=request("idInventory")%>"></td>
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
		document.form1.idWarehouse.focus()
  </script>
  </form>
</body>
</html>
<%
	rsWarehouses.Close
	set rsWarehouses = nothing
	rsCustomers.Close
	set rsCustomers = nothing
	rsDescriptions.Close
	set rsDescriptions = nothing
	rsInventory.Close
	set rsInventory = nothing
	dbConnection.Close
	set dbConnection = nothing
%>