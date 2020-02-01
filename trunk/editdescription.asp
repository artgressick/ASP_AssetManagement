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
	'Get a list of the Categories
	set rsCategories = server.CreateObject("adodb.recordset")
	sql = "execute ListCategories"
	set rsCategories = dbConnection.Execute(sql)
	
	'Get the description
	set rsDescription = server.CreateObject("adodb.recordset")
	sql = "execute FindDescriptionbyID " & request("idDescription")
	set rsDescription = dbConnection.Execute(sql)
	
	'Get a list of the Pictures
	set rsPictures = server.CreateObject("adodb.recordset")
	sql = "execute ListPictures"
	set rsPictures = dbConnection.Execute(sql)
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
<form name="form1" method="post" action="updatedescription.asp">
  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
    <!-- #Begin top part -->
	<!-- #include file="includes/top.htm" -->
	<!-- #Begin Middle part -->
    <tr>
      <td height="100%" valign="top" background="images/175leftbar.gif" bgcolor="#f5f5f5">
	  	<!-- #include file="includes/inventory-nav.htm" -->
      </td>
      <td width="100%" height="100%" valign="top">
	  	<table width="625" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15"><img src="images/ffffffdot.gif" width="15" height="1"></td>
            <td width="610"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Edit Description</strong></font></td>
                        <td width="50%">&nbsp;</td>
                      </tr>
                      <tr bgcolor="#f5f5f5"> 
                        <td colspan="2"><font size="1" face="Arial, Helvetica, sans-serif">To edit a description please enter all of the information below and press Update Description button.</font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="15"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Item Number<br>
                          <input name="chrItemNo" type="text" id="chrItemNo" size="17" maxlength="15" value="<%=trim(rsDescription("chrItemNo"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Item Name (example: PowerBook 17&quot; or IBM T - Series Laptop)<br>
                          <input name="chrItem" type="text" id="chrItem" size="50" maxlength="100" value="<%=trim(rsDescription("chrItem"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Category <br>
                          <select name="idCategory" size="1" id="idCategory">
<%
	if not rsCategories.EOF then
		do until rsCategories.EOF
%>
                            <option value="<%=rsCategories("idCategory")%>" <%if rsCategories("idCategory") = rsDescription("idCategory") then%>selected<%end if%>><%=trim(rsCategories("chrCategory"))%></option>
<%
		rsCategories.MoveNext
		loop
	end if
%>
                          </select>
                          </font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Item Image<br>
                          <select name="idPicture" size="1" id="idPicture">
<%
	if not rsPictures.EOF then
		do until rsPictures.EOF
%>
                            <option value="<%=rsPictures("idPicture")%>" <%if rsPictures("idPicture") = rsDescription("idPicture") then%>selected<%end if%>><%=trim(rsPictures("chrName"))%></option>
<%
		rsPictures.MoveNext
		loop
	end if
%>
                          </select>
                          </font></td>
                      </tr>
					  <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Cost Code (techIT Internal Code)<br>
                          <select name="idCost" size="1" id="idCost">
                            <option value="0" <%if rsDescription("idCost") = 0 then%>selected<%end if%>>0</option>
							<option value="1" <%if rsDescription("idCost") = 1 then%>selected<%end if%>>1</option>
							<option value="2" <%if rsDescription("idCost") = 2 then%>selected<%end if%>>2</option>
							<option value="3" <%if rsDescription("idCost") = 3 then%>selected<%end if%>>3</option>
							<option value="4" <%if rsDescription("idCost") = 4 then%>selected<%end if%>>4</option>
							<option value="5" <%if rsDescription("idCost") = 5 then%>selected<%end if%>>5</option>
							<option value="6" <%if rsDescription("idCost") = 6 then%>selected<%end if%>>6</option>
							<option value="7" <%if rsDescription("idCost") = 7 then%>selected<%end if%>>7</option>
							<option value="8" <%if rsDescription("idCost") = 8 then%>selected<%end if%>>8</option>
							<option value="9" <%if rsDescription("idCost") = 9 then%>selected<%end if%>>9</option>
                          </select>
                          </font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="25"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                </tr>
                <tr> 
                  <td bgcolor="#f5f5f5">
					<table width="100%" border="0" cellspacing="0" cellpadding="3">
                      <tr> 
                        <td><font size="2" face="Arial, Helvetica, sans-serif"><strong>Please only enter information for computers. Leave blank for all other items.</strong></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Processor/CPU (example: 3.06 GHz Intel Pentium 4 or 1.42 GHz Dual Processor Power PC G4)<br>
                          <input name="chrProcessor" type="text" id="chrProcessor" size="45" maxlength="35" value="<%=trim(rsDescription("chrProcessor"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Memory (example: 2 GB DDR 333 SDRAM or 1 GB of RDRAM 800MHz)<br>
                          <input name="chrMemory" type="text" id="chrMemory" size="30" maxlength="25" value="<%=trim(rsDescription("chrMemory"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Storage/Hard Drive (example: 80GB Ultra ATA/100 or 36GB Ultra160 SCSI)<br>
                          <input name="chrHDD" type="text" id="chrHDD" size="30" maxlength="30" value="<%=trim(rsDescription("chrHDD"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Optical Drive (example: SuperDrive (4x) or DVD/CD-RW Combo Drive)<br>
                          <input name="chrODrive" type="text" id="chrODrive" size="30" maxlength="30" value="<%=trim(rsDescription("chrODrive"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Removeable Storage/Second Optical Drive (example: Iomega ZIP 250MB or DVD-RAM)<br>
                          <input name="chrRStorage" type="text" id="chrRStorage" size="30" maxlength="30" value="<%=trim(rsDescription("chrRStorage"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">SCSI/Fibre Channel Device (example: Adaptec Ultra160 PCI or Fibre Channel PCI)<br>
                          <input name="chrSCSI" type="text" id="chrSCSI" size="30" maxlength="25" value="<%=trim(rsDescription("chrSCSI"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Graphics Card (example: ATI Radeon 9700 Pro w/128MB DDR or nVidia Titanium w/128MB DDR)<br>
                          <input name="chrGraphics" type="text" id="chrGraphics" size="30" maxlength="25" value="<%=trim(rsDescription("chrGraphics"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Wireless Card (example: Airport Extreme, Airport 802.11b or Intel Pro/Wireless 802.11b)<br>
                          <input name="chrWireless" type="text" id="chrWireless" size="30" maxlength="25" value="<%=trim(rsDescription("chrWireless"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Bluetooth (example: Bluetooth Module or Linksys Bluetooth Module)<br>
                          <input name="chrBluetooth" type="text" id="chrBluetooth" size="30" maxlength="25" value="<%=trim(rsDescription("chrBluetooth"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Modem (example: 56k Internal Modem or 56k PCI Modem)<br>
                          <input name="chrModem" type="text" id="chrModem" size="30" maxlength="25" value="<%=trim(rsDescription("chrModem"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">USB (example: 4 USB 1.0 or 2 USB 2.0) make sure to put the version of USB!!<br>
                          <input name="chrUSB" type="text" id="chrUSB" size="30" maxlength="25" value="<%=trim(rsDescription("chrUSB"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">FireWire (example: 2 FireWire 800 or 2 FireWire 400) make sure to put the version of FireWire!!<br>
                          <input name="chrFireWire" type="text" id="chrFireWire" size="30" maxlength="25" value="<%=trim(rsDescription("chrFireWire"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Ethernet (example: 10/100/1000 Ethernet LAN)<br>
                          <input name="chrEthernet" type="text" id="chrEthernet" size="30" maxlength="30" value="<%=trim(rsDescription("chrEthernet"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td><font size="1" face="Arial, Helvetica, sans-serif">Operating System (example: Windows XP Professional, Mac OS X Server or Mac OS 9/10 Dual Boot)<br>
                          <input name="chrOS" type="text" id="chrOS" size="35" maxlength="35" value="<%=trim(rsDescription("chrOS"))%>"></font></td>
                      </tr>
                      <tr> 
                        <td height="20" bgcolor="#ffffff"><img src="images/ffffffdot.gif" width="1" height="1"></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#ffffff"><input name="Submit" type="submit" id="Submit" onClick="YY_checkform('form1','chrItemNo','#q','0','Please enter the Item Number.','chrItem','#q','0','Please enter the Item Name.');return document.MM_returnValue" value="Update Description">
                        <input type="hidden" name="idDescription" value="<%=request("idDescription")%>"></td>
                      </tr>
                    </table></td>
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
		document.form1.chrItemNo.focus()
  </script>
  </form>
</body>
</html>
<%
	rsCategories.Close
	set rsCategories = nothing
	rsPictures.Close
	set rsPictures = nothing
	rsDescription.Close
	set rsDescription = nothing
	dbConnection.Close
	set dbConnection = nothing
%>