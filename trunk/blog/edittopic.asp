<%@ Language=VBScript %>
<%
	'check to see if the user is connected to the server
	if session("idUser") = "" then
		Response.Redirect "../logoff.asp"
	end if
%>
<!-- #include file="../includes/openconn.asp" -->
<%
	'get a list of orders that will return in 7 days
	set rsTopics = server.CreateObject("adodb.recordset")
	sql = "execute FindTopicbyID " & request("idBlog")
	set rsTopics = dbConnection.Execute(sql)
%>
<html>
<head>
<title>techIT Solutions Asset Management Blog</title>
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

<body>
<form name="form1" method="post" action="updatetopic.asp">
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td><table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr> 
            <td width="50%"><font size="3" face="Arial, Helvetica, sans-serif"><strong>Asset Management BLOG</strong></font></td>
            <td width="50%" align="right" valign="bottom">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr>
            <td bgcolor="#6699cc"><font color="#ffffff" size="3" face="Arial, Helvetica, sans-serif"><strong>Edit Topic</strong></font></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="20">&nbsp;</td>
    </tr>
    <tr> 
      <td><table width="100%" border="0" cellspacing="0" cellpadding="3">
          <tr> 
            <td><font size="1" face="Arial, Helvetica, sans-serif">Status<br>
              <select name="idStatus" size="1" id="idStatus">
                <option value="0" <%if rsTopics("idStatus") = 0 then%>selected<%end if%>>Please Choose</option>
                <option value="1" <%if rsTopics("idStatus") = 1 then%>selected<%end if%>>New</option>
                <option value="2" <%if rsTopics("idStatus") = 2 then%>selected<%end if%>>Completed</option>
              </select>
              </font></td>
          </tr>
		  <tr> 
            <td><font size="1" face="Arial, Helvetica, sans-serif">Type<br>
              <select name="idType" size="1" id="idType">
                <option value="0" <%if rsTopics("idType") = 0 then%>selected<%end if%>>Please Choose</option>
                <option value="1" <%if rsTopics("idType") = 1 then%>selected<%end if%>>Problem</option>
                <option value="2" <%if rsTopics("idType") = 2 then%>selected<%end if%>>Upgrade</option>
              </select>
              </font></td>
          </tr>
          <tr> 
            <td><font size="1" face="Arial, Helvetica, sans-serif">Priority<br>
              <select name="idPriority" size="1" id="idPriority">
                <option value="0" <%if rsTopics("idPriority") = 0 then%>selected<%end if%>>Please choose</option>
                <option value="1" <%if rsTopics("idPriority") = 1 then%>selected<%end if%>>High</option>
                <option value="2" <%if rsTopics("idPriority") = 2 then%>selected<%end if%>>Medium</option>
                <option value="3" <%if rsTopics("idPriority") = 3 then%>selected<%end if%>>Low</option>
              </select>
              </font></td>
          </tr>
          <tr> 
            <td><font size="1" face="Arial, Helvetica, sans-serif">Title<br>
              <input name="chrTitle" type="text" id="chrTitle" size="50" maxlength="75" value="<%=trim(rsTopics("chrTitle"))%>"></font></td>
          </tr>
          <tr> 
            <td><font size="1" face="Arial, Helvetica, sans-serif">Description<br>
              <textarea name="txtMessage" cols="50" rows="15" wrap="VIRTUAL" id="txtMessage"><%=trim(rsTopics("txtMessage"))%></textarea></font></td>
          </tr>
          <tr> 
            <td height="20">&nbsp;</td>
          </tr>
          <tr> 
            <td><input name="Submit" type="submit" onClick="YY_checkform('form1','chrTitle','#q','0','Please type a topic.','idType','#q','1','Please choose a Type.','idPriority','#q','1','Please choose a Priority.','idStatus','#q','1','Please choose a Status.');return document.MM_returnValue" value="Update Topic">
			<input type="hidden" name="idBlog" value="<%=request("idBlog")%>"></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="20">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
<%
	rsTopics.close
	set rsTopics = nothing
	dbConnection.Close
	set dbConnection = nothing
%>