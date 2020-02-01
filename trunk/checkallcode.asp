<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language = "Javascript">
<!-- 
function checkAllBoxes(e)
{

  for (i = 0; i < document.forms[0].length; i++)
  {
     if (document.forms[0].elements[i].type == "checkbox")
     {
        if (e.checked)
        {
	       document.forms[0].elements[i].checked = true;
        }
	    else
	    {
	       document.forms[0].elements[i].checked = false;
        }
     }
  }
}

// -->
</script>
</HEAD>

<BODY>

<form name="frmSample" method="post" action="#">
  <div align="left"><input type="checkbox" name="chkall" onClick="checkAllBoxes(this,'frmSample')">
  &nbsp;Check to Select/Deselect all</div><br><br>
<!--your dynamic table contents comes here (change the name chkSample everywhere if you want to) -->
<input type="checkbox" name="1" value="A"> A <br>
<input type="checkbox" name="2" value="B"> B <br>
<input type="checkbox" name="3" value="C"> C <br>
</form>

</BODY>
</HTML>
