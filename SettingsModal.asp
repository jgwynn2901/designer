<!--#include file="lib\common.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>

<html>
<head>
<link rel="stylesheet" type="text/css" href="FNSDESIGN.css">
<title>Settings</title>
<style TYPE="text/css">
HTML {width: 300pt; height:180pt}
</style>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
function BtnClear_onclick() 
{
	if (document.frames("DocFrame").document.frames("TOP").HasSelection() == true)
	{
		lret = confirm("Are you sure you want to clear the selected settings?");
		if (lret == true) 
		{
		document.frames("DocFrame").document.frames("TOP").document.all.ACTION.value="CLEAR";
		document.frames("DocFrame").document.frames("TOP").document.all.FrmSettings.submit();
		}
	}		
	else
		alert("Select a setting to clear.");
		
}


function BtnClose_onclick() 
{
	window.close();
}

function BtnSave_onclick()
{
var PageObj = document.frames("DocFrame").document.frames("TOP").document.all

	if (document.frames("DocFrame").document.frames("TOP").HasSelection() == true)
	{
	
		if ((document.frames("DocFrame").document.frames("TOP").ValidNumber(PageObj.TREELEVELS.value) == true) && (document.frames("DocFrame").document.frames("TOP").ValidNumber(PageObj.TREECOUNT.value) == true) && (document.frames("DocFrame").document.frames("TOP").ValidNumber(PageObj.SAVEMAXRECORDS.value) == true))
		{
			document.frames("DocFrame").document.frames("TOP").document.all.ACTION.value = "SAVE";
			document.frames("DocFrame").document.frames("TOP").document.all.FrmSettings.submit();
		}
		else
			alert ("Invalid settings, settings must be numeric, not null, and greater than one.")
	}
	else
		alert("Select a setting to save.");
}
</SCRIPT>


</head>
<body LEFTMARGIN="0" TOPMARGIN="0" bgcolor='<%= BODYBGCOLOR %>' bottommargin=0 rightmargin=0>
<iframe FRAMEBORDER="0" ID="DocFrame" WIDTH="100%" HEIGHT="80%" SRC="Settings-f.asp"></iframe>
<BR>
<TABLE>
<TR>
<td CLASS="LABEL"><button style="width:100" title="Save selected settings" CLASS="StdButton" NAME="BtnSave" ACCESSKEY="A" LANGUAGE=javascript onclick="return BtnSave_onclick()"><u>S</u>ave Settings</button></td>
<td CLASS="LABEL"><button style="width:100" title="Clear selected settings" CLASS="StdButton" NAME="BtnClear" ACCESSKEY="F" LANGUAGE=javascript onclick="return BtnClear_onclick()">C<U>l</U>ear Settings</button></td>
<td CLASS="LABEL"><button style="width:100" CLASS="StdButton" NAME="BtnClose" ACCESSKEY="O"  LANGUAGE=javascript onclick="return BtnClose_onclick()">Cl<U>o</U>se</button></td>
</TR>
</table>
</body>
</html>