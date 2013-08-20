<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ControlData.inc"-->

<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Claim Key Office Code Search</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<script>
function SelectOption(objSelect, strValue)
{
	var i, iRetVal=-1;

	for (i=0; i < objSelect.length; i ++)
	{
		if (strValue == objSelect(i).value)
		{
			objSelect(i).selected = true;
			return;
		}
	}
}
</script>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

Sub BtnClear_onclick()
	document.all.SearchKOCID.value = ""
	document.all.SearchBID.value = ""
	document.all.SearchAHSID.value = ""
	document.all.SearchKOC.value = ""
	document.all.SearchBNUM.value = ""
End Sub

Sub BtnSearch_onclick()
	'If document.all.SearchCARID.value = "" And document.all.SearchBID.value = "" And _
	 'document.all.SearchAHSID.value = "" And document.all.SearchRuleID.value = "" And _
	 'document.all.SearchRuleText.value = "" And  document.all.SearchLOBCD.value = "" Then
	'	MsgBox "Please enter search criteria", 0, "FNSNetDesigner"
	'Else
		document.all.SpanStatus.innerHTML = "<%=MSG_SEARCH%>"
		FrmSearch.submit
	'End If
End Sub

Sub window_onload
	'document.all.SearchName.focus ' Timing Problem
	document.all.SearchType(0).checked = True
	UpdateStatus("Ready")	
	
<%	If Request.QueryString <> "" Then %>
<%		If CStr(Request.QueryString("SearchType")) = "B" Then	%>
			document.all.SearchType(0).checked = True
<%		ElseIf CStr(Request.QueryString("SearchType")) = "C" Then	%>
			document.all.SearchType(1).checked = True
<%		ElseIf CStr(Request.QueryString("SearchType")) = "E" Then	%>
			document.all.SearchType(2).checked = True
<%		End If 
	End If %>	


	If document.all.SearchKOCID.value <> "" Or document.all.SearchBID.value <> "" Or _
	 document.all.SearchAHSID.value <> "" Or document.all.SearchKOC.value <> "" Or _
	 document.all.SearchBNUM.value <> "" Then
		UpdateStatus("<%=MSG_PROMPT%>")	
	End If
End Sub

Sub PostTo(strURL)
	curKOCID = Parent.frames("WORKAREA").GetKOCID
	temp = Split(curKOCID, "||")
	If UBound(temp) >= 0 Then 
		document.all.KOCID.value = temp(0)
	Else		
		document.all.KOCID.value = ""
	End If	
	FrmSearch.action = "ClaimKOCDetails-f.asp"
	FrmSearch.method = "GET"	
	FrmSearch.target = "_parent"	
	FrmSearch.submit
End Sub
Sub UpdateStatus(inStatus)
	document.all.SpanStatus.innerHTML = inStatus
End Sub

Sub StatusRpt_OnClick
	MsgBox "No other detail status reported.",0,"FNSNetDesigner"		
End Sub
sub enable_exact()
   document.all.SearchType(2).checked  = true 
   document.all.SearchType(0).disabled  = true 
   document.all.SearchType(1).disabled  = true 
end sub

sub enable_begin()
   document.all.SearchType(0).checked  = true 
   document.all.SearchType(0).disabled  = false 
   document.all.SearchType(1).disabled  = false 
   
end sub
<!--#include file="..\lib\Help.asp"-->
</script>

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Claim Key Office Code Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label" >
<tr>
<td VALIGN="CENTER" WIDTH="5" >
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER"  ALT="View Status Report">
</td>
<td width="485">
:<SPAN VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL>Ready</SPAN>
</td>
</tr>
</table>

<form Name="FrmSearch" METHOD="GET" ACTION="ClaimKOCSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="KOCID" value="<%=Request.QueryString("KOCID")%>">
<table width="100%" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
 <table>
	<tr></tr>
	<tr></tr>
	<tr>
	<td CLASS="LABEL">A.H. Step ID:<br><input CLASS="LABEL"  size="12" tabindex=1 TYPE="TEXT" NAME="SearchAHSID" VALUE="<%=Request.QueryString("SearchAHSID")%>" onfocus="enable_exact()" onBlur="enable_begin()" ></td>
	<td CLASS="LABEL">Branch ID:<br><input CLASS="LABEL" size="12" tabindex=2 TYPE="TEXT" NAME="SearchBID" VALUE="<%=Request.QueryString("SearchBID")%>"></td>
	<td CLASS="LABEL">Key Office Code:<br><input size="12" CLASS="LABEL" tabindex=4 TYPE="TEXT" NAME="SearchBNUM" VALUE="<%=Request.QueryString("SearchBNUM")%>"></td>
	</tr>				
	<tr>
	<td CLASS="LABEL" colspan=2>Sub Key Office Code:<br><input CLASS="LABEL" tabindex=5  size="12" TYPE="TEXT" NAME="SearchKOC" VALUE="<%=Request.QueryString("SearchKOC")%>"></td>
	<td CLASS="LABEL">Claim Key Office Code Type ID:<br><input CLASS="LABEL" tabindex=6  size="12" TYPE="TEXT" NAME="SearchKOCID" VALUE="<%=Request.QueryString("SearchKOCID")%>"></td>
	</tr>
 </table>
</td>			
<td VALIGN="TOP" rowspan="3">
 <table>
 <tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex=10 NAME="BtnSearch" TYPE="BUTTON" ACCESSKEY="H">Searc<u>h</u></button></td></tr>
 <tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex=11 NAME="BtnClear" ACCESSKEY="L">C<u>l</u>ear</button></td></tr>
 </table>
</td>	
</tr>
</table>

<table topmargin=0 bottommargin=0>
<tr>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" tabindex=7 NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" tabindex=8 NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" tabindex=9 NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
</tr>
</table>

</form>
</body>
</html>
