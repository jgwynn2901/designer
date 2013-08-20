<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Output Page Search</title>
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
	document.all.SearchOPID.value = ""
	document.all.SearchName.value = ""
	document.all.SearchODID.value = ""
	document.all.SearchBMP.value = ""
End Sub

Sub BtnSearch_onclick()
	'If document.all.SearchOPID.value = "" And document.all.SearchName.value = "" And _
	'document.all.SearchODID.value = "" Then
	'		MsgBox "Please enter search criteria", 0, "FNSNetDesigner"
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

		If CStr(Request.QueryString("SearchInputType")) <> "" Then	%>
			SelectOption document.all.SearchInputType,"<%=CStr(Request.QueryString("SearchInputType"))%>"
<%		End If 

	End If %>	

	If document.all.SearchOPID.value <> "" Or document.all.SearchName.value <> "" Or _
	document.all.SearchODID.value <> "" or document.all.SearchBMP.value <> "" Then
		UpdateStatus("<%=MSG_PROMPT%>")	
	End If
End Sub

Sub PostTo(strURL)
	curOPID = Parent.frames("WORKAREA").GetOPID
	temp = Split(curOPID, "||")
	If UBound(temp) >= 0 Then 
		document.all.OPID.value = temp(0)
	Else		
		document.all.OPID.value = ""
	End If
	FrmSearch.action = "OutputPageDetails-f.asp"
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
<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Output Page Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label">
<tr>
<td VALIGN="CENTER" WIDTH="5">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER" Title="View Status Report">
</td>
<td width="485">
:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>
<form Name="FrmSearch" METHOD="GET" ACTION="OutputPageSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="OPID" value="<%=Request.QueryString("OPID")%>">
<table width="100%" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
	<table CLASS="LABEL" style="width:300" align="left">
	<tr>
	<tr>
	<tr>
	<td CLASS="LABEL">Name:<br><input CLASS="LABEL" tabindex="1" TYPE="TEXT" NAME="SearchName" size="45" VALUE="<%=Request.QueryString("SearchName")%>"></td>
	<td CLASS="LABEL"><nobr>Output Def. ID:<br><input CLASS="LABEL" tabindex="2" TYPE="TEXT" NAME="SearchODID" size="10" VALUE="<%=Request.QueryString("SearchODID")%>"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">BMP:<br><input CLASS="LABEL" SIZE="45" TYPE="TEXT" tabindex="3" NAME="SearchBMP" VALUE="<%=Request.QueryString("SearchBMP")%>"></td>	
	<td CLASS="LABEL"><nobr>Output Page ID:<br><input CLASS="LABEL" SIZE="10" TYPE="TEXT" tabindex="4" NAME="SearchOPID" VALUE="<%=Request.QueryString("SearchOPID")%>"></td>
	</tr>
	</table>
</td>
<td VALIGN="TOP" rowspan="3">
	<table>
	<tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex="10" NAME="BtnSearch" TYPE="BUTTON" ACCESSKEY="H">Searc<u>h</u></button></td></tr>
	<tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex="11" NAME="BtnClear" ACCESSKEY="L">C<u>l</u>ear</button></td></tr>
	</table>
</td>	
</tr>
<tr>
<td>
	<table>
	<tr>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="7" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="8" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="9" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
	</tr>
	</table>
</td>
</tr>
</table>


</form>
</body>
</html>
