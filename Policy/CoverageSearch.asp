<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Coverage Search</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub BtnClear_onclick()
	document.all.SearchCOVID.value = ""
	document.all.SearchPOLICY_ID.value = ""
	document.all.SearchVEHICLE_ID.value = ""
	document.all.SearchEFFECTIVE_DATE.value = ""
	document.all.SearchEXPIRATION_DATE.value = ""
End Sub

Sub BtnSearch_onclick()
	'If document.all.SearchCOVID.value = "" And document.all.SearchPOLICY_ID.value = "" And _
	'document.all.SearchVEHICLE_ID.value = "" And _
	'document.all.SearchEFFECTIVE_DATE.value = ""  And document.all.SearchEXPIRATION_DATE.value = "" Then
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
If document.all.SearchCOVID.value <> "" And document.all.SearchPOLICY_ID.value <> "" And _
	document.all.SearchVEHICLE_ID.value <> "" And _
	document.all.SearchEFFECTIVE_DATE.value <> ""  And document.all.SearchEXPIRATION_DATE.value <> "" Then
		UpdateStatus("<%=MSG_PROMPT%>")	
	End If

End Sub

Sub PostTo(strURL)
	curCOVID = Parent.frames("WORKAREA").GetCOVID
	temp = Split(curCOVID, "||")
	If UBound(temp) >= 0 Then 
		document.all.COVID.value = temp(0)
	Else		
		document.all.COVID.value = ""
	End If
	FrmSearch.action = "CoverageDetails-f.asp"
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
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Coverage Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<table style="{position:absolute;top:20;}" class="Label">
<tr>
<td VALIGN="CENTER" WIDTH="5">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER" ALT="View Status Report">
</td>
<td width="485">
:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>
<form Name="FrmSearch" METHOD="GET" ACTION="CoverageSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="COVID" value="<%=Request.QueryString("COVID")%>">
<table width="100%" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
	<table CLASS="LABEL" style="width:300" align="left">
	<tr>
	<td CLASS="LABEL"><nobr>Coverage ID:<br><input CLASS="LABEL" tabindex="1" TYPE="TEXT" NAME="SearchCOVID" size="14" VALUE="<%=Request.QueryString("SearchCOVID")%>"></td>
	<td CLASS="LABEL"><nobr>Policy ID:<br><input CLASS="LABEL" tabindex="2" TYPE="TEXT" NAME="SearchPOLICY_ID" size="14" VALUE="<%=Request.QueryString("SearchPOLICY_ID")%>"></td>
	<td CLASS="LABEL"><nobr>Vehicle ID:<br><input CLASS="LABEL" tabindex="3" TYPE="TEXT" NAME="SearchVEHICLE_ID" size="14" VALUE="<%=Request.QueryString("SearchVEHICLE_ID")%>"></td>
	</tr>
	<tr>
	<td CLASS="LABEL"><nobr>Effective Date:<br><input size="14" MAXLENGTH="10" tabindex="5" CLASS="LABEL" TYPE="TEXT" NAME="SearchEFFECTIVE_DATE" VALUE="<%=Request.QueryString("SearchEFFECTIVE_DATE")%>"></td>
	<td CLASS="LABEL"><nobr>Expiration Date:<br><input CLASS="LABEL" size="14" MAXLENGTH="10" TYPE="TEXT" tabindex="6" NAME="SearchEXPIRATION_DATE" VALUE="<%=Request.QueryString("SearchEXPIRATION_DATE")%>"></td>
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
