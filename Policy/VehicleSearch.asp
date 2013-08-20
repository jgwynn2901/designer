<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Vehicle Search</title>
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
	document.all.SearchVID.value = ""
	document.all.SearchPOLICY_ID.value = ""
	document.all.SearchVIN.value = ""
	document.all.SearchYEAR.value = ""
	document.all.SearchMAKE.value = ""
	document.all.SearchMODEL.value = ""
	document.all.SearchLICENSE_PLATE.value = ""
	document.all.SearchLICENSE_PLATE_STATE.value = ""
	document.all.SearchREGISTRATION_STATE.value = ""
	document.all.SearchCOLOR.value = ""
End Sub

Sub BtnSearch_onclick()
	'If document.all.SearchVID.value = "" And document.all.SearchPOLICY_ID.value = "" And _
	'document.all.SearchVIN.value = "" And document.all.SearchYEAR.value = "" And _
	'document.all.SearchMAKE.value = "" And document.all.SearchMODEL.value = "" And _
	'document.all.SearchLICENSE_PLATE_STATE.value = "" And document.all.SearchREGISTRATION_STATE.value = "" And _
	'document.all.SearchLICENSE_PLATE.value = ""  And document.all.SearchCOLOR.value = "" Then
	'		MsgBox "Please enter search criteria", 0, "FNSNetDesigner"
	'Else
		document.all.SpanStatus.innerHTML = "<%=MSG_SEARCH%>"
		FrmSearch.submit
	'End If
End Sub

Sub window_onload
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

If document.all.SearchVID.value <> "" And document.all.SearchPOLICY_ID.value <> "" And _
	document.all.SearchVIN.value <> "" And document.all.SearchYEAR.value <> "" And _
	document.all.SearchMAKE.value <> "" And document.all.SearchMODEL.value <> "" And _
	document.all.SearchLICENSE_PLATE_STATE.value <> "" And document.all.SearchREGISTRATION_STATE.value <> "" And _
	document.all.SearchLICENSE_PLATE.value <> ""  And document.all.SearchCOLOR.value <> "" Then

		UpdateStatus("<%=MSG_PROMPT%>")	
	End If
<% If Request.QueryString("SearchLICENSE_PLATE_STATE") <> "" Then %>
document.all.SearchLICENSE_PLATE_STATE.value = "<%= Request.QueryString("SearchLICENSE_PLATE_STATE") %>"
<% End If %>
<% If Request.QueryString("SearchREGISTRATION_STATE") <> "" Then %>
document.all.SearchREGISTRATION_STATE.value = "<%= Request.QueryString("SearchREGISTRATION_STATE") %>"
<% End If %>
End Sub

Sub PostTo(strURL)

	curVID = Parent.frames("WORKAREA").GetVID
	temp = Split(curVID, "||")
	If UBound(temp) >= 0 Then 
		document.all.VID.value = temp(0)
	Else		
		document.all.VID.value = ""
	End If
	FrmSearch.action = "VehicleDetails-f.asp"
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
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Vehicle Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
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
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER" ALT="View Status Report">
</td>
<td width="485">
:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>

<form Name="FrmSearch" METHOD="GET" ACTION="VehicleSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="VID" value="<%=Request.QueryString("VID")%>">
<table width="100%" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
	<table CLASS="LABEL" style="width:300" align="left">
	<tr>
	<tr>
	<tr>
	<td CLASS="LABEL">Policy ID:<br><input SIZE="25" CLASS="LABEL" tabindex="1" MAXLENGTH="10" TYPE="TEXT" NAME="SearchPOLICY_ID" size="20" VALUE="<%=Request.QueryString("SearchPOLICY_ID")%>"></td>
	<td CLASS="LABEL">VIN:<br><input SIZE="25" CLASS="LABEL" MAXLENGTH="40" tabindex="2" TYPE="TEXT" NAME="SearchVIN" size="23" VALUE="<%=Request.QueryString("SearchVIN")%>"></td>
	<td CLASS="LABEL" COLSPAN="2">Year:<br><input MAXLENGTH="4" tabindex="4" CLASS="LABEL" TYPE="TEXT" NAME="SearchYEAR" VALUE="<%=Request.QueryString("SearchYEAR")%>"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">Make:<br><input tabindex="4" SIZE="25" MAXLENGTH="40" CLASS="LABEL" TYPE="TEXT" NAME="SearchMAKE" VALUE="<%=Request.QueryString("SearchMAKE")%>"></td>
	<td CLASS="LABEL">Model:<br><input tabindex="4" SIZE="25" MAXLENGTH="30" CLASS="LABEL" TYPE="TEXT" NAME="SearchMODEL" VALUE="<%=Request.QueryString("SearchMODEL")%>"></td>
	<td CLASS="LABEL" COLSPAN="2">License Plate:<br><input tabindex="4" SIZE="20" MAXLENGTH="30" CLASS="LABEL" TYPE="TEXT" NAME="SearchLICENSE_PLATE" VALUE="<%=Request.QueryString("SearchLICENSE_PLATE")%>"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">Vehicle ID:<br><input tabindex="4" SIZE="25" MAXLENGTH="10" CLASS="LABEL" TYPE="TEXT" NAME="SearchVID" VALUE="<%=Request.QueryString("SearchVID")%>"></td>
	<td CLASS="LABEL">Color:<br><input tabindex="4" SIZE="25" MAXLENGTH="30" CLASS="LABEL" TYPE="TEXT" NAME="SearchCOLOR" VALUE="<%=Request.QueryString("SearchCOLOR")%>"></td>
	<td CLASS="LABEL"><nobr>License State:<br>
	<select NAME="SearchLICENSE_PLATE_STATE" CLASS="LABEL" tabindex="4">
	<option VALUE>
	<!--#include file="..\lib\States.asp"-->
	</select>
	</td>
	<td CLASS="LABEL"><nobr>Reg. State:<br>
	<select NAME="SearchREGISTRATION_STATE" CLASS="LABEL" tabindex="4">
	<option VALUE>
	<!--#include file="..\lib\States.asp"-->
	</select>
	</td>
	
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
