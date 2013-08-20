<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Attribute Search</title>
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
	document.all.SearchDID.value = ""
	document.all.SearchPOLICY_ID.value = ""
	document.all.SearchNAME_LAST.value = ""
	document.all.SearchNAME_FIRST.value = ""
	document.all.SearchSSN.value = ""
	document.all.SearchADDRESS.value = ""
	document.all.SearchCITY.value = ""
	document.all.SearchSTATE.value = ""
	document.all.SearchZIP.value = ""
	
End Sub

Sub BtnSearch_onclick()
	'If document.all.SearchDID.value = "" And document.all.SearchVEHICLE_ID.value = "" And _
	'document.all.SearchNAME_LAST.value = "" And document.all.SearchNAME_FIRST.value = "" And _
	'document.all.SearchADDRESS.value = "" And document.all.SearchCITY.value = "" And _
	'document.all.SearchCITY.value = "" And document.all.SearchSTATE.value = "" And _
	'document.all.SearchZIP.value = "" And document.all.SearchSSN.value = "" Then
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
	If document.all.SearchDID.value <> "" And document.all.SearchPOLICY_ID.value <> "" And _
	document.all.SearchNAME_LAST.value <> "" And document.all.SearchNAME_FIRST.value <> "" And _
	document.all.SearchADDRESS.value <> "" And document.all.SearchCITY.value <> "" And _
	document.all.SearchCITY.value <> "" And document.all.SearchSTATE.value <> "" And _
	document.all.SearchZIP.value <> "" And document.all.SearchSSN.value <> "" Then
		UpdateStatus("<%=MSG_PROMPT%>")	
	End If
<% If Request.QueryString("SearchState") <> "" Then %>
document.all.SearchState.Value = "<%= Request.QueryString("SearchState") %>"
<% End If %>
End Sub

Sub PostTo(strURL)

	curDID = Parent.frames("WORKAREA").GetDID
	temp = Split(curDID, "||")
	If UBound(temp) >= 0 Then 
		document.all.DID.value = temp(0)
	Else		
		document.all.DID.value = ""
	End If
	FrmSearch.action = "DriverDetails-f.asp"
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
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Driver Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
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
<form Name="FrmSearch" METHOD="GET" ACTION="DriverSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="DID" value="<%=Request.QueryString("DID")%>">
<table width="100%" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
	<table CLASS="LABEL" style="width:300" align="left">
	<tr>
	<td CLASS="LABEL">Driver ID:<br><input SIZE="22" CLASS="LABEL" tabindex="1" MAXLENGTH="10" TYPE="TEXT" NAME="SearchDID" size="20" VALUE="<%=Request.QueryString("SearchDID")%>"></td>
	<td CLASS="LABEL" ALIGN="LEFT">Policy ID:<br><input SIZE="15" CLASS="LABEL" tabindex="2" MAXLENGTH="10" TYPE="TEXT" NAME="SearchPOLICY_ID" size="20" VALUE="<%=Request.QueryString("SearchPOLICY_ID")%>"></td>
	<td CLASS="LABEL" COLSPAN="2">Last Name:<br><input SIZE="40" CLASS="LABEL" MAXLENGTH="40" tabindex="3" TYPE="TEXT" NAME="SearchNAME_LAST" size="23" VALUE="<%=Request.QueryString("SearchNAME_LAST")%>"></td>
	</tr>
	<tr>
	<td CLASS="LABEL" COLSPAN="2">First Name:<br><input MAXLENGTH="80" SIZE="42" tabindex="4" CLASS="LABEL" TYPE="TEXT" NAME="SearchNAME_FIRST" VALUE="<%=Request.QueryString("SearchNAME_FIRST")%>"></td>
	<td CLASS="LABEL" COLSPAN="2">Address:<br><input tabindex="5" SIZE="40" MAXLENGTH="40" CLASS="LABEL" TYPE="TEXT" NAME="SearchADDRESS" VALUE="<%=Request.QueryString("SearchADDRESS")%>"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">City:<br><input tabindex="6" SIZE="20" MAXLENGTH="30" CLASS="LABEL" TYPE="TEXT" NAME="SearchCITY" VALUE="<%=Request.QueryString("SearchCITY")%>"></td>
	<td CLASS="LABEL">State:<br>
	<select NAME="SearchSTATE" CLASS="LABEL" tabindex="7">
	<option VALUE>
	<!--#include file="..\lib\States.asp"-->
	</select>
	</td>
	<td CLASS="LABEL">Zip:<br><input tabindex="8" SIZE="10" MAXLENGTH="9" CLASS="LABEL" TYPE="TEXT" NAME="SearchZip" VALUE="<%=Request.QueryString("SearchZip")%>"></td>
<td CLASS="LABEL">SSN:<br><input tabindex="9" SIZE="15" MAXLENGTH="9" CLASS="LABEL" TYPE="TEXT" NAME="SearchSSN" VALUE="<%=Request.QueryString("SearchSSN")%>"></td>
	</tr>
	</table>
	
</td>
<td VALIGN="TOP" rowspan="3">
	<table>
	<tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex="13" NAME="BtnSearch" TYPE="BUTTON" ACCESSKEY="H">Searc<u>h</u></button></td></tr>
	<tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex="14" NAME="BtnClear" ACCESSKEY="L">C<u>l</u>ear</button></td></tr>
	</table>
</td>	
</tr>

<tr>
<td>
	<table>
	<tr>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="10" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="11" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="12" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
	</tr>
	</table>
</td>
</tr>
</table>
</form>
</body>
</html>
