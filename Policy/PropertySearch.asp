<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Property Search</title>
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
	document.all.SearchPROPID.value = ""
	document.all.SearchPOLICY_ID.value = ""
	document.all.SearchPROPERTY_DESCRIPTION.value = ""
	document.all.SearchADDRESS.value = ""
	document.all.SearchCITY.value = ""
	document.all.SearchSTATE.value = ""
	document.all.SearchZIP.value = ""
	
End Sub

Sub BtnSearch_onclick()
	If document.all.SearchPROPID.value = "" And document.all.SearchPOLICY_ID.value = "" And _
	document.all.SearchPROPERTY_DESCRIPTION.value = "" And document.all.SearchSTATE.value = "" And _
	document.all.SearchADDRESS.value = "" And document.all.SearchCITY.value = "" And _
	document.all.SearchZIP.value = "" Then
			MsgBox "Please enter search criteria", 0, "FNSNetDesigner"
	Else
		document.all.SpanStatus.innerHTML = "<%=MSG_SEARCH%>"
		FrmSearch.submit
	End If
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
	If document.all.SearchPROPID.value <> "" And document.all.SearchPOLICY_ID.value <> "" And _
	document.all.SearchPROPERTY_DESCRIPTION.value <> "" And document.all.SearchSTATE.value <> "" And _
	document.all.SearchADDRESS.value <> "" And document.all.SearchCITY.value <> "" And _
	document.all.SearchZIP.value <> "" Then
		UpdateStatus("<%=MSG_PROMPT%>")	
	End If

End Sub

Sub PostTo(strURL)

	curPROPID = Parent.frames("WORKAREA").GetPROPID
	temp = Split(curPROPID, "||")
	If UBound(temp) >= 0 Then 
		document.all.PROPID.value = temp(0)
	Else		
		document.all.PROPID.value = ""
	End If
	FrmSearch.action = "PropertyDetails-f.asp"
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
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Property Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
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
<form Name="FrmSearch" METHOD="GET" ACTION="PropertySearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="PROPID" value="<%=Request.QueryString("PROPID")%>">
<table width="100%" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
	<table CLASS="LABEL" style="width:300" align="left">
	<tr>
	<td CLASS="LABEL">Property ID:<br><input SIZE="22" CLASS="LABEL" tabindex="1" MAXLENGTH="10" TYPE="TEXT" NAME="SearchPROPID" size="20" VALUE="<%=Request.QueryString("SearchPROPID")%>"></td>
	<td CLASS="LABEL" COLSPAN="2">Description:<br><input MAXLENGTH="80" SIZE="42" tabindex="2" CLASS="LABEL" TYPE="TEXT" NAME="SearchPROPERTY_DESCRIPTION" VALUE="<%=Request.QueryString("SearchPROPERTY_DESCRIPTION")%>"></td>
	</tr>
	<tr>
	<td CLASS="LABEL" COLSPAN="2">Address:<br><input tabindex="3" SIZE="40" MAXLENGTH="40" CLASS="LABEL" TYPE="TEXT" NAME="SearchADDRESS" VALUE="<%=Request.QueryString("SearchADDRESS")%>"></td>
	<td CLASS="LABEL" ALIGN="LEFT">Policy ID:<br><input SIZE="15" CLASS="LABEL" tabindex="4" MAXLENGTH="10" TYPE="TEXT" NAME="SearchPOLICY_ID" size="20" VALUE="<%=Request.QueryString("SearchPOLICY_ID")%>"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">City:<br><input tabindex="5" SIZE="20" MAXLENGTH="30" CLASS="LABEL" TYPE="TEXT" NAME="SearchCITY" VALUE="<%=Request.QueryString("SearchCITY")%>"></td>
	<td CLASS="LABEL">State:<br>
	<select NAME="SearchSTATE" CLASS="LABEL" tabindex="6">
	<option VALUE>
	<!--#include file="..\lib\States.asp"-->
	</select>
	</td>
	<td CLASS="LABEL">Zip:<br><input tabindex="7" SIZE="10" MAXLENGTH="9" CLASS="LABEL" TYPE="TEXT" NAME="SearchZip" VALUE="<%=Request.QueryString("SearchZip")%>"></td>
	</tr>
	</table>
	
</td>
<td VALIGN="TOP" rowspan="3">
	<table>
	<tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex="11" NAME="BtnSearch" TYPE="BUTTON" ACCESSKEY="H">Searc<u>h</u></button></td></tr>
	<tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex="12" NAME="BtnClear" ACCESSKEY="L">C<u>l</u>ear</button></td></tr>
	</table>
	
</td>	
</tr>
<tr>
<td>
	<table>
	<tr>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="8" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="9" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="10" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
	</tr>
	</table>
</td>
</tr>
</table>
</form>
</body>
</html>
