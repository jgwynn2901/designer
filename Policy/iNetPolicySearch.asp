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
<title>Agent Search</title>
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
	document.all.ClientCode.value = ""
	document.all.PolicyID.value = ""
	document.all.CarrierName.value = ""
	document.all.InsuredName.value = ""
	document.all.SearchState.value = ""
	document.all.SearchZip.value = ""
End Sub

Sub BtnSearch_onclick()
	document.all.SpanStatus.innerHTML = "<%=MSG_SEARCH%>"
	FrmSearch.submit
End Sub

Sub window_onload
	document.all.SearchType(0).checked = True
	UpdateStatus("Ready")	
	
<%	If Request.QueryString <> "" Then %>
<%		If Request.QueryString("SearchType") = "B" Then	%>
			document.all.SearchType(0).checked = True
<%		ElseIf Request.QueryString("SearchType") = "C" Then	%>
			document.all.SearchType(1).checked = True
<%		ElseIf Request.QueryString("SearchType") = "E" Then	%>
			document.all.SearchType(2).checked = True
<%		End If 

		If Request.QueryString("SearchState") <> "" Then	%>
			SelectOption document.all.SearchState,"<%=Request.QueryString("SearchState")%>"
<%		End If 

	End If %>	

	If document.all.ClientCode.value <> "" Or document.all.PolicyID.value <> "" Or _
	 document.all.CarrierName.value <> "" Or document.all.InsuredName.value <> "" Or document.all.SearchState.value <> "" Or _
	 document.all.SearchZip.value <> ""  Then
		UpdateStatus("<%=MSG_PROMPT%>")	
	End If
End Sub

Sub PostTo(strURL)
	dim cKey, temp
	
	document.all.txtAction.value = "UPDATE"
	cKey = Parent.frames("WORKAREA").getKey
	temp = Split(cKey, "||")
	If UBound(temp) >= 0 Then 
		document.all.recBM.value = temp(0)
	Else		
		document.all.recBM.value = ""
	End If
	FrmSearch.action = "iNetPolicyDetails-f.asp"
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
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» iNetPolicy Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
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

<form Name="FrmSearch" METHOD="GET" TARGET="WORKAREA" ACTION="iNetPolicySearchResults.asp">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="recBM">
<input type="hidden" NAME="txtAction">
<table width="100%" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
 <table CELLPADDING="0" CELLSPACING="0"> 
	<tr></tr>
	<tr></tr>
	<tr>
	<td CLASS="LABEL" width=30%>Client Code:<br><input CLASS="LABEL" tabindex="1" TYPE="TEXT" SIZE="3" NAME="ClientCode" VALUE="<%=Request.QueryString("ClientCode")%>"></td>					
	<td CLASS="LABEL" width=25%>State:<br><select tabindex="4" NAME="SearchState" CLASS="LABEL"><option VALUE="<%=Request.QueryString("SearchState")%>"><!--#include file="..\lib\states.asp"--></select></td>		
	<td CLASS="LABEL">Zip:<br><input CLASS="LABEL" tabindex="5" TYPE="TEXT" NAME="SearchZip" VALUE="<%=Request.QueryString("SearchZip")%>"></td>
	</tr>
	</table>
	<table>
	<tr>
	<td CLASS="LABEL">Policy ID:<br><input CLASS="LABEL" tabindex="2" TYPE="TEXT" SIZE="80" NAME="PolicyID" VALUE="<%=Request.QueryString("PolicyID")%>"></td>					
	</tr>
	<tr>
	<td CLASS="LABEL">Carrier Name:<br><input CLASS="LABEL" tabindex="1" TYPE="TEXT" SIZE="80" NAME="CarrierName" VALUE="<%=Request.QueryString("CarrierName")%>"></td>
	</tr>
	<tr>
	<td CLASS="LABEL">Insured Name:<br><input CLASS="LABEL" tabindex="1" TYPE="TEXT" SIZE="80" NAME="InsuredName" VALUE="<%=Request.QueryString("InsuredName")%>"></td>												
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
</table>

<table topmargin="0" bottommargin="0">
<tr>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" tabindex="8" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" tabindex="9" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" tabindex="10" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
</tr>
</table>

</form>
</body>
</html>
