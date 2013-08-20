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
<title>Vendor Search</title>
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
	document.all.SearchVendorID.value = ""
	document.all.SearchName.value = ""
	document.all.SearchAddress.value = ""
	document.all.SearchCity.value = ""
	document.all.SearchState.value = ""
	document.all.SearchZip.value = ""
	document.all.SearchServType.value = ""
	document.all.SearchEnabled.checked = false
End Sub

Sub BtnSearch_onclick()
	document.all.SpanStatus.innerHTML = "<%=MSG_SEARCH%>"
	FrmSearch.submit
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

		If Request.QueryString("SearchState") <> "" Then	%>
			SelectOption document.all.SearchState,"<%=Request.QueryString("SearchState")%>"
<%		End If 

		If Request.QueryString("ServiceType") <> "" Then %>
			SelectOption document.all.SearchServType,"<%=Request.QueryString("ServiceType")%>"
<%
		End If 
	End If %>	

	If document.all.SearchVendorID.value <> "" Or document.all.SearchName.value <> "" Or _
	document.all.SearchAddress.value <> ""  Or document.all.SearchCity.value <> "" Or _
	document.all.SearchState.value <> ""  Or document.all.SearchZip.value <> "" Or _
	document.all.SearchServType.value <> "" Or _
	document.all.SearchEnabled.value <> "" Or _
	document.all.SearchVendorID.value <> ""  Then
		UpdateStatus("<%=MSG_PROMPT%>")	
	End If

End Sub

Sub PostTo(strURL)
	curVID = Parent.frames("WORKAREA").GetVID
	temp = Split(curVID, "||")
	If UBound(temp) >= 0 Then 
		document.all.VID.value = temp(0)
	Else		
		document.all.VID.value = ""
	End If
	FrmSearch.action = "VendorDetails-f.asp"
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
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Vendor Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
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

<form Name="FrmSearch" METHOD="GET" ACTION="VendorSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="VID" value="<%=Request.QueryString("VID")%>">
<table width="100%" CELLPADDING="0" CELLSPACING="0">
<tr>
<td width="600">
	<table CLASS="LABEL" style="width:250" align="left">
	<tr>
	<td CLASS="LABEL">Vendor ID:<br><input size="5" tabindex="1" CLASS="LABEL" TYPE="TEXT" NAME="SearchVendorID" COLSPAN=2 VALUE="<%=Request.QueryString("SearchVendorID")%>"></td>
	<td CLASS="LABEL">Name:<br><input size="25" tabindex="2" CLASS="LABEL" TYPE="TEXT" NAME="SearchName" VALUE="<%=Request.QueryString("SearchName")%>"></td>
    <td >&nbsp;</td>
    <td >&nbsp;</td>
    </tr>
	<tr>
	<td CLASS="LABEL" >Address:<br><input size="25" tabindex="3" CLASS="LABEL" TYPE="TEXT" NAME="SearchAddress" VALUE="<%=Request.QueryString("SearchAddress")%>"></td>
	<td CLASS="LABEL">City:<br><input size="25" tabindex="4" CLASS="LABEL" TYPE="TEXT" NAME="SearchCity" VALUE="<%=Request.QueryString("SearchCity")%>"></td>
	<td CLASS="LABEL">State:<br><select tabindex="5" NAME="SearchState" CLASS="LABEL"><option VALUE><!--#include file="..\lib\states.asp"--></select></td>	
	<td CLASS="LABEL">Zip:<br><input size="16" tabindex="6" CLASS="LABEL" TYPE="TEXT" NAME="SearchZip" VALUE="<%=Request.QueryString("SearchZip")%>"></td>		
	</tr>
	<tr>
	<td CLASS="LABEL" COLSPAN="2">Service Type:<br>
	<select ID="SearchServType" CLASS="LABEL" ScrnBtn="TRUE">
	<option VALUE>
	<%
	SQLST = "SELECT * FROM SERVICE_TYPE WHERE TYPE IS NOT NULL"
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	Set RS = oConn.Execute(SQLST)
	Do While Not RS.EOF
	%>
	<option VALUE="<%=RS("SERVICE_TYPE_ID")%>"><%= RS("TYPE") %>
	<%
	RS.MoveNext
	Loop
	RS.CLose
	oConn.close
	set RS=nothing
	set oConn=nothing
	%>
	</select></td>
    <td ><input type="checkbox" name="SearchEnabled" CHECKED> Enabled</td>
    <td >&nbsp;</td>
    <td >&nbsp;</td>
	</tr>
	</table>
</td>
<td VALIGN="TOP" rowspan="3">
	<table>
	<tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex="16" NAME="BtnSearch" TYPE="BUTTON" ACCESSKEY="H">Searc<u>h</u></button></td></tr>
	<tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex="17" NAME="BtnClear" ACCESSKEY="L">C<u>l</u>ear</button></td></tr>
	</table>
</td>	
</tr>
<tr>
<td>
	<table>
	<tr>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="13" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="14" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex="15" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
	</tr>
	</table>
</td>
</tr>
</table>
</form>
</body>
</html>
