<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Routing Address Search</title>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

Sub BtnClear_onclick()
	document.all.SearchRAID.value = ""
	document.all.SearchDescription.value = ""
	document.all.SearchState.value = ""
	document.all.SearchFIPS.value = ""
	document.all.SearchZip.value = ""
		
End Sub

Sub BtnSearch_onclick()
	'If document.all.SearchRAID.value = "" And document.all.SearchDescription.value = "" And _
	 '  document.all.SearchState.value = "" And document.all.SearchFIPS.value = "" And _
	  ' document.all.SearchZip.value = "" Then
		'	MsgBox "Please enter search criteria.", 0, "FNSNetDesigner"
	'Else
		document.all.SpanStatus.innerHTML = "<%=MSG_SEARCH%>"
		FrmSearch.submit
	'End If
End Sub

Sub window_onload
	'document.all.SearchRuleId.focus timing issue
	FrmSearch.SearchType(1).checked = True
	UpdateStatus("Ready")	

<%	If Request.QueryString <> "" Then %>
<%		If CStr(Request.QueryString("SearchType")) = "B" Then	%>
			document.all.SearchType(0).checked = True
<%		ElseIf CStr(Request.QueryString("SearchType")) = "C" Then	%>
			document.all.SearchType(1).checked = True
<%		ElseIf CStr(Request.QueryString("SearchType")) = "E" Then	%>
			document.all.SearchType(2).checked = True
<%		End If %>				
<%	End If %>	

	If document.all.SearchRAID.value <> "" Or document.all.SearchDescription.value <> "" Or _
	   document.all.SearchState.value <> "" Or  document.all.SearchFIPS.value <> "" Or _
		document.all.SearchZip.value <> "" Then
		UpdateStatus("<%=MSG_PROMPT%>")	
	End If
End Sub


Sub PostTo(strURL)
	document.all.RAID.value = Parent.frames("WORKAREA").GetRAID
	FrmSearch.action = "RoutingAddressDetails-f.asp"
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
</script>

</head>
<body BGCOLOR="<%=BODYBGCOLOR%>"  topmargin=0 leftmargin=0  rightmargin=0>
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Routing Address Search</TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
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

<form Name="FrmSearch" METHOD="GET" ACTION="RoutingAddressSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="RAID" value="<%=Request.QueryString("RAID")%>">

<table CLASS="LABEL" width="100%" CELLPADDING="0" CELLSPACING="0">
<tr>
 <td>
  <table border=0 CLASS="LABEL" style="width:300" align=left>
    <tr>
    <tr>
    <tr>
	<tr><td colspan=3>Description:<br><input type="text" NAME="SearchDescription" tabindex=1 CLASS="LABEL" size=80 value="<%=Request.QueryString("SearchDescription")%>"></input></td></tr>
	<tr>
		<td>State:<br><input type="text" NAME="SearchState" tabindex=2 CLASS="LABEL" size=24 value="<%=Request.QueryString("SearchState")%>"></input></td>
		<td>FIPS:<br><input type="text" NAME="SearchFIPS" tabindex=3 CLASS="LABEL" size=24 value="<%=Request.QueryString("SearchFIPS")%>"></input></td>
		<td>Zip:<br><input type="text" NAME="SearchZip" tabindex=4 CLASS="LABEL" size=24 value="<%=Request.QueryString("SearchZip")%>"></input></td>
	</tr>
	<tr><td>Routing Address ID:<br><input class="LABEL" name="SearchRAID" type="text" size="16" tabindex=5 VALUE="<%=Request.QueryString("SearchRAID")%>"></td></tr>
   </table>
 </td>
 <td VALIGN=top>
   <table>
	<tr><td><button CLASS="StdButton" NAME="BtnSearch" ACCESSKEY="H"  tabindex=9>Searc<u>h</u></button></td></tr>
	<tr><td><button CLASS="StdButton" NAME="BtnClear" ACCESSKEY="L"  tabindex=10>C<u>l</u>ear</button></td></tr>
	</table>
 </td>
</tr>
<tr>
<td>
<table>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL" tabindex=6>Begins With</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL" tabindex=7>Contains</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL" tabindex=8>Exact</td>
</table>
</td>
</tr>
</table>
</body>
</form>
</html>
