<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ControlData.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<!--#include file="..\lib\tablecommon.inc"-->
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Frame Search</title>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub BtnClear_onclick()
	document.all.FRAME_ID.value = ""
	document.all.NAME.value = ""
	document.all.TITLE.value = ""
	document.all.AHS_ID.value = ""
	document.all.ClientNode_ID.value = ""
	document.all.LOB_CD.value = ""
End Sub

Sub UpdateStatus(inStatus)
	document.all.SpanStatus.innerHTML = inStatus
End Sub

Sub BtnSearch_onclick()
	document.all.SpanStatus.innerHTML = "<%=MSG_SEARCH%>"
	FrmSearch.submit
End Sub

Sub window_onload
	'document.all.SearchRuleId.focus timing issue
	FrmSearch.SearchType(0).checked = True
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
</script>

</head>
<body BGCOLOR="#d6cfbd"  leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0>
<form Name="FrmSearch" METHOD="GET" ACTION="FrameSearchResults.asp" TARGET="WORKAREA">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
	<TR><TD colspan=2 HEIGHT=4></TD></TR>
	<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Frame Search</TD>
		<TD HEIGHT=5 ALIGN=LEFT>
			<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
				<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
				<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
					<TD WIDTH=300 HEIGHT=8></TD></TR>
			</TABLE></TD></TR>
	<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
	<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
<table LANGUAGE="JScript" ONDRAGSTART="return false;" class="Label"  cellspacing=0 cellpadding=0>
	<tr><td VALIGN="CENTER" WIDTH="5" >
			<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER"></td>
			<td width="485">
				:<SPAN VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL>Ready</SPAN></td>
	</tr>
</table>
<table CLASS="LABEL" width="100%">
	<tr><td>A.H.S.ID:<br><input class="LABEL" TabIndex=1 Name="AHS_ID" Type="Text" Size="10" onfocus="enable_exact()" onBlur="enable_begin()"></td>
		<td>Client Node ID:<br><input class="LABEL" TabIndex=2 Name="ClientNode_ID" Type="Text" Size="10"></td></tr>
	<tr><td>Frame ID:<br><input class="LABEL" TABINDEX=3 name="FRAME_ID" type="text" size="10"></td>
		<td>Name:<br><input class="LABEL" TABINDEX=4 name="NAME" type="text" size="33" ></td>
		<td rowspan="2" VALIGN="TOP"> 
			<table>
				<tr><td><button CLASS="StdButton" NAME="BtnSearch" ACCESSKEY="C">Sear<u>c</u>h</button></td></tr>
				<tr><td><button CLASS="StdButton" NAME="BtnClear" ACCESSKEY="L">C<u>l</u>ear</button></td></tr>
			</table>
			</td></tr>
	<tr><td>L.O.B:<br><Select Class="LABEL" TabIndex=5 Name="LOB_CD"><%=GetControlDataHTML("LOB", "LOB_CD", "LOB_CD", "", true)%></Select></td>
		<td colspan="2">Title:<br><input TABINDEX=6 class="LABEL" name="TITLE" type="text" size="33"></td></tr>
</table>
<table>
	<tr><td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
		<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
		<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td></tr>
</table>
</body>
</form>
</html>
