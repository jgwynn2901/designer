<!--#include file="..\lib\common.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<!--#include file="..\lib\tablecommon.inc"-->
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Rule Search</title>

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub BtnClear_onclick()
	document.all.OUTPUTDEF_ID.value = ""
	document.all.NAME.value = ""
	document.all.DESCRIPTION.value = ""
End Sub

Sub BtnSearch_onclick()
	If document.all.OUTPUTDEF_ID.value = "" And document.all.NAME.value = "" And _
	   document.all.DESCRIPTION.value = ""  Then
			MsgBox "Please enter search criteria.", 0, "FNSNetDesigner"
	Else
		FrmSearch.submit
	End If
End Sub

Sub window_onload
	'document.all.SearchRuleId.focus timing issue
	FrmSearch.SearchType(0).checked = True
End Sub


</script>

</head>
<body BGCOLOR='<%=BODYBGCOLOR%>'   leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0>
<form Name="FrmSearch" METHOD="GET" ACTION="ODSearchResults.asp" TARGET="WORKAREA">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Output Definition Search</TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<table CLASS="LABEL" width="100%" CELLSPACING=0 CELLPADDING=0>
<tr>
<td>Output ID:<br><input class="LABEL" TABINDEX=1 name="OUTPUTDEF_ID" type="text" size="10"></td>
<td>Name:<br><input class="LABEL" TABINDEX=2 name="NAME" type="text" size="40" ></td>
<td rowspan="2" VALIGN="TOP"> 
<table>
<tr><td><button CLASS="StdButton" NAME="BtnSearch" ACCESSKEY="C">Sear<u>c</u>h</button></td></tr>
<tr><td><button CLASS="StdButton" NAME="BtnClear" ACCESSKEY="L">C<u>l</u>ear</button></td></tr>
</table>
</td>
</tr>
<tr><td colspan="2">Description:<br><input TABINDEX=3 class="LABEL" name="DESCRIPTION" type="text" size="50"></td></tr>
</table>
<table>
<tr>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
</tr>
</table>
</body>
</form>
</html>
