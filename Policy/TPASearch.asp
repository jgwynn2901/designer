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
	document.all.SearchTPAID.value = ""
	document.all.SearchName.value = ""
	document.all.SearchTitle.value = ""
	document.all.SearchBusinessType.value = ""
	document.all.SearchAddress.value = ""
	document.all.SearchCity.value = ""
	document.all.SearchState.value = ""
	document.all.SearchZip.value = ""
	document.all.SearchTPANumber.value = ""
End Sub

Sub BtnSearch_onclick()
	'If document.all.SearchTPAID.value = "" And document.all.SearchName.value = "" And _
	'document.all.SearchTitle.value = "" And document.all.SearchBusinessType.value = "" And _
	'document.all.SearchState.value = "" And document.all.SearchZip.value = "" And _
	'document.all.SearchAddress.value = ""  And document.all.SearchCity.value = "" Then
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

If document.all.SearchTPAID.value <> "" And document.all.SearchName.value <> "" And _
	document.all.SearchTitle.value <> "" And document.all.SearchBusinessType.value <> "" And _
	document.all.SearchState.value <> "" And document.all.SearchZip.value <> "" And _
	document.all.SearchAddress.value <> ""  And document.all.SearchCity.value <> "" And _
	document.all.SearchTPANumber.value <> ""Then

		UpdateStatus("<%=MSG_PROMPT%>")	
	End If
<% If Request.QueryString("SearchState") <> "" Then %>
document.all.SearchState.Value = "<%= Request.QueryString("SearchState")%>"
<%End If %>
End Sub

Sub PostTo(strURL)

	curTPAID = Parent.frames("WORKAREA").GetTPAID
	temp = Split(curTPAID, "||")
	If UBound(temp) >= 0 Then 
		document.all.TPAID.value = temp(0)
	Else		
		document.all.TPAID.value = ""
	End If
	FrmSearch.action = "TPADetails-f.asp"
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
<BODY  topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 TPA Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
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
<form Name="FrmSearch" METHOD="GET" ACTION="TPASearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="TPAID" value="<%=Request.QueryString("TPAID")%>">
<table width=100% CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
	<table CLASS="LABEL" style="width:300" align=left>
	<tr>
	<tr>
	<Tr>
		<Td CLASS="LABEL">TPA Number:<br><input CLASS="LABEL" tabindex=1 TYPE="TEXT" MAXLENGTH="30" NAME="SearchTPANumber" VALUE="<%=Request.QueryString("TPA_NUMBER")%>"></td>
		<td CLASS="LABEL">Name:<br><input CLASS="LABEL" tabindex=2 TYPE="TEXT" NAME="SearchName" size=20 VALUE="<%=Request.QueryString("SearchName")%>"></td>
	</Tr>
	<tr>
		<td CLASS="LABEL">TPA ID:<br><input CLASS="LABEL" TYPE="TEXT" tabindex=3 NAME="SearchTPAID" VALUE="<%=Request.QueryString("SearchTPAID")%>"></td>
		<td CLASS="LABEL">Business Type:<br><input tabindex=4 CLASS="LABEL" TYPE="TEXT" NAME="SearchBusinessType" VALUE="<%=Request.QueryString("SearchBusinessType")%>"></td>
		<td CLASS="LABEL" COLSPAN=2>Title:<br><input CLASS="LABEL" tabindex=5 TYPE="TEXT" NAME="SearchTitle" size=23 VALUE="<%=Request.QueryString("SearchTitle")%>"></td>
	</TR>
	<TR>
		<td CLASS="LABEL">Address:<br><input tabindex=6 CLASS="LABEL" TYPE="TEXT" NAME="SearchAddress" VALUE="<%=Request.QueryString("SearchAddress")%>"></td>
		<td CLASS="LABEL">City:<br><input tabindex=7 CLASS="LABEL" TYPE="TEXT" NAME="SearchCity" VALUE="<%=Request.QueryString("SearchCity")%>"></td>
		<td CLASS="LABEL">State:<br>
			<SELECT NAME=SearchState CLASS=LABEL tabindex=8>
				<OPTION VALUE="">
				<!--#include file="..\lib\States.asp"-->
			</SELECT>
		</TD>
		<td CLASS="LABEL">Zip:<br><input tabindex=9 SIZE=7 CLASS="LABEL" TYPE="TEXT" NAME="SearchZip" VALUE="<%=Request.QueryString("SearchZip")%>"></td>
	</tr>
	</table>
</td>
<td VALIGN=TOP rowspan="3" >
	<TABLE>
	<TR><TD CLASS=LABEL><BUTTON CLASS=StdButton tabindex=12 NAME=BtnSearch TYPE="BUTTON" ACCESSKEY="H">Searc<U>h</U></BUTTON></TD></TR>
	<TR><TD CLASS=LABEL><BUTTON CLASS=StdButton tabindex=13 NAME=BtnClear ACCESSKEY="L">C<U>l</U>ear</BUTTON></TD></TR>
	</TABLE>
</td>	
</tr>
<tr>
<td>
	<table>
	<tr>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex=9 STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex=10 STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex=11 STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
	</tr>
	</table>
</td>
</tr>
</table>


</form>
</body>
</html>
