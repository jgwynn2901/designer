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
	document.all.SearchABID.value = ""
	document.all.SearchNAME.value = ""
	document.all.SearchDESCRIPTION.value = ""
	document.all.SearchCFID.value = ""
End Sub

Sub BtnSearch_onclick()
	'If document.all.SearchABID.value = "" And document.all.SearchNAME.value = "" And _
	'document.all.SearchDESCRIPTION.value = "" Then
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
	If document.all.SearchABID.value <> "" And document.all.SearchNAME.value <> "" And _
	document.all.SearchDESCRIPTION.value <> "" Then
		UpdateStatus("<%=MSG_PROMPT%>")	
	End If

End Sub

Sub PostTo(strURL)

	curABID = Parent.frames("WORKAREA").GetABID
	temp = Split(curABID, "||")
	If UBound(temp) >= 0 Then 
		document.all.ABID.value = temp(0)
	Else		
		document.all.ABID.value = ""
	End If
	FrmSearch.action = "AddressBookDetails-f.asp"
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
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Address Book Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
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
<form Name="FrmSearch" METHOD="GET" ACTION="AddressBookSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="ABID" value="<%=Request.QueryString("ABID")%>">
<table width=100% CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
	<table CLASS="LABEL" style="width:300" align=left CELLSPACING=0 CELLPADDING=0>
	<TR>
	<td CLASS="LABEL" COLSPAN=3>Description:<br><input SIZE=78 CLASS="LABEL" MAXLENGTH=80 tabindex=3 TYPE="TEXT" NAME="SearchDESCRIPTION" size=23 VALUE="<%=Request.QueryString("SearchDESCRIPTION")%>"></td>
	</TR>
	<TR>
	<td CLASS="LABEL"><NOBR>Add. Book ID:<br><input SIZE=10 CLASS="LABEL" tabindex=1 MAXLENGTH=10 TYPE="TEXT" NAME="SearchABID" size=20 VALUE="<%=Request.QueryString("SearchABID")%>"></td>
	<td CLASS="LABEL"><NOBR>Call Flow ID:<br><input SIZE=10 CLASS="LABEL" tabindex=1 MAXLENGTH=10 TYPE="TEXT" NAME="SearchCFID" size=20 VALUE="<%=Request.QueryString("SearchCFID")%>"></td>
	<td CLASS="LABEL" ALIGN=LEFT><NOBR>Name:<br><input SIZE=50 CLASS="LABEL" tabindex=2 MAXLENGTH=40 TYPE="TEXT" NAME="SearchNAME" size=40 VALUE="<%=Request.QueryString("SearchNAME")%>"></td>
	</TR>
	</table>
</td>
<td VALIGN=TOP rowspan="3" >
	<TABLE>
	<TR><TD CLASS=LABEL><BUTTON CLASS=StdButton tabindex=10 NAME=BtnSearch TYPE="BUTTON" ACCESSKEY="H">Searc<U>h</U></BUTTON></TD></TR>
	<TR><TD CLASS=LABEL><BUTTON CLASS=StdButton tabindex=11 NAME=BtnClear ACCESSKEY="L">C<U>l</U>ear</BUTTON></TD></TR>
	</TABLE>
</td>	
</tr>
<tr>
<td>
	<table>
	<tr>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex=7 STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex=8 STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex=9 STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
	</tr>
	</table>
</td>
</tr>
</table>
</form>
</body>
</html>
