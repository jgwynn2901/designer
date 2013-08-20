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
<title>Branch Search</title>
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
	document.all.SearchBranchNumber.value = ""
	document.all.SearchAHLoadID.value = ""
	document.all.SearchStatus.value = ""
	document.all.SearchOfficeNumber.value = ""
	document.all.SearchOfficeType.value = ""
	document.all.SearchOfficeName.value = ""
	document.all.SearchAddress.value = ""
	document.all.SearchCity.value = ""
	document.all.SearchState.value = ""
	document.all.SearchZip.value = ""
	document.all.SearchBID.value = ""
	document.all.SearchBranchType.value = document.all.BranchTypeFilter.value
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

		If CStr(Request.QueryString("SearchState")) <> "" Then	%>
			SelectOption document.all.SearchState,"<%=CStr(Request.QueryString("SearchState"))%>"
<%		End If 

		If CStr(Request.QueryString("BranchTypeFilter")) <> "" Then %>
			SelectOption document.all.SearchBranchType,"<%=CStr(Request.QueryString("BranchTypeFilter"))%>"
<%
		Else
			If CStr(Request.QueryString("SearchBranchType")) <> "" Then	%>
				SelectOption document.all.SearchBranchType,"<%=CStr(Request.QueryString("SearchBranchType"))%>"
<%			End If
		End If 

	End If %>	

	If document.all.SearchBranchNumber.value <> "" Or document.all.SearchAHLoadID.value <> "" Or _
	document.all.SearchStatus.value <> "" Or document.all.SearchOfficeNumber.value <> "" Or _
	document.all.SearchOfficeType.value <> ""  Or document.all.SearchOfficeName.value <> "" Or _
	document.all.SearchAddress.value <> ""  Or document.all.SearchCity.value <> "" Or _
	document.all.SearchState.value <> ""  Or document.all.SearchZip.value <> "" Or _
	document.all.SearchBranchType.value <> "" Or _
	document.all.SearchBID.value <> ""  Then
		UpdateStatus("<%=MSG_PROMPT%>")	
	End If

<%	if CStr(Request.QueryString("BranchTypeFilter")) <> "" then %>
	SetBranchTypeFieldReadOnly true 
<%	end if %>

End Sub

Sub PostTo(strURL)
	curBID = Parent.frames("WORKAREA").GetBID
	temp = Split(curBID, "||")
	If UBound(temp) >= 0 Then 
		document.all.BID.value = temp(0)
	Else		
		document.all.BID.value = ""
	End If
	FrmSearch.action = "BranchDetails-f.asp"
	FrmSearch.method = "GET"	
	FrmSearch.target = "_parent"	
	FrmSearch.submit
End Sub

sub SetBranchTypeFieldReadOnly(bReadOnly)
	for iCount = 0 to document.all.length-1
		if document.all(iCount).getAttribute("SpecialFilterBtn") = "TRUE" then
			document.all(iCount).disabled = bReadOnly
		end if
	next
end sub

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
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Branch Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
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

<form Name="FrmSearch" METHOD="GET" ACTION="BranchSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="BID" value="<%=Request.QueryString("BID")%>">
<input type="hidden" NAME="BranchTypeFilter" value="<%=Request.QueryString("BranchTypeFilter")%>">
<table width=100% CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
	<table CLASS="LABEL" style="width:300" align=left>
	<tr>
	<tr>
	<tr>
	<td CLASS="LABEL">Branch Number:<br><input size=25 tabindex=1 CLASS="LABEL" TYPE="TEXT" NAME="SearchBranchNumber"  VALUE="<%=Request.QueryString("SearchBranchNumber")%>"></td>
	<td CLASS="LABEL">AH Load ID:<br><input size=25 tabindex=2 CLASS="LABEL" TYPE="TEXT" NAME="SearchAHLoadID" VALUE="<%=Request.QueryString("SearchAHLoadID")%>"></td>
	<td COLSPAN=2 CLASS="LABEL">Status:<br><input tabindex=3 CLASS="LABEL" TYPE="TEXT" NAME="SearchStatus" VALUE="<%=Request.QueryString("SearchStatus")%>"></td>
	<tr>
	<td CLASS="LABEL">Office Number:<br><input size=25 tabindex=4 CLASS="LABEL" TYPE="TEXT" NAME="SearchOfficeNumber"  VALUE="<%=Request.QueryString("SearchOfficeNumber")%>"></td>
	<td CLASS="LABEL">Office Type:<br><input  size=25 tabindex=5 CLASS="LABEL" TYPE="TEXT" NAME="SearchOfficeType" VALUE="<%=Request.QueryString("SearchOfficeType")%>"></td>
	<td COLSPAN=2 CLASS="LABEL">Office Name:<br><input tabindex=6 size=25 CLASS="LABEL" TYPE="TEXT" NAME="SearchOfficeName" VALUE="<%=Request.QueryString("SearchOfficeName")%>"></td>	
	</tr>
	<tr>
	<td CLASS="LABEL">Address:<br><input size=25 tabindex=7 CLASS="LABEL" TYPE="TEXT" NAME="SearchAddress"  VALUE="<%=Request.QueryString("SearchAddress")%>"></td>
	<td CLASS="LABEL">City:<br><input size=25 tabindex=8 CLASS="LABEL" TYPE="TEXT" NAME="SearchCity" VALUE="<%=Request.QueryString("SearchCity")%>"></td>
	<td CLASS="LABEL">State:<br><SELECT tabindex=9 NAME=SearchState CLASS=LABEL><OPTION VALUE=""><!--#include file="..\lib\states.asp"--></SELECT></td>	
	<td CLASS="LABEL">Zip:<br><input size=16 tabindex=10 CLASS="LABEL" TYPE="TEXT" NAME="SearchZip" VALUE="<%=Request.QueryString("SearchZip")%>"></td>		
	</tr>
	<tr>
	<td CLASS="LABEL">Branch ID:<br><input CLASS="LABEL" tabindex=11 TYPE="TEXT"  NAME="SearchBID" VALUE="<%=Request.QueryString("SearchBID")%>"></td>
	<td CLASS="LABEL">Branch Type:<br><SELECT SpecialFilterBtn="TRUE" tabindex=12 NAME=SearchBranchType CLASS=LABEL><OPTION VALUE=""></OPTION><OPTION VALUE="CLAIMHANDLING">CLAIMHANDLING</OPTION><OPTION VALUE="MANAGEDCARE">MANAGEDCARE</OPTION></SELECT></td>	
	</tr>
	</table>
</td>
<td VALIGN=TOP rowspan="3" >
	<TABLE>
	<TR><TD CLASS=LABEL><BUTTON CLASS=StdButton tabindex=16 NAME=BtnSearch TYPE="BUTTON" ACCESSKEY="H">Searc<U>h</U></BUTTON></TD></TR>
	<TR><TD CLASS=LABEL><BUTTON CLASS=StdButton tabindex=17 NAME=BtnClear ACCESSKEY="L">C<U>l</U>ear</BUTTON></TD></TR>
	</TABLE>
</td>	
</tr>
<tr>
<td>
	<table>
	<tr>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex=13 STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex=14 STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
	<td CLASS="LABEL"><input TYPE="RADIO" tabindex=15 STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
	</tr>
	</table>
</td>
</tr>
</table>


</form>
</body>
</html>
