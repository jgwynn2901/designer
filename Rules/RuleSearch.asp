<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Rule Search</title>
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
	document.all.SearchRuleId.value = ""
	document.all.SearchRuleType.value = ""
	document.all.SearchRuleText.value = ""
	document.all.SearchComments.value = ""
	document.all.SearchUser.value = ""
End Sub

Sub BtnSearch_onclick()
	'If document.all.SearchRuleId.value = "" And document.all.SearchRuleType.value = "" And _
	 '  document.all.SearchRuleText.value = ""  Then
	'		MsgBox "Please enter search criteria.", 0, "FNSNetDesigner"
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
	
<%		If CStr(Request.QueryString("SearchRuleType")) <> "" Then	%>
			SelectOption document.all.SearchRuleType,"<%=CStr(Request.QueryString("SearchRuleType"))%>"
<%		End If %>		
<%	End If %>	

	If	document.all.SearchRuleId.value <> ""	Or _ 
		document.all.SearchRuleType.value <> ""	Or _
		document.all.SearchRuleText.value <> ""	Or _ 
		document.all.SearchComments <> "" 		Or _
		document.all.SearchUser <> "" 	Then
		UpdateStatus("<%=MSG_PROMPT%>")	
	End If
End Sub


Sub PostTo(strURL)
	document.all.RID.value = Parent.frames("WORKAREA").GetRID
	FrmSearch.action = "RuleEditor-f.asp"
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
<body BGCOLOR="<%=BODYBGCOLOR%>"  topmargin=0 leftmargin=0  rightmargin=0>
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Rule Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
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

<form Name="FrmSearch" METHOD="GET" ACTION="RuleSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="RID" value="<%=Request.QueryString("RID")%>">

<table CLASS="LABEL" width="100%" CELLPADDING="0" CELLSPACING="0">
<tr>
 <td>
  <table border=0 CLASS="LABEL" style="width:300" align=left>
    <tr>
    <tr>
    
		<td style="width:220">Type:<br><select NAME="SearchRuleType" tabindex=1 CLASS="LABEL"><%=GetValidValuesHTML("RULES_TYPE","",true)%></select></td>
		<td colspan="2">Rule Text:<br><input class="LABEL" name="SearchRuleText" type="text" size="30" tabindex=2 VALUE="<%=Request.QueryString("SearchRuleText")%>">
	</tr>
	<tr>
		<td>Rule ID:<br><input class="LABEL" name="SearchRuleId" type="text" size="16" tabindex=3 VALUE="<%=Request.QueryString("SearchRuleId")%>" ID="Text1"></td>
		<td>Comments:<br><input class="LABEL" name="SearchComments" type="text" size="30" tabindex=4 VALUE="<%=Request.QueryString("SearchComments")%>" ID="Text2"></td>
		
	</tr>
	<tr> <td>User:<br><input class="LABEL" name="SearchUser" type="text" size="20" tabindex=5 VALUE="<%=Request.QueryString("SearchUser")%>" ID="Text3">
		</td>
	</tr>
   
   </table>
 </td>

 <td VALIGN=top>
   <table>
	<tr><td><button CLASS="StdButton" NAME="BtnSearch" ACCESSKEY="H"  tabindex=7>Searc<u>h</u></button></td></tr>
	<tr><td><button CLASS="StdButton" NAME="BtnClear" ACCESSKEY="L"  tabindex=8>C<u>l</u>ear</button></td></tr>
	</table>
 </td>
</tr>
<tr>
<td>
<table>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL" tabindex=4>Begins With</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL" tabindex=5>Contains</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL" tabindex=6>Exact</td>

</table>
</td>
</tr>
</table>
</body>
</form>
</html>
