<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"

%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Owner Search</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">

Sub BtnClear_onclick()
	document.all.SearchOID.value = ""
	document.all.SearchTitle.value = ""
	document.all.SearchNameLast.value = ""
	document.all.SearchNameFirst.value = ""
	document.all.SearchAdd1.value = ""
	document.all.SearchAdd2.value = ""
	document.all.SearchCity.value = ""
	document.all.SearchState.value = ""
	document.all.SearchZip.value = ""
	document.all.SearchWphone.value = ""
	document.all.SearchHPhone.value = ""
	document.all.SearchFax.value = ""
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
<%		End If %>				
<%	End If %>	
    if document.all.SearchOID.value      <> "" Or document.all.SearchTitle.value     <> "" Or _
	   document.all.SearchNameLast.value <> "" Or document.all.SearchNameFirst.value <> "" Or _
	   document.all.SearchAdd1.value     <> "" Or document.all.SearchAdd2.value      <> "" Or _
	   document.all.SearchCity.value     <> "" Or document.all.SearchState.value     <> "" Or _
	   document.all.SearchZip.value      <> "" Or document.all.SearchWphone.value    <> "" Or _
	   document.all.SearchHPhone.value   <> "" Or document.all.SearchFax.value       <> "" Then
	   UpdateStatus("<%=MSG_PROMPT%>")	
	End If
End Sub

Sub PostTo(strURL)
	
	curOID = Parent.frames("WORKAREA").GetOID
	temp = Split(curOID, "||")
	If UBound(temp) >= 0 Then 
		document.all.SearchOID.value = temp(0)
	Else		
		document.all.SearchOID.value = ""
	End If	
	FrmSearch.action = "OwnerDetails-f.asp"
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
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Owner Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

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

<form Name="FrmSearch" METHOD="GET" ACTION="OwnerSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="OID" value="<%=Request.QueryString("searchOID")%>">
<table width="100%" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
 <table>
	<tr></tr>
	<tr></tr>
	<tr nowrap>
	   <td CLASS="LABEL">Owner ID:<br><input CLASS="LABEL" size="10" tabindex=1 TYPE="TEXT" NAME="SearchOID" VALUE="<%=Request.QueryString("SearchOID")%>"></td>
	   <td CLASS="LABEL">Title:<br><input CLASS="LABEL" tabindex=2 TYPE="TEXT" NAME="SearchTitle" VALUE="<%=Request.QueryString("SearchTitle")%>"></td>
	   <td CLASS="LABEL">Name Last:<br><input  CLASS="LABEL" tabindex=3 TYPE="TEXT" NAME="SearchNamelast" VALUE="<%=Request.QueryString("SearchNameLast")%>"></td>
	   <td CLASS="LABEL">Name First:<br><input CLASS="LABEL" tabindex=4 TYPE="TEXT" NAME="SearchNameFirst" VALUE="<%=Request.QueryString("SearchNameFirst")%>"></td>
	</tr>				
	<tr>
	   <td CLASS="LABEL">Address Line1:<br><input CLASS="LABEL" tabindex=5 TYPE="TEXT" NAME="SearchAdd1" VALUE="<%=Request.QueryString("SearchAdd1")%>"></td>
	   <td CLASS="LABEL">Address Line2:<br><input CLASS="LABEL" tabindex=5 TYPE="TEXT" NAME="SearchAdd2" VALUE="<%=Request.QueryString("SearchAdd2")%>"></td>
	   <td CLASS="LABEL">City:<br><input CLASS="LABEL" tabindex=5 TYPE="TEXT" NAME="SearchCity" VALUE="<%=Request.QueryString("SearchCity")%>"></td>
	   <td CLASS="LABEL">State:<br><input size="2" CLASS="LABEL" tabindex=5 TYPE="TEXT" NAME="SearchState" VALUE="<%=Request.QueryString("SearchState")%>"></td>
	 </tr>
	 <tr>
	   <td CLASS="LABEL">Zip:<br><input CLASS="LABEL" tabindex=5 TYPE="TEXT" NAME="SearchZip" VALUE="<%=Request.QueryString("SearchZip")%>"></td>
	   <td CLASS="LABEL">Work Phone:<br><input CLASS="LABEL" tabindex=5 TYPE="TEXT" NAME="SearchWphone" VALUE="<%=Request.QueryString("SearchWPhone")%>"></td>
	   <td CLASS="LABEL">Home Phone:<br><input CLASS="LABEL" tabindex=5 TYPE="TEXT" NAME="SearchHphone" VALUE="<%=Request.QueryString("SearchHphone")%>"></td>
	   <td CLASS="LABEL">Fax:<br><input CLASS="LABEL" tabindex=5 TYPE="TEXT" NAME="SearchFax" VALUE="<%=Request.QueryString("SearchFax")%>"></td>
	   
	</tr>
 </table>
</td>			
<td VALIGN="TOP" rowspan="3">
 <table>
 <tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex=9 NAME="BtnSearch" TYPE="BUTTON" ACCESSKEY="H">Searc<u>h</u></button></td></tr>
 <tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex=10 NAME="BtnClear" ACCESSKEY="L">C<u>l</u>ear</button></td></tr>
 </table>
</td>	
</tr>
</table>

<table topmargin=0 bottommargin=0>
<tr>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" tabindex=6 NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" tabindex=7 NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" tabindex=8 NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
</tr>
</table>

</form>
</body>
</html>
