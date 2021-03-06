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
<title>User Search</title>
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
	document.all.SearchUID.value = ""
	document.all.SearchName.value = ""
	document.all.SearchSite.value = ""
	document.all.SearchActive.value = ""
	document.all.SearchFName.value = ""
	document.all.SearchLName.value = ""
	document.all.SearchAHSID.value = ""
	document.all.SearchEMAIL.value = ""
End Sub

Sub BtnSearch_onclick()
	'If document.all.SearchUID.value = "" And document.all.SearchName.value = "" And _
	'document.all.SearchSite.value = ""  Then
		'MsgBox "Please enter search criteria", 0, "FNSNetDesigner"
	'Else
		document.all.SpanStatus.innerHTML = "<%=MSG_SEARCH%>"
		FrmSearch.submit
	'End If
End Sub

Sub window_onload
	'document.all.SearchName.focus ' Timing Problem
	parent.parent.disableTab("Details")
	parent.parent.disableTab("Groups")
	parent.parent.disableTab("Permissions")
	parent.parent.disableTab("Accounts")
	parent.parent.disableTab("Locations")
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

		If CStr(Request.QueryString("SearchSite")) <> "" Then	%>
			SelectOption document.all.SearchSite,"<%=CStr(Request.QueryString("SearchSite"))%>"
<%		End If 

		If CStr(Request.QueryString("SearchActive")) <> "" Then	%>
			SelectOption document.all.SearchActive,"<%=CStr(Request.QueryString("SearchActive"))%>"
<%		End If 

	End If %>	

	If document.all.SearchUID.value <> "" Or document.all.SearchName.value <> "" Or _
		document.all.SearchSite.value <> "" Or document.all.SearchActive.value <> "" Then
		UpdateStatus("<%=MSG_PROMPT%>")	
	End If
End Sub

Sub PostTo(strURL)
	document.all.UID.value = Parent.frames("WORKAREA").GetUID
	FrmSearch.action = strURL
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
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;� User Search</td>
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

<form Name="FrmSearch" METHOD="GET" ACTION="UserSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="UID" value="<%=Request.QueryString("UID")%>">
<table width="100%" CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
 <table>
	<tr></tr>
	<tr></tr>
	<tr nowrap>
	<td CLASS="LABEL">Name:<br>
	   <input CLASS="LABEL" 
	          size="50" 
	          tabindex=1 
	          TYPE="TEXT" 
	          NAME="SearchName"
	          VALUE="<%=Request.QueryString("SearchName")%>"></td>
			  
			  <td CLASS="LABEL">First Name:<br>
	   <input CLASS="LABEL" 
	          tabindex=2 
	          TYPE="TEXT" 
	          NAME="SearchFName"
	          VALUE="<%=Request.QueryString("SearchFName")%>"></td>
    <td CLASS="LABEL">Last Name:<br>
	   <input CLASS="LABEL" 
	          tabindex=3 
	          TYPE="TEXT" 
	          NAME="SearchLName"
	          VALUE="<%=Request.QueryString("SearchLName")%>"></td> 		
	<td CLASS="LABEL">Site:<br>
	<SELECT NAME="SearchSite" 
	        tabindex=4 
	        CLASS="LABEL"><%=GetControlDataHTML("SITE","SITE_ID","NAME","",true)%></SELECT></td>
	<td CLASS="LABEL">Active:<br>
	<SELECT NAME="SearchActive" 
	        tabindex=5 
	        CLASS="LABEL">
	        <OPTION VALUE='' SELECTED></OPTION>
	        <OPTION VALUE='Y'>YES</OPTION>
	        <OPTION VALUE='N'>NO</OPTION>
	 </SELECT></td>
	 	  
	</tr>				
	<tr>
	<td CLASS="LABEL">User ID:<br>
	     <input CLASS="LABEL" 
	     tabindex=6 
	     TYPE="TEXT" 
	     NAME="SearchUID" 
	     VALUE="<%=Request.QueryString("SearchUID")%>">
	 </td>
	<td CLASS="LABEL">AHSID:<br>
	     <input CLASS="LABEL" 
	     tabindex=7 
	     TYPE="TEXT" 
	     NAME="SearchAHSID" 
	     VALUE="<%=Request.QueryString("SearchAHSID")%>" ID="Text1">
	 </td>
	 
	 <td CLASS="LABEL">EMAIL:<br>
	     <input CLASS="LABEL" 
	     tabindex=8 
	     TYPE="TEXT" 
	     NAME="SearchEMAIL" 
	     VALUE="<%=Request.QueryString("SearchEMAIL")%>" ID="Text1">
	 </td>
	</tr>
 </table>
</td>			
<td VALIGN="TOP" rowspan="3">
 <table>
 <tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex=12 NAME="BtnSearch" TYPE="BUTTON" ACCESSKEY="H">Searc<u>h</u></button></td></tr>
 <tr><td CLASS="LABEL"><button CLASS="StdButton" tabindex=13 NAME="BtnClear" ACCESSKEY="L">C<u>l</u>ear</button></td></tr>
 </table>
</td>	
</tr>
</table>

<table topmargin=0 bottommargin=0>
<tr>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" tabindex=9 NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" tabindex=10 NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" tabindex=11 NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
</tr>
</table>

</form>
</body>
</html>
