<!--#include file="..\lib\common.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Attribute Search</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--

Sub BtnClear_onclick()
	document.all.SearchAID.value = ""
	document.all.SearchName.value = ""
	document.all.SearchCaption.value = ""
	document.all.SearchDescription.value = ""
	document.all.SearchHelpString.value = ""
End Sub

Sub BtnSearch_onclick()
	If document.all.SearchAID.value = "" And document.all.SearchName.value = "" And _
	document.all.SearchCaption.value = "" And document.all.SearchDescription.value = "" And _
	document.all.SearchHelpString.value = "" Then
			MsgBox "Please enter search criteria", 0, "FNSNetDesigner"
	Else
		FrmSearch.submit
	End If
End Sub

Sub window_onload
	'document.all.SearchName.focus ' Timing Problem
	document.all.SearchType(0).checked = True
	
<%	If Request.QueryString <> "" Then %>
<%		If CStr(Request.QueryString("SearchType")) = "B" Then	%>
			document.all.SearchType(0).checked = True
<%		ElseIf CStr(Request.QueryString("SearchType")) = "C" Then	%>
			document.all.SearchType(1).checked = True
<%		ElseIf CStr(Request.QueryString("SearchType")) = "E" Then	%>
			document.all.SearchType(2).checked = True
<%		End If %>				
	
<%	End If %>	

<%	'AID will only be populated if we are returning from the details screen
	'so we need to submit the form to simulate the tabs appearance before the
	'detail tab was loaded
	If CStr(Request.QueryString("AID")) <> "" And CStr(Request.QueryString("AID")) <> "NEW" Then 
%>
	FrmSearch.submit
<%	End If %>	

End Sub

Sub PostTo
	document.all.AID.value = Parent.frames("WORKAREA").GetAID
	FrmSearch.action = "AttributeDetails-f.asp"
	FrmSearch.method = "GET"	
	FrmSearch.target = "_parent"	
	FrmSearch.submit
End Sub

-->
</script>

</head>
<BODY  topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="#d6cfbd">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Attribute Search</TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
<form Name="FrmSearch" METHOD="GET" ACTION="AttributeSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="AID" value="<%=Request.QueryString("AID")%>">



<table width=100%>


<tr>
<td CLASS="LABEL">Attribute ID:<br><input CLASS="LABEL" TYPE="TEXT" NAME="SearchAID" VALUE="<%=Request.QueryString("SearchAID")%>"></td>
<td CLASS="LABEL">Name:<br><input CLASS="LABEL" TYPE="TEXT" NAME="SearchName" VALUE="<%=Request.QueryString("SearchName")%>"></td>
<td CLASS="LABEL">Caption:<br><input CLASS="LABEL" TYPE="TEXT" NAME="SearchCaption" VALUE="<%=Request.QueryString("SearchCaption")%>"></td>
<td VALIGN=TOP rowspan="3" >
 <TABLE>
 <TR><TD CLASS=LABEL><BUTTON CLASS=StdButton NAME=BtnSearch TYPE="BUTTON" ACCESSKEY="H">Searc<U>h</U></BUTTON></TD></TR>
 <TR><TD CLASS=LABEL><BUTTON CLASS=StdButton NAME=BtnClear ACCESSKEY="L">C<U>l</U>ear</BUTTON></TD></TR>
 </TABLE>
</td>	
</tr>
<tr>
<td CLASS="LABEL">Description:<br><input  CLASS="LABEL" TYPE="TEXT" NAME="SearchDescription" VALUE="<%=Request.QueryString("SearchDescription")%>"></td>
<td colspan=2 CLASS="LABEL">Help String:<br><input size="48" CLASS="LABEL" TYPE="TEXT" NAME="SearchHelpString" VALUE="<%=Request.QueryString("SearchHelpString")%>"></td>
</tr>
<tr>
<tr>
<tr>
</table>

<table>
<tr>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
</tr>
</table>

</form>
</body>
</html>
