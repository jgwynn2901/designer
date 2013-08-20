<%
'***************************************************************
'implements the search tab for Department.
'
'$History: DepartmentSearch.asp $ 
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 1/25/07    Time: 9:09a
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/DEPARTMENT
'* Moved the Department interface to Account Related and created a new
'* permission FNSD_DEPARTMENT based on Doug's recommondation.
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 1/24/07    Time: 1:39p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/Policy
'* Added Department Interface due to ESIS Project.  It allows User to
'* create Department record attached to the AHSID in PROD Designer. The
'* permission used is the same as for Branch.
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 1/24/07    Time: 12:10p
'* Created in $/FNS_DESIGNER/Source/Designer/Policy
'* Added Department Interface due to the ESIS Project.  It allows user to
'* attach AHSID to the department record.  Also, it allows user to delete,
'* create a new record and Edit an record in PROD Designer.  Permission
'* setup is the same as for Branch.  
'* 

%>
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
<title>Department Search</title>
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
	document.all.SearchDeptName.value = ""
	document.all.SearchAHSID.value = ""
	document.all.SearchDeptCode.value = ""
End Sub

Sub BtnSearch_onclick()
	document.all.SpanStatus.innerHTML = "<%=MSG_SEARCH%>"
	FrmSearch.submit
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

	End If %>	

	If document.all.SearchDeptCode.value <> "" Or document.all.SearchAHSID.value <> "" or document.all.SearchDeptName.value <> ""Then
		UpdateStatus("<%=MSG_PROMPT%>")	
	End If

End Sub

Sub PostTo(strURL)
	curDEPTID = Parent.frames("WORKAREA").GetDEPTID
	temp = Split(curDEPTID, "||")
	If UBound(temp) >= 0 Then 
		document.all.DEPTID.value = temp(0)
	Else		
		document.all.DEPTID.value = ""
	End If
	FrmSearch.action = "DEPARTMENTDetails-f.asp"
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
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Department Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
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

<form Name="FrmSearch" METHOD="GET" ACTION="departmentSearchResults.asp" TARGET="WORKAREA">
<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
<input type="hidden" NAME="DEPTID" value="<%=Request.QueryString("DEPTID")%>">
<table width=100% CELLPADDING="0" CELLSPACING="0">
<tr>
<td>
	<table CLASS="LABEL" style="width:300" align=left>
	<tr>
	<tr>
	<tr>
	<td CLASS="LABEL">Department Name:<br><input size=25 tabindex=1 CLASS="LABEL" TYPE="TEXT" NAME="SearchDeptName"  VALUE="<%=Request.QueryString("SearchDeptName")%>"></td>
	<td CLASS="LABEL">Department Code:<br><input size=25 tabindex=2 CLASS="LABEL" TYPE="TEXT" NAME="SearchDeptCode" VALUE="<%=Request.QueryString("SearchDeptCode")%>"></td>
	<td CLASS="LABEL">AHS ID:<br><input size=25 tabindex=2 CLASS="LABEL" TYPE="TEXT" NAME="SearchAHSID" VALUE="<%=Request.QueryString("SearchAHSID")%>" onfocus="enable_exact()" onBlur="enable_begin()" ></td>
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
