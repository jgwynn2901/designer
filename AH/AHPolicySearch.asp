<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE></TITLE>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub ClearSearch()
	document.all.POLICY_NUMBER.value = ""
	document.all.POLICY_DESC.value = ""
	document.all.LOB_CD.value = ""
	document.all.EFFECTIVE_DATE.value = ""
	document.all.EXPIRATION_DATE.value = ""
End Sub

Sub ExeSearch()
errstr = ""
	If document.all.POLICY_NUMBER.value = "" AND document.all.POLICY_DESC.value = "" AND document.all.EXPIRATION_DATE.value = "" AND document.all.EFFECTIVE_DATE.value = "" AND document.all.LOB_CD.value = "" Then
		errstr = errstr & "Please enter search criteria!" & VBCRLF
	End If
	If Not CheckDate(document.all.EFFECTIVE_DATE.value) AND document.all.EFFECTIVE_DATE.value <> "" Then
		errstr = errstr & "Effective Date has an incorrect format. Format as MM/DD/YYYY" & VBCRLF
	End If
	If Not CheckDate(document.all.EXPIRATION_DATE.value) AND document.all.EXPIRATION_DATE.value <> "" Then
		errstr = errstr & "Expiration Date has an incorrect format. Format as MM/DD/YYYY" & VBCRLF
	End If
	If errstr = "" Then
		SPANSTATUS.innerHTML = "<%= MSG_SEARCH %>"
		FrmSearch.submit
	Else
		MsgBox errstr, 0 , "FNSNetDesigner"
	End If
End Sub

Sub window_onload

End Sub

Sub BtnSearch_OnClick
	Call ExeSearch()
End Sub

Sub BtnClear_OnCLick
	Call ClearSearch()
End Sub

Function CheckDate( InDate )
	If Not IsDate(InDate) Then
		CheckDate = false
		Exit Function
	End If
	If Len(InDate) <> 10 OR Mid(InDate,1,2) > 12 Then
		CheckDate = false
		Exit Function
	End If
	CheckDate = true
End Function
-->
</SCRIPT>
</HEAD>
<BODY  topmargin=0 leftmargin=0 bgcolor='<%= BODYBGCOLOR %>'>
<FORM Name="FrmSearch" TARGET="WORKAREA" METHOD=POST ACTION="AHPolicySearchResults.asp">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Policy Search</TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME=AT_AHSID VALUE="<%= Request.QueryString("AHSID") %>">

<TABLE  cellspacing=0 cellpadding=0>
<TR>
<TD CLASS=LABEL><img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" Title=""></TD>
<TD CLASS=LABEL><SPAN ID=SPANSTATUS STYLE="COLOR:#006699" CLASS=LABEL>: Ready</SPAN></TD>
</TR>
</TABLE>


<TABLE WIDTH="100%"><TR><TD ALIGN=LEFT>
<TABLE>
<TR>
<TD CLASS=LABEL>Policy number:<BR><INPUT TYPE=TEXT NAME=POLICY_NUMBER CLASS=LABEL></TD>
<TD CLASS=LABEL>Effective Date:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=EFFECTIVE_DATE MAXLENGTH=10 SIZE=20></TD>
<TD CLASS=LABEL>Expiration Date:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=EXPIRATION_DATE MAXLENGTH=10 SIZE=20></TD>
</TR>
<TR>
<TD CLASS=LABEL COLSPAN=2>Policy Description:<BR><INPUT TYPE=TEXT NAME=POLICY_DESC SIZE=45 CLASS=LABEL STYLE="TEXT-TRANSFORM:UPPERCASE"></TD>
<TD CLASS=LABEL>LOB:<BR>
<SELECT NAME=LOB_CD CLASS=LABEL>
<OPTION VALUE="">
<%
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLST = ""
	SQLST = SQLST & "SELECT LOB_CD,LOB_NAME FROM LOB WHERE LOB_CD IS NOT NULL"
	Set RS = Conn.Execute(SQLST)
Do While Not RS.EOF
%>
<OPTION VALUE="<%= RS("LOB_CD") %>"><%= RS("LOB_NAME") %>
<%
RS.MoveNext
Loop
RS.CLose
%>
</SELECT></TD>
</TR>
</TABLE>
<TABLE CELLPADDING=0 CELLSPACING=0>
<tr>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL" CHECKED>Begins With</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
<td width=75>
<TD ALIGN=RIGHT CLASS=LABEL> Direction:
<input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchDirection" VALUE="UP" CLASS="LABEL">Up
<input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchDirection" VALUE="Down" CLASS="LABEL" CHECKED>Down
</td>
</tr>
</TABLE>
</TD><TD ALIGN=RIGHT VALIGN=TOP>
<TABLE>
<TR>
<td ALIGN=RIGHT CLASS="LABEL"><button CLASS="StdButton" NAME="BtnSearch" ACCESSKEY="C">Sear<u>c</u>h</button></td>
</TR>
<TR>
<td ALIGN=RIGHT CLASS="LABEL"><button CLASS="StdButton" NAME="BtnClear" ACCESSKEY="L">C<U>l</U>ear</button></td>
</TR>
</TABLE>
</TD></TR></TABLE>
</FORM>
</BODY>
</HTML>
