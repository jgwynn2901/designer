<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache" %>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<%
Sub GetStatusRptHTML(ShowMostRecent, GrpCount)

	On Error Resume Next

	Dim objStatusRpt
	Set objStatusRpt = Session("StatusRptRS")

	If IsObject(objStatusRpt) Then

		objStatusRpt.MoveFirst

		If Err.Number <> 0 Then
			Exit Sub
		End If

		Do While Not objStatusRpt.EOF

			If Session("StatusRptMsgNextGrpID") = objStatusRpt("GrpID") Then
%>
<tr ID="FieldRow" CLASS="ResultRow">
	<td NOWRAP CLASS="LABEL" ID="Severity"><%=objStatusRpt("Severity")%></td>
	<td NOWRAP CLASS="LABEL" ID="Message"><%=objStatusRpt("Message")%></td>
	<td NOWRAP CLASS="LABEL" ID="SourceRowID"><%=objStatusRpt("SourceRowID")%></td>
	<td NOWRAP CLASS="LABEL" ID="SourceRowDesc" ><%=objStatusRpt("SourceRowDesc")%></td>
	<td NOWRAP CLASS="LABEL" ID="SourceTable"><%=objStatusRpt("SourceTable")%></td>
	<td NOWRAP CLASS="LABEL" ID="SourceField"><%=objStatusRpt("SourceField")%></td>
</tr>
<%
		Else
			objStatusRpt.Delete
		End If

		objStatusRpt.MoveNext
		Loop
	End If
End Sub
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Status Report - FNS Net Designer</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub BtnClose_OnClick
	window.close
End Sub
</script>
</head>
<body BGCOLOR="#d6cfbd">
<div style="position:absolute;top:4;left:10";width:'100%'>
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Status Report</TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<fieldset align="center" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699">
<div align="center" style="display:block;height:254;width:'100%';overflow:scroll">
<table cellPadding="2" cellSpacing="0" frame="void" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Severity</div></td>
			<td class="thd"><div id><nobr>Message</div></td>
			<td class="thd"><div id><nobr>Src Row ID</div></td>
			<td class="thd"><div id><nobr>Src Row Desc</div></td>
			<td class="thd"><div id><nobr>Src Table</div></td>
			<td class="thd"><div id><nobr>Src Field</div></td>

		</tr>
	</thead>
	<tbody ID="TableRows">

<%
Dim bShowMostRecent, nGrpCount

If Request.QueryString("ShowMostRecent") = "True" Then
	bShowMostRecent = True
Else
	bShowMostRecent = False
End If

If Request.QueryString("GrpCount") = "True" Then
	nGrpCount = True
Else
	nGrpCount = False
End If

Call GetStatusRptHTML(bShowMostRecent, nGrpCount)
%>
</tbody>
</table>
</div>
</fieldset>
<br>
<br>
<div align="center"><button ID=Details>Details...</button>&nbsp<button ID=BtnClose>Close</button></div>
</div>
</body>
</html>
