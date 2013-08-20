<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->
<!--#include file="..\lib\CheckSharedRule.inc"-->
<!--#include file="..\lib\CheckSharedAttribute.inc"-->
<!--#include file="..\lib\CheckSharedRoutingAddress.inc"-->
<!--#include file="..\lib\CheckSharedOffice.inc"-->
<!--#include file="..\lib\CheckSharedAgent.inc"-->
<!--#include file="..\lib\CheckSharedFrame.inc"-->
<!--#include file="..\lib\CheckSharedCallFlow.inc"-->
<!--#include file="..\lib\CheckSharedLUType.inc"-->
<!--#include file="..\lib\CheckSharedCarrier.inc"-->
<!--#include file="..\lib\CheckSharedOD.inc"-->
<!--#include file="..\lib\RefCountRptInc.asp"-->
<%
Dim SharedCount, ID

ID = Request.QueryString("ID")	
If Request.QueryString("CheckSharedAttribute") = "True" Then
	SharedCount = CheckSharedAttribute(ID, False, False, 0, True, True, 1)
	If (SharedCount = 0) Then
		ClearRefCountRpt
	End If
End If

If Request.QueryString("CheckSharedRoutingAddress") = "True" Then
	SharedCount = CheckSharedRoutingAddress(ID, False, False, 0, True, True, 1)
	If (SharedCount = 0) Then
		ClearRefCountRpt
	End If
End If

If Request.QueryString("CheckSharedOffice") = "True" Then
	SharedCount = CheckSharedOffice(ID, False, False, 0, True, True, 1)
	If (SharedCount = 0) Then
		ClearRefCountRpt
	End If
End If

If Request.QueryString("CheckSharedAgent") = "True" Then
	SharedCount = CheckSharedAgent(ID, False, False, 0, True, True, 1)
	If (SharedCount = 0) Then
		ClearRefCountRpt
	End If
End If

If Request.QueryString("CheckSharedFrame") = "True" Then
	SharedCount = CheckSharedFrame(ID, False, False, 0, True, True, 1)
	If (SharedCount = 0) Then
		ClearRefCountRpt
	End If
End If

If Request.QueryString("CheckSharedRule") = "True" Then
	SharedCount = CheckSharedRule(ID, False, False, 0, True, True, 1)
	If (SharedCount = 0) Then
		ClearRefCountRpt
	End If
End If

If Request.QueryString("CheckSharedCallFlow") = "True" Then
	SharedCount = CheckSharedCallFlow(ID, False, False, 0, True, True, 1)
	If (SharedCount = 0) Then
		ClearRefCountRpt
	End If
End If

If Request.QueryString("CheckSharedLUType") = "True" Then
	SharedCount = CheckSharedLUType(ID, False, False, 0, True, True, 1)
	If (SharedCount = 0) Then
		ClearRefCountRpt
	End If
End If

If Request.QueryString("CheckSharedCarrier") = "True" Then
	SharedCount = CheckSharedCarrier(ID, False, False, 0, True, True, 1)
	If (SharedCount = 0) Then
		ClearRefCountRpt
	End If
End If


If Request.QueryString("CheckSharedOD") = "True" Then
	SharedCount = CheckSharedOD(ID, False, False, 0, True, True, 1)
	If (SharedCount = 0) Then
		ClearRefCountRpt
	End If
End If


On Error Resume Next

%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Reference Count Report - FNS Net Designer</title>
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
<TR><TD CLASS=GrpLabel WIDTH="180" HEIGHT=10><NOBR>&nbsp&#187 Reference Count Report (<%=SharedCount%>)</TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
<%

	Dim objRefCountRpt
	Set objRefCountRpt = Session("RefCountRptRS")
	
	If IsObject(objRefCountRpt) Then

		Do While Not objRefCountRpt.EOF

			If Session("RefCountRptMsgNextGrpID") = objRefCountRpt("GrpID") Then
%>
<SPAN CLASS=LABEL><%=objRefCountRpt("Message")%></SPAN>
<%
			Else
				objRefCountRpt.Delete
			End If

			objRefCountRpt.MoveNext
		Loop

	End If
%>
<fieldset align="center" STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699">
<div align="center" style="display:block;height:254;width:'100%';overflow:scroll">
<table cellPadding="2" cellSpacing="0" frame="void" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
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


objRefCountRpt.MoveFirst

Do While Not objRefCountRpt.EOF

	If Session("RefCountRptMsgNextGrpID") = objRefCountRpt("GrpID") Then
%>
<tr ID="FieldRow" CLASS="ResultRow">
	<td NOWRAP CLASS="LABEL" ID="SourceRowID"><%=objRefCountRpt("SourceRowID")%></td>
	<td NOWRAP CLASS="LABEL" ID="SourceRowDesc" ><%=objRefCountRpt("SourceRowDesc")%></td>
	<td NOWRAP CLASS="LABEL" ID="SourceTable"><%=objRefCountRpt("SourceTable")%></td>
	<td NOWRAP CLASS="LABEL" ID="SourceField"><%=objRefCountRpt("SourceField")%></td>
</tr>
<%
	Else
		objRefCountRpt.Delete
	End If

	objRefCountRpt.MoveNext
Loop

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
