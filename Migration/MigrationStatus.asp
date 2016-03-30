<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\commonError.inc"-->
<!--#include file="..\lib\security.inc"-->
<% Response.Expires = 0 %>
<%
dim oConn, cSQL, cJobID, oRS
Dim cDeleteMsg, cStatMsg
dim cORACLETime
dim cStart, cEnd, cStatus, cStyle

PREV_ID = "-1"
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open CONNECT_STRING

set oRS = oConn.Execute( "Select to_char(sysdate, 'HH:MI:SS PM') from dual" )
cORACLETime = oRS.Fields(0).Value
set oRS = nothing 
	
Function ShowDetail(ID)
	Select Case ID
		Case "0"
			ShowDetail = "Pending"
		Case "1"
			ShowDetail = "Ready"
		Case "2"
			ShowDetail = "Started"
		Case "3"
			ShowDetail = "Completed"
		Case "4"
			ShowDetail = "<B>Failed</B>"
	End Select
End Function

Function MyClass(ID, cStart, cEnd)
	Select Case ID
		Case "0"
			MyClass = "ResultCell"
		Case "1"
			MyClass = "ResultCell"
		Case "2"
			MyClass = "ResultCell"
		Case "3"
			MyClass = "ResultCellComplete"
			if datediff("n", cdate(cStart), cdate(cEnd)) > 30 then
				MyClass = "ResultCellLongTime"
			end if
		Case "4"
			MyClass = "ResultCellFailed"
	End Select
End Function

cJobID = Request.QueryString("JOBID")
If Request.QueryString("DELETE") = "TRUE" AND cJobID <> ""Then
	' check if job is running
	cSQL = "SELECT START_TIME FROM MIGRATION_JOB WHERE JOB_ID=" & cJobID
	Set oRS = Server.CreateObject("ADODB.Recordset")
	oRS.Open cSQL, oConn, adOpenStatic, adLockReadOnly   
	if oRS.Fields("START_TIME").Value <> 0 then
		cDeleteMsg = "Cannot delete. Job already started!"
	else
		cSQL = "DELETE FROM MIGRATION_DETAIL WHERE JOB_ID=" & cJobID
		oConn.Execute(cSQL)
		cSQL = "DELETE FROM MIGRATION_JOB WHERE JOB_ID=" & cJobID
		oConn.Execute(cSQL)
	end if
	oRS.Close 
	set oRS = nothing
End If
cSQL = "SELECT * FROM MIGRATION_JOB WHERE ROWNUM < " & MAXRECORDCOUNT + 1 & " ORDER BY JOB_ID DESC"
cSQL = "SELECT * FROM MIGRATION_JOB WHERE SCHEDULED_START>sysdate-180 and SCHEDULED_START is not null ORDER BY SCHEDULED_START DESC"
Set oRS = oConn.Execute(cSQL)
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow){

multiselect(objRow)
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
		//return objRow.getAttribute("JOBID")
		document.frames("Status").location.href = "MigrationStatusDetails.asp?JOBID=" + objRow.getAttribute("JOBID")
}
</SCRIPT>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub BtnDelete_onclick
	GetIndex = getselectedindex(document.all.tblFields)
	GetJOBID = document.all.tblFields.Rows(GetIndex).GetAttribute("JOBID")
	If GetJOBID <> "" Then
		lret = msgbox ("Are you sure you wish to delete Job ID: " & GetJOBID, 1, "FNSDesigner")
		If lret = 1 Then
			self.location.href = "MigrationStatus.asp?DELETE=TRUE&JOBID=" & GetJOBID
		End If
	Else
		msgbox "Please select a migration job.",0,"FNSDesigner"
	End If
End Sub

Sub Btnrefresh_onclick
	self.location.href = "MigrationStatus.asp"
End Sub

Sub window_onload
	lret = window.setInterval ("Refresh()",  60000)
	<% If Request.QueryString("PREV_ID") <> "-1" AND Request.QueryString("PREV_ID") <> "" Then %>
		Call dblhighlight(document.all.tblFields.Rows(<%= Request.QueryString("INDEX") %>))
	<% End If 
	 IF Not(cDeleteMsg = "") THEN
			s_Quote = """"
			s_MsgboxTitle = s_Quote & "FNSDesigner DataMigration" & s_Quote
			Response.Write "Msgbox " & s_Quote & cDeleteMsg & s_Quote & ", 0, " & s_MsgboxTitle
	 END IF 
	%>
End Sub

Function Refresh()
If document.all.AUTOREFRESH.checked = true Then
	GetIndex = getselectedindex(document.all.tblFields)
	If GetIndex = "-1" Then
		GetJOBID = "-1"
	Else
		GetJOBID = document.all.tblFields.Rows(GetIndex).GetAttribute("JOBID")
	End If
	self.location.href = "MigrationStatus.asp?PREV_ID=" & GetJOBID & "&INDEX=" & GetIndex
End If
End Function	
-->
</SCRIPT>
</HEAD>
<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<SPAN ID=TimeStamp STYLE="Position:Absolute;Left:'80%';Top:1" CLASS=LABEL><NOBR>As of <%=cORACLETime%></SPAN>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Migration Status Job</td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'45%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Job ID</div></td>
			<td class="thd"><div id><nobr>A.H.S. ID</div></td>
			<td class="thd"><div id><nobr>LOB</div></td>
			<td class="thd"><div id><nobr>Status Code</div></td>
			<td class="thd"><div id><nobr>Start Time</div></td>
			<td class="thd"><div id><nobr>End Time</div></td>
			<td class="thd"><div id><nobr>Scheduled Start</div></td>
			<td class="thd"><div id><nobr>Status Msg</div></td>
			<td class="thd"><div id><nobr>All R.P.</div></td>
			<td class="thd"><div id><nobr>All O.D.</div></td>
			<td class="thd"><div id><nobr>Reference ID</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<% 
If Not oRS.EOF AND Not oRS.BOF Then
Do While not oRS.EOF
	cStart = oRS("START_TIME")
	cEnd = oRS("END_TIME")
	cStatus = oRS("STATUS_CD")
	cStyle = MyClass(cStatus, cStart, cEnd)
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:dblhighlight(this);" OnDblClick="Javascript:dblhighlight(this);"  JOBID='<%=oRS("JOB_ID")%>'>
	<td NOWRAP CLASS="<%=cStyle%>"><%=renderCell(oRS("JOB_ID"))%></td>
	<td NOWRAP CLASS="<%=cStyle%>"><%=renderCell(oRS("ACCNT_HRCY_STEP_ID"))%></td>
	<td NOWRAP CLASS="<%=cStyle%>"><%=renderCell(oRS("LOB_CD"))%></td>
	<td NOWRAP CLASS="<%=cStyle%>">(<%=renderCell(cStatus)%>)&nbsp;<%= ShowDetail(oRS("STATUS_CD"))%></td>
	<td NOWRAP CLASS="<%=cStyle%>"><%=renderCell(cStart)%></td>
	<td NOWRAP CLASS="<%=cStyle%>"><%=renderCell(cEnd)%></td>
	<td NOWRAP CLASS="<%=cStyle%>"><%=renderCell(oRS("SCHEDULED_START"))%></td>
	<td NOWRAP CLASS="<%=cStyle%>"><%=renderCell(oRS("STATUS_MSG"))%></td>
	<td NOWRAP CLASS="<%=cStyle%>"><%=renderCell(oRS("MOVE_ALL_ROUTING_PLANS"))%></td>
	<td NOWRAP CLASS="<%=cStyle%>"><%=renderCell(oRS("MOVE_ALL_OUTPUT_DEFS"))%></td>
	<td NOWRAP CLASS="<%=cStyle%>"><%=renderCell(oRS("REFERENCE_ID"))%></td>
	</tr>
<%
oRS.MoveNext
Loop
oRS.CLose
set oRS = nothing
oConn.Close 
set oConn = nothing
Else
%>
<tr ID="FieldRow" CLASS="ResultRow" JOBID='' >
	<td COLSPAN=7 NOWRAP CLASS="ResultCell">No migration jobs found.</td>
</tr>
<% End If %>
</TABLE>
</DIV>
</FIELDSET>
<TABLE>
<TR>
<TD CLASS=LABEL><BUTTON NAME=BtnDelete CLASS=STDBUTTON ACCESSKEY="D"><U>D</U>elete</BUTTON></TD>
<TD CLASS=LABEL><BUTTON NAME=Btnrefresh CLASS=STDBUTTON ACCESSKEY="R"><U>R</U>efresh</BUTTON></TD>
<TD CLASS=LABEL WIDTH="60%" ALIGN=RIGHT><INPUT TYPE=CHECKBOX CLASS=LABEL NAME=AUTOREFRESH CHECKED>Auto Refresh?</TD>
</TR>
</TABLE>
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Migration Job Details </td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="100%"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'40%';width:'100%'">
<iframe FRAMEBORDER="0" ID="Status" WIDTH="100%" HEIGHT="100%" SRC="MigrationStatusDetails.asp?JOBID=<%= PREV_ID %>" scrolling=no>
</iframe>
</FIELDSET>
</BODY>
</HTML>
