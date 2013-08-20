<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires = 0
Response.Buffer = true
Response.AddHeader  "Pragma", "no-cache"
%>
<!--#include file="..\lib\tablecommon.inc"-->
<% Response.Expires = 0 %>
<html>
<head>
<title>Escalation Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		'Parent.frames("TOP").document.all.EPID.value = ""
	end if
End Sub

Function GetEPID
	GetEPID = getmultipleindex(document.all.tblFields, "EPID")
End Function

Function GetEPIDName
	GetEPIDName = getmultipleindex(document.all.tblFields, "NAME")
End Function


</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "EPID")
		return objRow.getAttribute("EPID");
	else if (whichCol == "NAME")		
		return objRow.cells("NAME").innerText;
		
}
</SCRIPT>
</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Escalation Id</div></td>
			<td class="thd"><div id><nobr>AHS ID</div></td>
			<td class="thd"><div id><nobr>Lob</div></td>
			<td class="thd"><div id><nobr>Description</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
	dim RecCount
	RecCount = -1
	
	If Request.QueryString("SEARCHTYPE") <> "" Then
	RecCount = 0
		Select Case Request.QueryString("SEARCHTYPE")
			Case "B"
				EPID = Request.QueryString("SearchEPID") & "%"
				LOB_CD = Request.QueryString("SearchLOB_CD") & "%"
				ACCNT_HRCY_STEP_ID = Request.QueryString("SearchACCNT_HRCY_STEP_ID") & "%"
			Case "C"
				EPID = "%" & Request.QueryString("SearchEPID") & "%"
				LOB_CD = "%" & Request.QueryString("SearchLOB_CD") & "%"
				ACCNT_HRCY_STEP_ID = "%" & Request.QueryString("SearchACCNT_HRCY_STEP_ID") & "%"
			Case "E"
				EPID = Request.QueryString("SearchEPID")
				LOB_CD = Request.QueryString("SearchLOB_CD")
				ACCNT_HRCY_STEP_ID = Request.QueryString("SearchACCNT_HRCY_STEP_ID")
		End Select
		EPID = Replace(EPID, "'", "''")
		LOB_CD = Replace(LOB_CD, "'", "''")
		ACCNT_HRCY_STEP_ID = Replace(ACCNT_HRCY_STEP_ID, "'", "''")
		
		If Request.QueryString("SearchACCNT_HRCY_STEP_ID") <> "" Then
			WHERECLS = WHERECLS & "UPPER(ACCNT_HRCY_STEP_ID) LIKE '" & UCASE(ACCNT_HRCY_STEP_ID)  & "'"
		End If
		If Request.QueryString("SearchEPID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "ESCALATION_PLAN_ID LIKE '" & EPID & "'"
		End If
		If Request.QueryString("SearchLOB_CD") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(LOB_CD) LIKE '" & UCASE(LOB_CD) & "'"
		End If

			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT ACCNT_HRCY_STEP_ID,ESCALATION_PLAN_ID, LOB_CD, DESCRIPTION FROM ESCALATION_PLAN "
			
			If WHERECLS <> "" Then
				SQLST = SQLST & "WHERE " & WHERECLS
			End If
			
			SQLST = SQLST & " ORDER BY ACCNT_HRCY_STEP_ID, LOB_CD" 
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" EPID='' >
	<td COLSPAN=10 NOWRAP CLASS="ResultCell">No Escalations found re-check your criteria</td>
</tr>
	
	<%		Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  EPID='<%=RS("ESCALATION_PLAN_ID")%>'>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ESCALATION_PLAN_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("ACCNT_HRCY_STEP_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("LOB_CD"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("DESCRIPTION"))%></td>
	</tr>

<%
				RS.MoveNext
				Loop
				
			End If
			RS.Close
			Set RS = Nothing
			Conn.Close
			Set Conn = Nothing
	End If
%>

</tbody>
</table>
</div>
</fieldset>
<SCRIPT LANGUAGE="VBScript">
<%	If RecCount >= 0 Then %>
if Parent.frames("TOP").document.readyState = "complete" then
	curCount = <%=RecCount%>
	if curCount = <%=MAXRECORDCOUNT%> then
		Parent.frames("TOP").UpdateStatus("<%=MSG_MAXRECORDS%>")
	else		
		Parent.frames("TOP").UpdateStatus("Record count is <%=RecCount%>")
	end if		
end if
<%	End If %>
</SCRIPT>
</body>
</html>
