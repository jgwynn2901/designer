<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Rule Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--

Sub Window_onLoad
<% If Request.QueryString <> "" Then %>
if 0 < Document.all.tblFields.rows.length then
		call multiselect( Document.all.tblFields.rows(1))
		call Document.all.tblFields.focus()
end if
<% End if %>
End Sub

Function GetFrame()
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		idx = document.all.tblFields.rows(idx).getAttribute("FRAMEID")
	End If
	GetFrame = idx
End Function

Function DblGetFrameID(ID)
' Modal does not support double clicking on row
End Function
-->
</script>

<SCRIPT LANGUAGE="JavaScript">
<!--
function dblhighlight( objRow )
{
var qry
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	DblGetFrameID(objRow.getAttribute("FRAMEID"))
}
//-->
</SCRIPT>
</head>
<BODY  bgcolor='<%= BODYBGCOLOR %>'  leftmargin=2 topmargin=2 rightmargin=0 bottommargin=2>
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Frame Id</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>Title</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
	dim RecCount
	RecCount = -1
	WHERECLS = ""
	If Request.QueryString <> "" Then
	RecCount = 0
	LOB_CD = Request.QueryString("LOB_CD")
		Select Case Request.QueryString("SEARCHTYPE")
			Case "B"
				FRAME_ID = Request.QueryString("FRAME_ID") & "%"
				NAME = Request.QueryString("NAME") & "%"
				TITLE = Request.QueryString("TITLE") & "%"
				AHS_ID = Request.QueryString("AHS_ID") 
				CLIENTNODE_ID = Request.QueryString("ClientNode_ID") & "%"
			Case "C"
				FRAME_ID = "%" & Request.QueryString("FRAME_ID") & "%"
				NAME = "%" & Request.QueryString("NAME") & "%"
				TITLE = "%" & Request.QueryString("TITLE") & "%"
				AHS_ID =  Request.QueryString("AHS_ID") 
				CLIENTNODE_ID = "%" & Request.QueryString("ClientNode_ID") & "%"
			Case "E"
				FRAME_ID = Request.QueryString("FRAME_ID")
				NAME = Request.QueryString("NAME")
				TITLE = Request.QueryString("TITLE")
				AHS_ID = Request.QueryString("AHS_ID")
				CLIENTNODE_ID = Request.QueryString("ClientNode_ID")
		End Select

		FRAME_ID = Replace(FRAME_ID,"'","''")
		NAME = Replace(NAME,"'","''")
		TITLE = Replace(TITLE,"'","''")		

		IF Request.QueryString("ClientNode_ID") <> "" THEN
				s_SQLQuery = "SELECT DISTINCT F.* FROM FRAME F, FRAME_ORDER FO, CALLFLOW CF, ACCOUNT_CALLFLOW ACF, ACCOUNT_HIERARCHY_STEP AHS "
				WHERECLS = "(F.FRAME_ID = FO.FRAME_ID) AND (FO.CALLFLOW_ID = CF.CALLFLOW_ID) AND (CF.CALLFLOW_ID = ACF.CALLFLOW_ID) AND (ACF.ACCNT_HRCY_STEP_ID = AHS.ACCNT_HRCY_STEP_ID)"
		ELSEIf Request.QueryString("AHS_ID") <> "" OR Request.QueryString("LOB_CD") <> "" Then
				s_SQLQuery = "SELECT DISTINCT F.* FROM FRAME F, FRAME_ORDER FO, CALLFLOW CF, ACCOUNT_CALLFLOW ACF"
				WHERECLS = "(F.FRAME_ID = FO.FRAME_ID) AND (FO.CALLFLOW_ID = CF.CALLFLOW_ID) AND (CF.CALLFLOW_ID = ACF.CALLFLOW_ID)"
		ELSE
				s_SQLQuery = "SELECT F.* From FRAME F"
		END IF

		If Request.QueryString("FRAME_ID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "F.FRAME_ID LIKE '" & FRAME_ID  & "'"
		End If
		If Request.QueryString("NAME") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(F.NAME) LIKE '" & UCASE(NAME)  & "'"
		End If
		If Request.QueryString("TITLE") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(F.TITLE) LIKE '" & UCASE(TITLE) & "'"
		End If
		
		If Request.QueryString("AHS_ID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "ACF.ACCNT_HRCY_STEP_ID = " & AHS_ID 
		End If
		If Request.QueryString("ClientNode_ID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "AHS.CLIENT_NODE_ID LIKE '" & ClientNode_ID & "'"
		End If
		If Request.QueryString("LOB_CD") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "ACF.LOB_CD = '" & LOB_CD & "'"
		End If 
		
		Set RS = Server.CreateObject("ADODB.RecordSet")
		RS.MaxRecords = 30
		ConnectionString = CONNECT_STRING
		if WHERECLS <> "" then
			s_SQLQuery = s_SQLQuery & " WHERE " & WHERECLS & " ORDER BY F.NAME" 
		else
			s_SQLQuery = s_SQLQuery & " ORDER BY F.NAME" 		
		end if
		'response.Write(s_SQLQuery)
		RS.Open s_SQLQuery, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
	
		If RS.EOF AND RS.BOF Then
%>
<tr ID="FieldRow" CLASS="ResultRow" FRAMEID='X' >
	<td COLSPAN=3 NOWRAP CLASS="LABEL" ID="FRAME_ID">No frames found re-check your criteria</td>
</tr>
				
<%Else
	Do While Not RS.EOF 
	RecCount = RecCount + 1%>
<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);"  FRAMEID='<%= RS("FRAME_ID") %>' >
	<td NOWRAP CLASS="ResultCell" ID="FRAME_ID"><%=renderCell(RS("FRAME_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("NAME"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="TITLE" ><%=renderCell(RS("TITLE"))%></td>
</tr>
<%
		RS.MoveNext
		Loop
	End If
End If
%>
</tbody>
</table>
</div>
</fieldset>
<% If Request.QueryString <> "" Then %>
<SCRIPT LANGUAGE="VBScript">
if Parent.frames("TOP").document.readyState = "complete" then
	curCount = <%=RecCount%>
	if curCount = <%=MAXRECORDCOUNT%> then
		Parent.frames("TOP").UpdateStatus("<%=MSG_MAXRECORDS%>")
	elseif curCount = -1 then
		Parent.frames("TOP").UpdateStatus("<%=MSG_PROMPT%>")
	else		
		Parent.frames("TOP").UpdateStatus("Record count is <%=RecCount%>")
	end if		
end if
</SCRIPT>
<% End If %>
</body>
</html>
