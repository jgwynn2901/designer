<%
'***************************************************************
'display the results of a Department query in table format.
'
'$History: DepartmentSearchResults.asp $ 
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

%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires = 0
Response.Buffer = true
Response.AddHeader  "Pragma", "no-cache"
%>

<!--#include file="..\lib\tablecommon.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Department Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.DEPTID.value = ""
	end if
End Sub

Function GetDEPTID
	GetDEPTID = getmultipleindex(document.all.tblFields, "DEPTID")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "DEPTID")
		return objRow.getAttribute("DEPTID");
}
</SCRIPT>
</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0"  rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Department ID</div></td>
			<td class="thd"><div id><nobr>Department Name</div></td>
			<td class="thd"><div id><nobr>Department Code</div></td>
			<td class="thd"><div id"Div1"><nobr>AHS ID</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
	dim RecCount
	RecCount = -1
	WHERECLS = ""
	If Request.QueryString <> "" Then
		RecCount = 0
		Select Case Request.QueryString("SEARCHTYPE")
			Case "B"
				DEPTCODE = Request.QueryString("SearchDEPTCODE") & "%"
				DeptName = Request.QueryString("SearchDeptName") & "%"
				AHSID = Request.QueryString("SearchAHSID") & "%"
			Case "C"
				DEPTCODE = "%" & Request.QueryString("SearchDEPTCODE") & "%"
				DeptName = "%" & Request.QueryString("SearchDeptName") & "%"
				AHSID = "%" & Request.QueryString("SearchAHSID") & "%"
			Case "E"
				DEPTCODE = Request.QueryString("SearchDeptCODE")
				DeptName = Request.QueryString("SearchDeptName")
				AHSID = Request.QueryString("SearchAHSID")
		End Select
		DEPTCODE = Replace(DEPTCODE, "'", "''")
		DeptName = Replace(DeptName, "'", "''")
		AHSID = Replace(AHSID, "'", "''")
	
		If Request.QueryString("SearchDEPTCODE") <> "" Then
			WHERECLS = WHERECLS & "DEPARTMENT_CODE LIKE '" & DEPTCODE  & "'"
		End If
		If Request.QueryString("SearchDeptName") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "DEPARTMENT_NAME LIKE '" & DEPTNAME & "'"
		End If
		If Request.QueryString("SearchAHSID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "ACCNT_HRCY_STEP_ID LIKE '" & AHSID & "'"
		End If
			
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM DEPARTMENT_CODES "
		if WHERECLS <> "" then SQLST = SQLST & " WHERE " & WHERECLS
		SQLST = SQLST & " ORDER BY DEPARTMENT_CODES_ID" 
				
		Set RS = Server.CreateObject("ADODB.Recordset")
		RS.MaxRecords = MAXRECORDCOUNT
		RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
		if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow"  BID='' >
	<td COLSPAN=8 NOWRAP CLASS="ResultCell" >No Department found re-check your criteria</td>
</tr>
	
<%		Else
			Do While Not RS.EOF
				RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);"  DEPTID='<%=RS("DEPARTMENT_CODES_ID")%>'>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("DEPARTMENT_CODES_ID"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("DEPARTMENT_NAME"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("DEPARTMENT_CODE"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ACCNT_HRCY_STEP_ID"))%></td>
	</tr>

<%
			RS.MoveNext
			Loop
					
		End If
		RS.Close
		Set RS = Nothing
		Conn.Close
		Set Conn = Nothing
	End if
	
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
