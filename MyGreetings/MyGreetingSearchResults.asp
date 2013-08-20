<%
'***************************************************************
'display the results of a Mailbox query in table format.
'
'$History: MyGreetingSearchResults.asp $ 
'* 
'* *****************  Version 3  *****************
'* User: Jenny.cheung Date: 6/18/08    Time: 1:35p
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/MyGreetings
'* 
'* *****************  Version 3  *****************
'* User: Jenny.cheung Date: 6/18/08    Time: 1:33p
'* Updated in $/FNS_DESIGNER/Source/Designer/MyGreetings
'* 
'* *****************  Version 2  *****************
'* User: Jenny.cheung Date: 6/11/08    Time: 4:31p
'* Updated in $/FNS_DESIGNER/Source/Designer/MyGreetings
'* JCHE-0021 To Incorporate Greeting table in Designer for user setup on
'* the Location User page.
'* 
'* *****************  Version 2  *****************
'* User: Jenny.cheung Date: 6/11/08    Time: 4:26p
'* Updated in $/FNS_DESIGNER/Release/C-INetPub/Designer/MyGreetings
'* took out stop
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 6/11/08    Time: 4:14p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/MyGreeting
'* JCHE-0021 To Incorporate Greeting table in Designer for user setup on
'* the Location User page.
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 6/11/08    Time: 4:09p
'* Created in $/FNS_DESIGNER/Source/Designer/Greeting
'* JCHE-0021 To Incorporate Greeting table in Designer for user setup on
'* the Location User page.
'* 
'* *****************  Version 1  *****************
'* User: Jenny.cheung Date: 4/21/08    Time: 9:23a
'* Created in $/FNS_DESIGNER/Source/Designer
'* created for Sedgwick.  Just want to save my work for now
'* 
'* *****************  Version 1  *****************
'* User: Alex.shimberg Date: 4/30/06    Time: 9:46p
'* Created in $/FNS_DESIGNER/Release/C-INetPub/Designer/Mailbox
'* Hartford SRS: Initial revision
'***************************************************************
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
<title>Greeting Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.GreetingID.value = ""
	end if
End Sub

Function GetGreetingID
	GetGreetingID = getmultipleindex(document.all.tblFields, "GreetingID")
End Function

Function GetGreetingText
	GetGreetingText = getmultipleindex(document.all.tblFields, "GreetingText")
End Function




Function ExeDelete
	bRet = false
	if CStr(document.body.getAttribute("ScreenMode")) = "RO" then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeDelete = bRet
		exit Function
	end if
	
	if document.all.GreetingID.value = "" then
		ExeDelete = false
		exit function
	end if


		document.all.TxtAction.value = "DELETE"
		sResult = document.all.AID.value
		document.all.TxtSaveData.Value = sResult
		document.all.FrmDetails.Submit()
		ExeDelete = true
	
End Function
</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "GreetingID")
		return objRow.getAttribute("GreetingID");
	else if (whichCol == "GreetingText")		
		return objRow.cells("GreetingText").innerText;
}
</SCRIPT>
</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0"  rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Greeting ID</div></td>
			<td class="thd"><div id><nobr>Contract Number</div></td>
			<td class="thd"><div id><nobr>Greeting Text</div></td>
				<td class="thd"><div id><nobr>LOB Codes</div></td>
			<td class="thd"><div id><nobr>Employee Feed</div></td>
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
				GreetingID = Request.QueryString("SearchGreetingID") & "%"
				ContractNumber = Request.QueryString("SearchContractNumber") & "%"
			Case "C"
				GreetingID = "%" & Request.QueryString("SearchGreetingID") & "%"
				ContractNumber = "%" & Request.QueryString("SearchContractNumber") & "%"
			Case "E"
				GreetingID = Request.QueryString("SearchGreetingID")
				ContractNumber = Request.QueryString("SearchContractNumber")
		End Select
		GreetingID = Replace(GreetingID, "'", "''")
		ContractNumber = Replace(ContractNumber, "'", "''")
	
		If Request.QueryString("SearchGreetingID") <> "" Then
			WHERECLS = WHERECLS & "GREETINGS_ID LIKE '" & GreetingID  & "'"
		End If
		
		If Request.QueryString("SearchContractNumber") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "CONTRACT_NUM LIKE '" & ContractNumber & "'"
		End If
			
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM GREETINGS "
		if WHERECLS <> "" then SQLST = SQLST & " WHERE " & WHERECLS
		SQLST = SQLST & " ORDER BY GREETINGS_ID" 
				
		Set RS = Server.CreateObject("ADODB.Recordset")
		RS.MaxRecords = MAXRECORDCOUNT
		RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
		if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow"  BID='' >
	<td COLSPAN=5 NOWRAP CLASS="ResultCell" >No Greeting found re-check your criteria</td>
</tr>
	
<%		Else
			Do While Not RS.EOF
				RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);"  GreetingID='<%=RS("GREETINGS_ID")%>'>
		<td NOWRAP CLASS="ResultCell" id = "GreetingId"><%=renderCell(RS("GREETINGS_ID"))%></td>
		<td NOWRAP CLASS="ResultCell" id = "ContractNumber"><%=renderCell(RS("CONTRACT_NUM"))%></td>
		<td WRAP CLASS="ResultCell" id = "GreetingText"><%=renderCell(RS("TEXT"))%></td>
		<td WRAP CLASS="ResultCell" id = "LOBText"><%=renderCell(RS("LOB_CODES"))%></td>
		<td NOWRAP CLASS="ResultCell" id = "EmployeeFeedFlag"><%=renderCell(RS("HAS_EMPLOYEE_FEED"))%></td>
				
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
