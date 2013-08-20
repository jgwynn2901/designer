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
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Attribute Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.BID.value = ""
	end if
End Sub

Function GetBID
	GetBID = Trim(getmultipleindex(document.all.tblFields, "BID"))
End Function

Function GetBIDName
	GetBIDName = getmultipleindex(document.all.tblFields, "NAME")
End Function

Function GetBIDCaption
	GetBIDCaption = getmultipleindex(document.all.tblFields, "CAPTION")
End Function

Function GetBIDInputType
	GetBIDInputType = getmultipleindex(document.all.tblFields, "INPUTTYPE")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "BID")
		return objRow.getAttribute("BID");
	else if (whichCol == "NAME")		
		return objRow.cells("NAME").innerText;
	else if (whichCol == "CAPTION")		
		return objRow.cells("CAPTION").innerText;
	else if (whichCol == "INPUTTYPE")		
		return objRow.cells("INPUTTYPE").innerText;
		
}
</SCRIPT>
</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Billing Id</div></td>
			<td class="thd"><div id><nobr>Accnt ID</div></td>
			<td class="thd"><div id><nobr>LOB</div></td>
			<td class="thd"><div id><nobr>Fee Amount</div></td>
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
				BID = Request.QueryString("SearchBID") & "%"
				ACCNT_HRCY_STEP_ID = Request.QueryString("SearchACCNT_HRCY_STEP_ID") 
				LOB_CD = Request.QueryString("SearchLOB_CD") 
				FEE_TYPE_ID = Request.QueryString("SearchFEE_TYPE_ID") & "%"
				DESCRIPTION = Request.QueryString("SearchDESCRIPTION") & "%"
			Case "C"
				BID = "%" & Request.QueryString("SearchBID") & "%"
				ACCNT_HRCY_STEP_ID =  Request.QueryString("SearchACCNT_HRCY_STEP_ID") 
				LOB_CD =  Request.QueryString("SearchLOB_CD") 
				FEE_TYPE_ID = "%" & Request.QueryString("SearchFEE_TYPE_ID") & "%"
				DESCRIPTION = "%" & Request.QueryString("SearchDESCRIPTION") & "%"
			Case "E"
				BID = Request.QueryString("SearchBID")
				ACCNT_HRCY_STEP_ID = Request.QueryString("SearchACCNT_HRCY_STEP_ID")
				LOB_CD = Request.QueryString("SearchLOB_CD")
				FEE_TYPE_ID = Request.QueryString("SearchFEE_TYPE_ID")
				DESCRIPTION = Request.QueryString("SearchDESCRIPTION")
		End Select
		BID = Replace(BID, "'", "''")
		ACCNT_HRCY_STEP_ID = Replace(ACCNT_HRCY_STEP_ID, "'", "''")
		LOB_CD = Replace(LOB_CD, "'", "''")
		FEE_TYPE_ID = Replace(FEE_TYPE_ID, "'", "''")
		DESCRIPTION = Replace(DESCRIPTION, "'", "''")
		
		If Request.QueryString("SearchBID") <> "" Then
			WHERECLS = WHERECLS & "UPPER(FEE_ID) LIKE '" & UCASE(BID)  & "'"
		End If
		If Request.QueryString("SearchACCNT_HRCY_STEP_ID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "ACCNT_HRCY_STEP_ID =" & ACCNT_HRCY_STEP_ID 
		End If
		If Request.QueryString("SearchLOB_CD") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "LOB_CD = '" & LOB_CD & "'"
		End If
		If Request.QueryString("SearchFEE_TYPE_ID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(FEE_TYPE_ID) LIKE '" & UCASE(FEE_TYPE_ID) & "'"
		End If
		If Request.QueryString("SearchDESCRIPTION") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(DESCRIPTION) LIKE '" & UCASE(DESCRIPTION) & "'"
		End If
		
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT FEE_ID, ACCNT_HRCY_STEP_ID, LOB_CD, FEE_AMOUNT, DESCRIPTION FROM FEE "
			
			If WHERECLS <> "" Then
				SQLST = SQLST & " WHERE " & WHERECLS 
			End If
			SQLST = SQLST & " ORDER BY FEE_ID" 
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" BID='' >
	<td COLSPAN=7 NOWRAP CLASS="ResultCell">No fees found re-check your criteria</td>
</tr>
	
	<%		Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  BID='<%=RS("FEE_ID")%>'>
	<td NOWRAP CLASS="ResultCell" ID="AID"><%=renderCell(RS("FEE_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("ACCNT_HRCY_STEP_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("LOB_CD"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="CAPTION" >
	<%If NOT IsNull(RS("FEE_AMOUNT"))  Then
			Response.write(FormatCurrency(RS("FEE_AMOUNT")))
		else
			Response.write("&nbsp;")
		End If %></td>
	<td NOWRAP CLASS="ResultCell" ID="CAPTION" ><%= renderCell(RS("DESCRIPTION"))%></td>
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

