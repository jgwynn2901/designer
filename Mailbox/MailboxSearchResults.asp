<%
'***************************************************************
'display the results of a Mailbox query in table format.
'
'$History: MailboxSearchResults.asp $ 
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
<title>Mailbox Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.MBID.value = ""
	end if
End Sub

Function GetMBID
	GetMBID = getmultipleindex(document.all.tblFields, "MBID")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "MBID")
		return objRow.getAttribute("MBID");
}
</SCRIPT>
</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0"  rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Mailbox ID</div></td>
			<td class="thd"><div id><nobr>Mailbox Number</div></td>
			<td class="thd"><div id><nobr>AH Load ID</div></td>
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
				MBID = Request.QueryString("SearchMBID") & "%"
				MAILBOXNUMBER = Request.QueryString("SearchMailboxNumber") & "%"
				AHLOADID = Request.QueryString("SearchAHLoadID") & "%"
			Case "C"
				MBID = "%" & Request.QueryString("SearchMBID") & "%"
				MAILBOXHNUMBER = "%" & Request.QueryString("SearchMailboxNumber") & "%"
				AHLOADID = "%" & Request.QueryString("SearchAHLoadID") & "%"
			Case "E"
				MBID = Request.QueryString("SearchMBID")
				MAILBOXNUMBER = Request.QueryString("SearchMailboxNumber")
				AHLOADID = Request.QueryString("SearchAHLoadID")
		End Select
		MBID = Replace(MBID, "'", "''")
		MAILBOXNUMBER = Replace(MAILBOXNUMBER, "'", "''")
		AHLOADID = Replace(AHLOADID, "'", "''")
	
		If Request.QueryString("SearchMBID") <> "" Then
			WHERECLS = WHERECLS & "MAILBOX_ID LIKE '" & MBID  & "'"
		End If
		If Request.QueryString("SearchMailboxNumber") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "MAILBOX_NUMBER LIKE '" & MAILBOXNUMBER & "'"
		End If
		If Request.QueryString("SearchAHLoadID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "ACCOUNT_HIERARCHY_LOAD_ID LIKE '" & AHLOADID & "'"
		End If
			
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM MAILBOX "
		if WHERECLS <> "" then SQLST = SQLST & " WHERE " & WHERECLS
		SQLST = SQLST & " ORDER BY MAILBOX_ID" 
				
		Set RS = Server.CreateObject("ADODB.Recordset")
		RS.MaxRecords = MAXRECORDCOUNT
		RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
		if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow"  BID='' >
	<td COLSPAN=8 NOWRAP CLASS="ResultCell" >No Mailboxes found re-check your criteria</td>
</tr>
	
<%		Else
			Do While Not RS.EOF
				RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);"  MBID='<%=RS("MAILBOX_ID")%>'>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("MAILBOX_ID"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("MAILBOX_NUMBER"))%></td>
		<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ACCOUNT_HIERARCHY_LOAD_ID"))%></td>
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
