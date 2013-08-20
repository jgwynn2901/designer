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
<title>Greeting Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		'Parent.frames("TOP").document.all.EPID.value = ""
	end if
End Sub

Function GetIBID
	GetIBID = getmultipleindex(document.all.tblFields, "IBID")
End Function

Function GetIBIDName
	GetIBIDName = getmultipleindex(document.all.tblFields, "NAME")
End Function


</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{

	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "IBID")
		return objRow.getAttribute("IBID");
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
		    <td class="thd"><div id><nobr>Inbound Call ID</div></td>
			<td class="thd"><div id><nobr>A.H. Step ID</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>Phone Number</div></td>
			<td class="thd"><div id><nobr>Description</div></td>
			<td class="thd"><div id><nobr>Greeting</div></td>
			
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
			
			    AHSID       = Request.QueryString("SearchAHSID")
				NAME        = Request.QueryString("SearchName") & "%"
				PHONENUMBER = Request.QueryString("SearchPhoneNumber") & "%"
				DESCRIPTION = Request.QueryString("SearchDescription") & "%"
				GREETING    = Request.QueryString("SearchGreeting") & "%"
			Case "C"
			   
				AHSID       = Request.QueryString("SearchAHSID") & "%"
				NAME        = "%" & Request.QueryString("SearchName") & "%"
				PHONENUMBER = "%" &  Request.QueryString("SearchPhoneNumber") & "%"
				DESCRIPTION = "%" & Request.QueryString("SearchDescription") & "%"
				GREETING    = "%" & Request.QueryString("SearchGreeting") & "%"
			Case "E"
			   
				AHSID       = Request.QueryString("SearchAHSID")
				NAME        = Request.QueryString("SearchName") 
				PHONENUMBER = Request.QueryString("SearchPhoneNumber")
				DESCRIPTION = Request.QueryString("SearchDescription") 
				GREETING    = Request.QueryString("SearchGreeting")
		End Select
		
		  'IBID        = Replace(INBOUNDCALL_ID, "'", "''")
		  'AHSID       = Replace(AHSID, "'", "''")
		  NAME        = Replace(NAME , "'", "''")
		  PHONENUMBER = Replace(PHONENUMBER , "'", "''")
		  DESCRIPTION = Replace(DESCRIPTION , "'", "''")
		  GREETING    = Replace(GREETING , "'", "''")
		  
		If Request.QueryString("SearchAHSID") <> "" Then
			WHERECLS = WHERECLS & "ACCNT_HRCY_STEP_ID = " & AHSID
		End If
		If Request.QueryString("SearchName") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "NAME LIKE '" & NAME & "'"
		End If
		If Request.QueryString("SearchPhoneNumber") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(PHONENUMBER) LIKE '" & UCASE(PHONENUMBER) & "'"
		End If
		
		If Request.QueryString("SearchDescription") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(DESCRIPTION) LIKE '" & UCASE(DESCRIPTION) & "'"
		End If
		
		If Request.QueryString("SearchGreeting") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(GREETING) LIKE '" & UCASE(GREETING) & "'"
		End If
		
          Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT INBOUNDCALL_ID,ACCNT_HRCY_STEP_ID,NAME,PHONENUMBER,DESCRIPTION,GREETING FROM INBOUNDCALL "
			
			if WHERECLS <> "" Then
				SQLST = SQLST & " WHERE " & WHERECLS
			End if
			
			SQLST = SQLST & " ORDER BY INBOUNDCALL_ID,ACCNT_HRCY_STEP_ID "
			
			'Response.write(SQLST & "<BR>") 
			
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" EPID='' >
	<td COLSPAN=10 NOWRAP CLASS="ResultCell">No Greetings found. Re-check your criteria</td>
</tr>
	
	<%		Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  IBID='<%=RS("INBOUNDCALL_ID")%>'>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("INBOUNDCALL_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("ACCNT_HRCY_STEP_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("NAME"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("PHONENUMBER"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("DESCRIPTION"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("GREETING"))%></td>
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
