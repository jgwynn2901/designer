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
<title>Dictionary Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.RID.value = ""
	end if
End Sub


Function GetDictText
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("DICT_TEXT").innerText
	End If
	GetDictText = strText
End Function

</script>
</head>
<body BGCOLOR="<%=BODYBGCOLOR%>" topmargin=0 leftmargin=0  rightmargin=0 >
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:100%;width:100%">
<div align="LEFT" style="{display:block;height:100%;width:100%;overflow:auto}">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">

			<td class="thd"><div id><nobr>Word</div></td>

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
				DICTTEXT = Request.QueryString("SearchDictText") & "%"
			Case "C"
				DICTTEXT = "%" & Request.QueryString("SearchDictText") & "%"
			Case "E"
				DICTTEXT = Request.QueryString("SearchDictText")
		End Select

		DICTTEXT = Replace(DICTTEXT,"'","''")
		
		If Request.QueryString("SearchDictText") <> "" Then
			WHERECLS = WHERECLS & "WORD  LIKE '" & DICTTEXT  & "'"
		End If
		
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = SQLST & "SELECT * FROM SPELL_CHECK "
			If WHERECLS <> "" Then
				SQLST = SQLST & " WHERE " & WHERECLS 
			End If
			SQLST = SQLST & " ORDER BY word" 

			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText

			If RS.EOF And RS.BOF then
%>
<tr ID="FieldRow" CLASS="ResultRow" RULEID='' >
	<td COLSPAN=3 NOWRAP CLASS="ResultCell">No rules found re-check your criteria</td>
</tr>
<%			Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>
<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" WORD='<%= RS("WORD") %>'  >
	<td NOWRAP CLASS="ResultCell" ID="DICT_TEXT" ><%=renderCell(RS("WORD"))%></td>

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
