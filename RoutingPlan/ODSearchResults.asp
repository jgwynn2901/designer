<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Output Definition Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">


<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--
Sub Window_onLoad
<% If Request.QueryString <> "" Then %>
if 0 < Document.all.tblFields.rows.length then
		call multiselect( Document.all.tblFields.rows(1))
		'call Document.all.tblFields.focus()
end if
<% End if %>
End Sub

Function GetDef()
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		idx = document.all.tblFields.rows(idx).getAttribute("OUTPUTDEFID")
	End If
	GetDef = idx
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
<body BGCOLOR='<%=BODYBGCOLOR%>'  leftmargin=0 rightmargin=0 topmargin=0 bottommargin=0>
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699">
<div align="LEFT" style="display:block;height:135;width:'100%';overflow:scroll">
<table cellPadding="2" cellSpacing="0" frame="void" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Output Id</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>Description</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
	If Request.QueryString <> "" Then

		Select Case Request.QueryString("SEARCHTYPE")
			Case "B"
				OUTPUTDEF_ID = Request.QueryString("OUTPUTDEF_ID") & "%"
				NAME = Request.QueryString("NAME") & "%"
				DESCRIPTION = Request.QueryString("DESCRIPTION") & "%"
			Case "C"
				OUTPUTDEF_ID = "%" & Request.QueryString("OUTPUTDEF_ID") & "%"
				NAME = "%" & Request.QueryString("NAME") & "%"
				DESCRIPTION = "%" & Request.QueryString("DESCRIPTION") & "%"
			Case "E"
				OUTPUTDEF_ID = Request.QueryString("OUTPUTDEF_ID")
				NAME = Request.QueryString("NAME")
				DESCRIPTION = Request.QueryString("DESCRIPTION")
		End Select

	
		OUTPUTDEF_ID = Replace(OUTPUTDEF_ID,"'","''")
		NAME = Replace(NAME,"'","''")
		DESCRIPTION = Replace(DESCRIPTION,"'","''")
		
		If Request.QueryString("OUTPUTDEF_ID") <> "" Then
			WHERECLS = WHERECLS & "OUTPUTDEF_ID LIKE '" & OUTPUTDEF_ID  & "'"
		End If
		If Request.QueryString("NAME") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(NAME) LIKE '" & UCASE(NAME)  & "'"
		End If
		If Request.QueryString("DESCRIPTION") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(DESCRIPTION) LIKE '" & UCASE(DESCRIPTION) & "'"
		End If

		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		SQLST = "SELECT * FROM OUTPUT_DEFINITION WHERE " & WHERECLS & " ORDER BY NAME" 
		Set RS = Conn.Execute(SQLST)
		If RS.EOF AND RS.BOF Then
%>
<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);"  OUTPUTDEFID='X' >
	<td COLSPAN=3 NOWRAP CLASS="LABEL" ID="FRAME_ID">No output definitions found re-check your criteria</td>
</tr>
				
<%Else
	Do While Not RS.EOF %>
<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);"  OUTPUTDEFID='<%= RS("OUTPUTDEF_ID") %>' >
	<td NOWRAP CLASS="LABEL" ID="FRAME_ID"><%=renderCell(RS("OUTPUTDEF_ID"))%></td>
	<td NOWRAP CLASS="LABEL" ID="NAME"><%=renderCell(RS("NAME"))%></td>
	<td NOWRAP CLASS="LABEL" ID="DESCRIPTION" ><%=renderCell(RS("DESCRIPTION"))%></td>
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
</body>
</html>
