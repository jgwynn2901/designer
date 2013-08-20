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
<title>Rule Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.RID.value = ""
	end if
End Sub

Function GetRID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetRID = document.all.tblFields.rows(idx).getAttribute("RULEID")
	else 
		GetRID = ""
	End If
End Function

Function GetRIDText
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("RULE_TEXT").innerText
	End If
	GetRIDText = strText
End Function

Function GetRIDComments
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("COMMENTS").innerText
	End If
	GetRIDComments = strText
End Function

Function GetRIDUser
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("USER").innerText
	End If
	GetRIDUser = strText
End Function

Function GetRIDType
	dim idx, strText
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then 
		strText = document.all.tblFields.rows(idx).cells("RULE_TYPE").innerText
	End If
	GetRIDType = strText
End Function

</script>
</head>
<body BGCOLOR="<%=BODYBGCOLOR%>" topmargin=0 leftmargin=0  rightmargin=0 >
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:100%;width:100%">
<div align="LEFT" style="{display:block;height:100%;width:100%;overflow:auto}">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Rule Id</div></td>
			<td class="thd"><div id><nobr>Type</div></td>
			<td class="thd"><div id><nobr>Rule Text</div></td>
			<td class="thd"><div id><nobr>Comment</div></td>

		</tr>
	</thead>
	<tbody ID="TableRows">

<%
	dim RecCount
	dim UseUsersTable
	
	RecCount = -1
	UseUsersTable=0 
	
	If Request.QueryString("SEARCHTYPE") <> "" Then
	RecCount = 0

		Select Case Request.QueryString("SEARCHTYPE")
			Case "B"
				RULEID = Request.QueryString("SearchRuleId") & "%"
				RULETYPE = Request.QueryString("SearchRuleType") & "%"
				RULETEXT = Request.QueryString("SearchRuleText") & "%"
				RULECOMMENTS = Request.QueryString("SearchComments") & "%"
				RULEUSER = Request.QueryString("SearchUser") & "%"
			Case "C"
				RULEID = "%" & Request.QueryString("SearchRuleId") & "%"
				RULETYPE = "%" & Request.QueryString("SearchRuleType") & "%"
				RULETEXT = "%" & Request.QueryString("SearchRuleText") & "%"
				RULECOMMENTS = "%" & Request.QueryString("SearchComments") & "%"
				RULEUSER = "%" & Request.QueryString("SearchUser") & "%"
			Case "E"
				RULEID = Request.QueryString("SearchRuleId")
				RULETYPE = Request.QueryString("SearchRuleType")
				RULETEXT = Request.QueryString("SearchRuleText")
				RULECOMMENTS = Request.QueryString("SearchComments")
				RULEUSER =Request.QueryString("SearchUser") 
	End Select
	
		RULEID = Replace(RULEID,"'","''")
		RULETYPE = Replace(RULETYPE,"'","''")
		RULETEXT = Replace(RULETEXT,"'","''")
		RULECOMMENTS = Replace(RULECOMMENTS,"'","''")
		RULEUSER = Replace(RULEUSER,"'","''")
		
		If Request.QueryString("SearchRuleId") <> "" Then
			WHERECLS = WHERECLS & "RULES.RULE_ID LIKE '" & RULEID  & "'"
		End If
		If Request.QueryString("SearchRuleType") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(RULES.TYPE) LIKE '" & UCASE(RULETYPE)  & "'"
		End If
		If Request.QueryString("SearchRuleText") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(RULES.RULE_TEXT) LIKE '" & UCASE(RULETEXT) & "'"
		End If

		If Request.QueryString("SearchComments") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(RULES.COMMENTS) LIKE '" & UCASE(RULECOMMENTS) & "'"
		End If
		
		If Request.QueryString("SearchUser") <> "" Then
			UseUsersTable=1 
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "UPPER(USERS.NAME) LIKE '" & UCASE(RULEUSER) & "'"
		End If

		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open CONNECT_STRING
		
		If UseUsersTable <> 0 Then
			SQLST = SQLST & "SELECT * FROM RULES,USERS " 
		else
			SQLST = SQLST & "SELECT * FROM RULES " 
		End If

		If WHERECLS <> "" Then
			SQLST = SQLST & " WHERE " & WHERECLS 
		End If

		If UseUsersTable <> 0 Then
			SQLST = SQLST & "AND RULES.USER_ID = USERS.USER_ID" 
		End If

		SQLST = SQLST & " ORDER BY RULES.RULE_TEXT" 

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
<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" RULEID='<%= RS("RULE_ID") %>'  >
	<td NOWRAP CLASS="ResultCell" ID="RULE_ID"><%=renderCell(RS("RULE_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="RULE_TYPE"><%=renderCell(RS("TYPE"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="RULE_TEXT" ><%=renderCell(RS("RULE_TEXT"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="COMMENTS" ><%=renderCell(RS("COMMENTS"))%></td>

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
