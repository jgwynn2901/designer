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
<title>AHS Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.AHSID.value = ""
	end if
End Sub

Function GetAHSID
	GetAHSID = getmultipleindex(document.all.tblFields, "AHSID")
End Function

Function getCientNode
	getCientNode = getmultipleindex(document.all.tblFields, "CLIENTNODE")
End Function

Function GetAHSIDName
	GetAHSIDName = getmultipleindex(document.all.tblFields, "NAME")
End Function

</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "AHSID")
		return objRow.getAttribute("AHSID");
	else if (whichCol == "NAME")		
		return objRow.cells("NAME").innerText;
	else if (whichCol == "CAPTION")		
		return objRow.cells("CAPTION").innerText;
	else if (whichCol == "INPUTTYPE")		
		return objRow.cells("INPUTTYPE").innerText;
	else if (whichCol == "CLIENTNODE")		
		return objRow.getAttribute("CLIENTNODE");
}
</SCRIPT>
</head>

<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%" >
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>AHSID</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>Client Code</div></td>
			<td class="thd"><div id><nobr>Type</div></td>
			<td class="thd"><div id><nobr>Upload key</div></td>
			<td class="thd"><div id><nobr>Location Code</div></td>
			<td class="thd"><div id><nobr>SUID</div></td>

			<td class="thd" style="display:none"><div id><nobr>Nature of Business</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">

<%
	dim RecCount, lExactSearch, cVerb
	RecCount = -1
	
	
If Request.QueryString("SEARCHTYPE") <> "" Then
	RecCount = 0
	lExactSearch = false
		Select Case Request.QueryString("SEARCHTYPE")
			Case "B"
				AHSID = Request.QueryString("SearchAHSID") 
				NAME = Request.QueryString("SearchName") & "%"
				FNS_CLIENT_CD = Request.QueryString("SearchFNS_CLIENT_CD") 
				LOB_CD = Request.QueryString("SearchLOB_CD") 
	
				ATYPE = Request.QueryString("Search_TYPE") & "%"
				UPLOAD_KEY = Request.QueryString("SearchUPLOAD_KEY") & "%"
				LOCATION_CODE = Request.QueryString("SearchLOCATION_CODE") & "%"
				SUID = Request.QueryString("SearchSUID") & "%"
				
				

			Case "C"
				AHSID =  Request.QueryString("SearchAHSID") 
				NAME = "%" & Request.QueryString("SearchName") & "%"
				FNS_CLIENT_CD =  Request.QueryString("SearchFNS_CLIENT_CD")
				LOB_CD =  Request.QueryString("SearchLOB_CD") 
				
				ATYPE = "%" & Request.QueryString("Search_TYPE") & "%"
				UPLOAD_KEY = "%" & Request.QueryString("SearchUPLOAD_KEY") & "%"
				LOCATION_CODE = "%" & Request.QueryString("SearchLOCATION_CODE") & "%"
				SUID = "%" & Request.QueryString("SearchSUID") & "%"
				
			Case "E"
				AHSID = Request.QueryString("SearchAHSID")
				NAME = Request.QueryString("SearchName")
				FNS_CLIENT_CD = Request.QueryString("SearchFNS_CLIENT_CD")
				LOB_CD = Request.QueryString("SearchLOB_CD")

				ATYPE = Request.QueryString("Search_TYPE")
				UPLOAD_KEY = Request.QueryString("SearchUPLOAD_KEY")
				LOCATION_CODE = Request.QueryString("SearchLOCATION_CODE")
				SUID = Request.QueryString("SearchSUID")
				lExactSearch = true
		End Select
		
		AHSID = Replace(AHSID, "'", "''")
		NAME = Replace(NAME, "'", "''")
		FNS_CLIENT_CD = Replace(FNS_CLIENT_CD, "'", "''")
		LOB_CD = Replace(LOB_CD, "'", "''")

				ATYPE = Replace(ATYPE, "'", "''")
				UPLOAD_KEY = Replace(UPLOAD_KEY, "'", "''")
				LOCATION_CODE = Replace(LOCATION_CODE, "'", "''")
				SUID = Replace(SUID, "'", "''")
		
		if lExactSearch then
			cVerb = "="
		else
			cVerb = "LIKE"
		end if
		If Request.QueryString("SearchName") <> "" Then
			WHERECLS = WHERECLS & "UPPER(NAME) " & cVerb & " '" & UCASE(NAME)  & "'"
		End If
		If Request.QueryString("SearchAHSID") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "ACCNT_HRCY_STEP_ID = " &  AHSID 
		else
			if Request.QueryString("SearchPARENT_NODE_ID") <> "" Then
           PARENT_NODE_ID = Request.QueryString("SearchPARENT_NODE_ID") 
				If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
				End If
				 WHERECLS = WHERECLS & " PARENT_NODE_ID = " &  PARENT_NODE_ID 
			end if
				
		End If
		If Request.QueryString("SearchFNS_CLIENT_CD") <> "" Then
			If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			End If
			WHERECLS = WHERECLS & "FNS_CLIENT_CD " & cVerb & " '" & UCASE(FNS_CLIENT_CD) & "'"
		End If
		
			      
      		

if Request.QueryString("Search_TYPE") <> "" Then
	If WHERECLS <> "" Then 
		WHERECLS = WHERECLS & " AND "
	End If
	WHERECLS = WHERECLS & "TYPE " & cVerb & " '" & UCASE(ATYPE) & "'"
End If
if Request.QueryString("SearchUPLOAD_KEY") <> "" Then
	If WHERECLS <> "" Then 
		WHERECLS = WHERECLS & " AND "
	End If
	WHERECLS = WHERECLS & "UPPER(UPLOAD_KEY) " & cVerb & " '" & UCASE(UPLOAD_KEY) & "'"
End If
if Request.QueryString("SearchLOCATION_CODE") <> "" Then
	If WHERECLS <> "" Then 
		WHERECLS = WHERECLS & " AND "
	End If
	WHERECLS = WHERECLS & "LOCATION_CODE " & cVerb & " '" & UCASE(LOCATION_CODE) & "'"
End If
if Request.QueryString("SearchSUID") <> "" Then
	If WHERECLS <> "" Then 
		WHERECLS = WHERECLS & " AND "
	End If
	WHERECLS = WHERECLS & "SUID " & cVerb & " '" & UCASE(SUID) & "'"
End If





		
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = ""
			SQLST = SQLST & "SELECT NAME, FNS_CLIENT_CD, TYPE, UPLOAD_KEY, LOCATION_CODE, SUID, "
			SQLST = SQLST & "ACCNT_HRCY_STEP_ID, NATURE_OF_BUSINESS, CLIENT_NODE_ID FROM "
			SQLST = SQLST & "ACCOUNT_HIERARCHY_STEP "
			
			If WHERECLS <> "" Then
				SQLST = SQLST & "WHERE ACTIVE_STATUS = 'ACTIVE' AND " & WHERECLS
			End If
			SQLST = SQLST & " ORDER BY NAME" 
			'Response.Write(SQLST)
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);"  AHSID='' >
	<td COLSPAN=7 align="center" NOWRAP CLASS="ResultCell">No AHS found re-check your criteria</td>
</tr>
	
	<%	Else
			Do While Not RS.EOF
			RecCount = RecCount + 1
%>
	<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);"  AHSID='<%=RS("ACCNT_HRCY_STEP_ID")%>' CLIENTNODE="<%=RS("CLIENT_NODE_ID")%>">
	<td NOWRAP CLASS="ResultCell" ID="AHSID"><%=renderCell(RS("ACCNT_HRCY_STEP_ID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="NAME"><%=renderCell(RS("NAME"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="CAPTION" ><%=renderCell(RS("FNS_CLIENT_CD"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="CAPTION" ><%=renderCell(RS("TYPE"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="CAPTION" ><%=renderCell(RS("UPLOAD_KEY"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="CAPTION" ><%=renderCell(RS("LOCATION_CODE"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="CAPTION" ><%=renderCell(RS("SUID"))%></td>
	<td NOWRAP CLASS="ResultCell" ID="INPUTTYPE" style="display:none"><%=RS("NATURE_OF_BUSINESS")%></td>
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
