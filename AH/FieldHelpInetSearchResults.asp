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
<title>Field Help Inetinternal Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		'Parent.frames("TOP").document.all.HPID.value = ""
	end if
End Sub

Function GetHPID

	GetHPID = getmultipleindex(document.all.tblFields, "HPID")
End Function

Function GetHPIDName

	GetHPIDName = getmultipleindex(document.all.tblFields, "NAME")
End Function


</SCRIPT>

<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
  currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "HPID")
		return objRow.getAttribute("HPID");
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
		    <td class="thd"><div id><nobr>Help ID</div></td>
		    <td class="thd"><div id><nobr>Tab Order</div></td>
		    <td class="thd"><div id><nobr>Field</div></td>
			<td class="thd"><div id><nobr>A.H. Step ID</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>LOB</div></td>
			<td class="thd"><div id><nobr>Help Type</div></td>
			<td class="thd"><div id><nobr>Help Text</div></td>
			
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
			   
				NAME        = Request.QueryString("SearchName")        & "%"
				AHSID       = Request.QueryString("SearchAHSID")       & "%"
			    HELP_TYPE   = Request.QueryString("SearchHelpType")   
				TABORDER    = Request.QueryString("SearchTAB_ORDER")   & "%"
				LOB_CD      = Request.QueryString("SearchLOBCD")       & "%"
				FIELD       = Request.QueryString("SearchField")       & "%"
				HELP_TEXT   = Request.QueryString("SearchHelpText") & "%"
				
			Case "C"
			   
				NAME       = "%" &  Request.QueryString("SearchName")      & "%"
				AHSID      = "%" &  Request.QueryString("SearchAHSID")     & "%"
			    HELP_TYPE  =      Request.QueryString("SearchHelpType")  
				TABORDER   = "%" &  Request.QueryString("SearchTAB_ORDER") & "%"
				LOB_CD     = "%" &  Request.QueryString("SearchLOBCD")     & "%"
				FIELD      = "%" &  Request.QueryString("SearchField")     & "%"
				HELP_TEXT  = "%" &  Request.QueryString("SearchHelpText")  & "%"
			Case "E"
			     
				 NAME      = Request.QueryString("SearchName") 
				 AHSID     = Request.QueryString("SearchAHSID") 
				 HELP_TYPE = Request.QueryString("SearchHelpType")
				 TABORDER  = Request.QueryString("SearchTAB_ORDER")
				 LOB_CD    = Request.QueryString("SearchLOBCD") 
				 FIELD     = Request.QueryString("SearchField") 
				 HELP_TEXT = Request.QueryString("SearchHelpText") 
		End Select
		'HPID = Replace(HPID, "'", "''")
		    
		     NAME      = Replace(NAME, "'", "''")
		     AHSID     = Replace(AHSID , "'", "''")
		     HELP_TYPE = Replace(HELP_TYPE, "'", "''")
		     LOB_CD    = Replace(LOB_CD, "'", "''")
		     FIELD = Replace(FIELD, "'", "''")
		     HELP_TEXT = Replace(HELP_TEXT, "'", "''")
	
		If Request.QueryString("SearchAHSID") <> "" Then
		    WHERECLS = WHERECLS & " AND "
			WHERECLS = WHERECLS & "UPPER(ACCNT_HRCY_STEP_ID) LIKE '" & UCASE(AHSID)  & "'"
		End If
		
		If Request.QueryString("SearchName") <> "" Then
		 
			'If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			'End If
			WHERECLS = WHERECLS & "CH.NAME LIKE '" & UCASE(NAME) & "'"
		End If
		If Request.QueryString("SearchHelpType") <> "" Then
		
			'If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			'End If
			
			WHERECLS = WHERECLS & "CH.HELP_TYPE_ID = '" & HELP_TYPE  & "'"
		End If
		
		If Request.QueryString("SearchTAB_ORDER") <> "" Then
		
			'If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			'End If
			
			WHERECLS = WHERECLS & "TAB_ORDER LIKE'" & TABORDER  & "'"
		End If
		
		
		If Request.QueryString("SearchLOBCD") <> "" Then
		
			'IF WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			'End If
			WHERECLS = WHERECLS & "UPPER( LOB_CD) LIKE '" & UCASE( LOB_CD) & "'"
		End If
		
		If Request.QueryString("SearchField") <> "" Then
		 
			'If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			'End If
			WHERECLS = WHERECLS & "UPPER(SUBSTR(FIELD,2)) LIKE '" & UCASE(FIELD) & "'"
			
		End If
		If Request.QueryString("SearchHelpText") <> "" Then
		
			'If WHERECLS <> "" Then 
				WHERECLS = WHERECLS & " AND "
			'End If
			WHERECLS = WHERECLS & "UPPER( HELP_TEXT) LIKE '" & UCASE( HELP_TEXT) & "'"
		End If



  
			Set Conn = Server.CreateObject("ADODB.Connection")
		    Conn.Open CONNECT_STRING
			SQLST = "SELECT CH.NAME as CallName ,CH.HELP_ID,HT.NAME AS TypeName,"
			SQLST = SQLST & " CH.LOB_CD,CH.TAB_ORDER,CH.ACCNT_HRCY_STEP_ID,"
			SQLST = SQLST & " CH.HELP_TYPE_ID AS HelpType ,CH.FIELD,CH.HELP_TEXT "
			SQLST = SQLST & " FROM CALL_HELP CH,HELP_TYPE HT "
			SQLST = SQLST & " WHERE CH.HELP_TYPE_ID = HT.HELP_TYPE_ID "
			
			If WHERECLS <> "" Then
				SQLST = SQLST & WHERECLS
			End If
			
			'Response.write(SQLST & "<BR>")
			
			SQLST = SQLST & "ORDER BY CH.HELP_ID" 
			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	 
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" HPID='' >
	<td COLSPAN=10 NOWRAP CLASS="ResultCell">No Field Help Inetinternal found re-check your criteria</td>
</tr>
	
	<%		Else
	
				Do While Not RS.EOF
					RecCount = RecCount + 1
					
			dim RsField,trimField
			trimField = trim(RS("FIELD"))
			RsField=MID(trimField,2)
					
%>   

	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);"  HPID='<%=RS("HELP_ID")%>'>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("HELP_ID"))%></td>
    <td NOWRAP CLASS="ResultCell"><%=renderCell(RS("TAB_ORDER"))%></td>
	<td NOWRAP CLASS="ResultCell"ID="NAME"><%=renderCell(RsField)%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ACCNT_HRCY_STEP_ID"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("CallName"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("LOB_CD"))%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("TypeName"))%>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(RS("HELP_TEXT"))%></td>
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
