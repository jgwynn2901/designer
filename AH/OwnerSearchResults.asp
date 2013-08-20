<%
	Response.Expires = 0
	Response.Buffer = true
%>

<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Owner Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		'Parent.frames("TOP").document.all.searchOID.value = ""
	end if
End Sub

Function GetOID
	GetOID = getmultipleindex(document.all.tblFields, "OID")
End Function
</script>
<SCRIPT LANGUAGE="JScript">
function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "OID")
		return objRow.getAttribute("OID");
		
}
</SCRIPT>

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Owner Id</div></td>
			<td class="thd"><div id><nobr>Title</div></td>
			<td class="thd"><div id><nobr>Name</div></td>
			<td class="thd"><div id><nobr>Address</div></td>
			<td class="thd"><div id><nobr>City</div></td>
			<td class="thd"><div id><nobr>State</div></td>
			<td class="thd"><div id><nobr>WPhone</div></td>
			<td class="thd"><div id><nobr>Fax</div></td>
<!--			<td class="thd"><div id><nobr>A.H.S ID</div></td>
			<td class="thd"><div id><nobr>Active Start Dt</div></td>
			<td class="thd"><div id><nobr>Active End Dt</div></td>-->
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
				OID       = Request.QueryString("SearchOID")       & "%"
				TITLE     = Request.QueryString("SearchTitle")     & "%"
				NAMELAST  = Request.QueryString("SearchNameLast")  & "%"
				NAMEFIRST = Request.QueryString("SearchNameFirst") & "%"
				ADD1      = Request.QueryString("SearchAdd1")      & "%"
				ADD2      = Request.QueryString("SearchAdd2")      & "%"
				CITY      = Request.QueryString("SearchCity")      & "%"
				STATE     = Request.QueryString("SearchState")     & "%"
				ZIP       = Request.QueryString("SearchZip")       & "%"
				WPHONE    = Request.QueryString("SearchWPhone")    & "%"
				HPHONE    = Request.QueryString("SearchHPhone")    & "%"
				FAX       = Request.QueryString("SearchFax")       & "%"
			Case "C"
				OID       = "%" & Request.QueryString("SearchOID")       & "%"
				TITLE     = "%" & Request.QueryString("SearchTitle")     & "%"
				NAMELAST  = "%" & Request.QueryString("SearchNameLast")  & "%"
				NAMEFIRST = "%" & Request.QueryString("SearchNameFirst") & "%"
				ADD1      = "%" & Request.QueryString("SearchAdd1")      & "%"
				ADD2      = "%" & Request.QueryString("SearchAdd2")      & "%"
				CITY      = "%" & Request.QueryString("SearchCity")      & "%"
				STATE     = "%" & Request.QueryString("SearchState")     & "%"
				ZIP       = "%" & Request.QueryString("SearchZip")       & "%"
				WPHONE    = "%" & Request.QueryString("SearchWPhone")    & "%"
				HPHONE    = "%" & Request.QueryString("SearchHPhone")    & "%"
				FAX       = "%" & Request.QueryString("SearchFax")       & "%"
			Case "E"
				OID       = Request.QueryString("SearchOID")       
				TITLE     = Request.QueryString("SearchTitle")     
				NAMELAST  = Request.QueryString("SearchNameLast")  
				NAMEFIRST = Request.QueryString("SearchNameFirst") 
				ADD1      = Request.QueryString("SearchAdd1")      
				ADD2      = Request.QueryString("SearchAdd2")      
				CITY      = Request.QueryString("SearchCity")      
				STATE     = Request.QueryString("SearchState")     
				ZIP       = Request.QueryString("SearchZip")       
				WPHONE    = Request.QueryString("SearchWPhone")   
				HPHONE    = Request.QueryString("SearchHPhone")   
				FAX       = Request.QueryString("SearchFax")     
		End Select
		
		
		OID       = Replace(OID, "'", "''")
		TITLE     = Replace(TITLE, "'", "''")
		NAMELAST  = Replace(NAMELAST, "'", "''")
		NAMEFIRST = Replace(NAMEFIRST, "'", "''")
		ADD1      = Replace(ADD1, "'", "''")
		ADD2      = Replace(ADD2 , "'", "''")
		CITY      = Replace(CITY, "'", "''")
		STATE     = Replace(STATE, "'", "''")
		ZIP       = Replace(ZIP, "'", "''")
		WPHONE    = Replace(WPHONE, "'", "''")
		HPHONE    = Replace(HPHONE, "'", "''")
		FAX       = Replace(FAX , "'", "''")
		
		WHERECLS = " WHERE 1 = 1 "

		If Request.QueryString("SearchOID") <> "" Then
			WHERECLS = WHERECLS & "AND O.OWNER_ID LIKE '" & OID & "'"
		End If
		If Request.QueryString("SearchTitle") <> "" Then
			WHERECLS = WHERECLS & "AND O.NAME_TITLE LIKE '" & TITLE & "'"
		End If
		If Request.QueryString("SearchNameLast") <> "" Then
			WHERECLS = WHERECLS & "AND O.name_last LIKE '" & NAMELAST & "'"
		End If
		If Request.QueryString("SearchNameFirst") <> "" Then
			WHERECLS = WHERECLS & "AND O.NAME_FIRST LIKE '" & NAMEFIRST & "'"
		End If
		If Request.QueryString("Searchadd1") <> "" Then
			WHERECLS = WHERECLS & "AND O.ADDRESs_LINE1 LIKE '" & ADD1 & "'"
		End If
        If Request.QueryString("Searchadd2") <> "" Then
			WHERECLS = WHERECLS & "AND O.ADDRESs_LINE2 LIKE '" & ADD2 & "'"
		End If
        If Request.QueryString("SearchCity") <> "" Then
			WHERECLS = WHERECLS & "AND O.ADDRESs_CITY LIKE '" & CITY & "'"
		End If
        If Request.QueryString("SearchState") <> "" Then
			WHERECLS = WHERECLS & "AND O.ADDRESs_STATE LIKE '" & STATE & "'"
		End If
        If Request.QueryString("SearchZip") <> "" Then
			WHERECLS = WHERECLS & "AND O.ADDRESS_ZIP LIKE '" & ZIP & "'"
		End If
		If Request.QueryString("SearchWPhone") <> "" Then
			WHERECLS = WHERECLS & "AND O.PHONE_WORK LIKE '" & WPHONE & "'"
		End If
		If Request.QueryString("SearchHPhone") <> "" Then
			WHERECLS = WHERECLS & "AND O.PHONE_HOME LIKE '" & HPHONE & "'"
		End If
		If Request.QueryString("SearchFax") <> "" Then
			WHERECLS = WHERECLS & "AND O.PHONE_FAX LIKE '" & FAX & "'"
		End If
			Set Conn = Server.CreateObject("ADODB.Connection")
			Conn.Open CONNECT_STRING
			SQLST = "SELECT O.*  " &_
			        "  FROM OWNER O " 
			        '"       AHSO.ACCNT_HRCY_STEP_ID, " &_
			        
					'"       AHSO.ACTIVE_START_DT, " &_
					'"       AHSO.ACTIVE_END_DT " &_
					
					
			
			
			If WHERECLS <> "" Then
				SQLST = SQLST &  WHERECLS 
			End If
			
			SQLST = SQLST & " ORDER BY O.OWNER_ID" 

			Set RS = Server.CreateObject("ADODB.Recordset")
			RS.MaxRecords = MAXRECORDCOUNT
			'RESPONSE.WRITE(SQLST)
			RS.Open SQLST, Conn, adOpenStatic,adLockReadOnly, adCmdText
	
			if RS.EOF And RS.BOF then %>

<tr ID="FieldRow" CLASS="ResultRow" >
	<td COLSPAN="8" NOWRAP CLASS="ResultCell">No owners found re-check your criteria</td>
</tr>

<%			Else
				Do While Not RS.EOF
					RecCount = RecCount + 1
%>

<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" OID="<%=RS("OWNER_ID")%>">
   <td NOWRAP CLASS="ResultCell"><%=renderCell(RS("OWNER_ID"))%></td>
   <td NOWRAP CLASS="ResultCell"><%=renderCell(RS("NAME_TITLE"))%></td>
   <td NOWRAP CLASS="ResultCell"><%=renderCell(RS("NAME_LAST"))%>-<%=renderCell(RS("NAME_FIRST"))%></td>
   <td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ADDRESS_LINE1"))%></td>
   <td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ADDRESS_CITY"))%></td>
   <td NOWRAP CLASS="ResultCell"><%=renderCell(RS("ADDRESS_STATE"))%></td>
   <td NOWRAP CLASS="ResultCell"><%=renderCell(RS("PHONE_WORK"))%></td>
   <td NOWRAP CLASS="ResultCell"><%=renderCell(RS("PHONE_FAX"))%></td>
  
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
