<%
	Response.Expires = 0
	Response.Buffer = true
%>

<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>User Search Results</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub window_onload
	if Parent.frames("TOP").document.readyState = "complete" then
		Parent.frames("TOP").document.all.PID.value = ""
	end if
End Sub

Function GetPID
	GetPID = getmultipleindex(document.all.tblFields, "PID")
End Function

Function GetLOB
	GetLOB = getmultipleindex(document.all.tblFields, "LOB")
End Function

</script>

<script language=jscript>
/*	function override	-	no multiple selection	*/
function multiselect( object )
{
	clearselection();
	object.className='ResultSelectRow';
	arraySelectedObjects[0]=object;
	currentRowIndex = object.rowIndex;
	lastObject = object;
}

function dblhighlight(objRow, whichCol)
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	if (whichCol == "PID")
		return objRow.getAttribute("PID");
	else if (whichCol == "LOB")
		return objRow.getAttribute("LOB");
}
</SCRIPT>

</head>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Policy Id</div></td>
			<td class="thd"><div id><nobr>Number</div></td>
			<td class="thd"><div id><nobr>Description</div></td>
			<td class="thd"><div id><nobr>LOB</div></td>
			<td class="thd"><div id><nobr>Effective</div></td>
			<td class="thd"><div id><nobr>Expiration</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<%
	dim nRecCount, lExactSearch, cSQL, cWhere, oConn, oRS
	
	nRecCount = -1
	If Request.QueryString("SEARCHTYPE") <> "" Then
	nRecCount = 0

	EFFECTIVE         = Request.QueryString("SearchEffective")
	ORIGINALEFFECTIVE = Request.QueryString("SearchOriginalEffective")
	EXPIRATION        = Request.QueryString("SearchExpiration")

	lExactSearch = false
	Select Case Request.QueryString("SEARCHTYPE")
		Case "B"
			PID          = Request.QueryString("SearchPID") & "%"
			NUMBER       = Request.QueryString("SearchNumber") & "%"
			AHSID        = Request.QueryString("SearchAHSID") 
			CARRIER      = Request.QueryString("SearchCarrier") & "%"
			TPADMIN		 = Request.QueryString("SearchTPADMIN") & "%"
			AGENT        = Request.QueryString("SearchAgent") & "%"
			LOBCD        = Request.QueryString("SearchLOBCD")
			MCTYPE       = Request.QueryString("SearchMCTYPE")
			SELFINSUREDFLG = Request.QueryString("SearchSelfInsuredFlg") & "%"
			COMPANYCODE    = Request.QueryString("SearchCompanyCode") & "%"
		Case "C"
			PID         = "%" & Request.QueryString("SearchPID") & "%"
			NUMBER      = "%" & Request.QueryString("SearchNumber") & "%"
			AHSID       =  Request.QueryString("SearchAHSID") 
			CARRIER     = "%" & Request.QueryString("SearchCarrier") & "%"
			TPADMIN		 = "%" & Request.QueryString("SearchTPADMIN") & "%"
			AGENT       = "%" & Request.QueryString("SearchAgent") & "%"
			LOBCD       = Request.QueryString("SearchLOBCD")
			MCTYPE      = Request.QueryString("SearchMCTYPE")
			SELFINSUREDFLG = "%" & Request.QueryString("SearchSelfInsuredFlg") & "%"
			COMPANYCODE = "%" & Request.QueryString("SearchCompanyCode") & "%"
		Case "E"
			lExactSearch = true
			PID         = Request.QueryString("SearchPID")
			NUMBER      = Request.QueryString("SearchNumber")
			AHSID       = Request.QueryString("SearchAHSID")
			CARRIER     = Request.QueryString("SearchCarrier")
			TPADMIN		= Request.QueryString("SearchTPADMIN")
			AGENT       = Request.QueryString("SearchAgent")
			LOBCD       = Request.QueryString("SearchLOBCD")
			MCTYPE      = Request.QueryString("SearchMCTYPE")
			SELFINSUREDFLG = Request.QueryString("SearchSelfInsuredFlg")
			COMPANYCODE    = Request.QueryString("SearchCompanyCode")
	End Select
	PID         = Replace(PID,"'","''")
	NUMBER      = UCase(Replace(NUMBER,"'","''"))
	AHSID       = Replace(AHSID,"'","''")
	CARRIER     = Replace(CARRIER,"'","''")
	AGENT       = Replace(AGENT,"'","''")
	LOBCD       = UCase(Replace(LOBCD,"'","''"))
	MCTYPE      = Replace(MCTYPE,"'","''")
	SELFINSUREDFLG = Replace(SELFINSUREDFLG,"'","''")
	COMPANYCODE    = UCase(Replace(COMPANYCODE,"'","''"))
		
	if lExactSearch then
		cVerb = "="
	else
		cVerb = "LIKE"
	end if
		
	cSQL = "SELECT /*+ USE_HASH(P) */ P.POLICY_ID, P.POLICY_NUMBER, P.POLICY_DESC, " &_
			"to_char(P.EXPIRATION_DATE, 'MM/DD/YYYY') F_EXPIRATION_DATE, " &_
			"To_Date(TO_CHAR(p.effective_date, 'MM/DD/YYYY'),'MM/DD/YYYY') F_EFFECTIVE_DATE, " &_
			"AHSP.ACCNT_HRCY_STEP_ID, AHSP.LOB_CD, " &_
			"AHSP.AHS_POLICY_ID " &_
			"FROM AHS_POLICY AHSP, POLICY P "
	cWhere = "WHERE AHSP.POLICY_ID = P.POLICY_ID + 0 "
		
	If Request.QueryString("SearchNumber") <> "" Then
	    cWhere = cWhere & " AND P.POLICY_NUMBER " & cVerb & " '" & NUMBER & "'"
	End If
	If Request.QueryString("SearchPID") <> "" Then
		cWhere = cWhere & " AND P.POLICY_ID " & cVerb & " '" & PID & "'"
	End If
	If Request.QueryString("SearchDescription") <> "" Then
		cWhere = cWhere & " AND P.POLICY_DESC " & cVerb & " '" & DESCRIPTION & "'"
	End If
	If Request.QueryString("SearchAHSID") <> "" Then
		cWhere = cWhere & " AND AHSP.ACCNT_HRCY_STEP_ID =" & AHSID 
	End If		
	If Request.QueryString("SearchCarrier") <> "" Then
		cWhere = cWhere & " AND P.CARRIER_ID " & cVerb & " '" & CARRIER & "'"
	End If
	If Request.QueryString("SearchTPADMIN") <> "" Then
		cWhere = cWhere & " AND P.TPA_ID " & cVerb & " '" & TPADMIN & "'"
	End If			
	If Request.QueryString("SearchAgent") <> "" Then
		cWhere = cWhere & " AND P.AGENT_ID " & cVerb & " '" & AGENT & "'"
	End If		
	If Request.QueryString("SearchLOBCD") <> "" Then
		cWhere = cWhere & " AND AHSP.LOB_CD " & cVerb & " '" & LOBCD & "'"
	End If
	If Request.QueryString("SearchMCTYPE") <> "" Then
		cWhere = cWhere & " AND P.MANAGED_CARE_TYPE " & cVerb & " '" & MCTYPE & "'"
	End If
	If Request.QueryString("SearchSelfInsuredFlg") <> "" Then
		cWhere = cWhere & " AND P.SELF_INSURED " & cVerb & " '" & SELFINSUREDFLG & "'"
	End If
	If Request.QueryString("SearchCompanyCode") <> "" Then
		cWhere = cWhere & " AND P.COMPANY_CODE " & cVerb & " '" & COMPANYCODE & "'"
	End If
		
	If Request.QueryString("SearchEffective") <> "" Then
		cWhere = cWhere & " AND (to_date(to_char(P.EFFECTIVE_DATE, 'MM-DD-YYYY'), 'MM-DD-YYYY') BETWEEN to_date('" & EFFECTIVE & "', 'MM-DD-YYYY') AND to_date('" & EFFECTIVE & "', 'MM-DD-YYYY'))"
	End If
	If Request.QueryString("SearchOriginalEffective") <> "" Then
		cWhere = cWhere & " AND (to_date(to_char(P.ORIGINAL_EFFECTIVE_DATE, 'MM-DD-YYYY'), 'MM-DD-YYYY') BETWEEN to_date('" & ORIGINALEFFECTIVE & "', 'MM-DD-YYYY') AND to_date('" & ORIGINALEFFECTIVE & "', 'MM-DD-YYYY'))"
	End If
	If Request.QueryString("SearchExpiration") <> "" Then
		cWhere = cWhere & " AND (to_date(to_char(P.EXPIRATION_DATE, 'MM-DD-YYYY'), 'MM-DD-YYYY') BETWEEN to_date('" & EXPIRATION & "', 'MM-DD-YYYY') AND to_date('" & EXPIRATION & "', 'MM-DD-YYYY'))"
	End If
	If Request.QueryString("SearchCancellation") <> "" Then
		cWhere = cWhere & " AND (to_date(to_char(P.CANCELLATION_DATE, 'MM-DD-YYYY'), 'MM-DD-YYYY') BETWEEN to_date('" & CANCELLATION & "', 'MM-DD-YYYY') AND to_date('" & CANCELLATION & "', 'MM-DD-YYYY'))"
	End If
	If Request.QueryString("SearchChange") <> "" Then
		cWhere = cWhere & " AND (to_date(to_char(P.CHANGE_DATE, 'MM-DD-YYYY'), 'MM-DD-YYYY') BETWEEN to_date('" & CHANGE & "', 'MM-DD-YYYY') AND to_date('" & CHANGE & "', 'MM-DD-YYYY'))"
	End If
	If Request.QueryString("SearchLoad") <> "" Then
		cWhere = cWhere & " AND (to_date(to_char(P.LOAD_DATE, 'MM-DD-YYYY'), 'MM-DD-YYYY') BETWEEN to_date('" & LOAD & "', 'MM-DD-YYYY') AND to_date('" & LOAD & "', 'MM-DD-YYYY'))"
	End If
	
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	
	cSQL = cSQL & cWhere & " ORDER BY  LOB_CD,F_EFFECTIVE_DATE DESC"
	Set oRS = Server.CreateObject("ADODB.Recordset")
	oRS.MaxRecords = MAXRECORDCOUNT
	oRS.Open cSQL, oConn, adOpenStatic, adLockReadOnly, adCmdText
	
	if oRS.EOF then %>

<tr ID="FieldRow" CLASS="ResultRow" PID=''>
	<td COLSPAN=7 NOWRAP CLASS="ResultCell" >No policies found re-check your criteria</td>
</tr>

<%	Else
		cLOB = ""
		cPolicyNo = ""
		Do While Not oRS.EOF
			if cLOB <> oRS("LOB_CD") OR cPolicyNo <> oRS("POLICY_NUMBER") then
				cLOB = oRS("LOB_CD")
				cPolicyNo = oRS("POLICY_NUMBER")
				nRecCount = nRecCount + 1
%>
<tr ID="FieldRow" CLASS="ResultRow" DYNKEY="1" OnClick="Javascript:multiselect(this);" LOB="<%=oRS("LOB_CD")%>" PID="<%=oRS("POLICY_ID")%>">
<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("POLICY_ID"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("POLICY_NUMBER"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("POLICY_DESC"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("LOB_CD"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("F_EFFECTIVE_DATE"))%></td>
<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS("F_EXPIRATION_DATE"))%></td>
</tr>
<%
			end if
			oRS.MoveNext
		Loop
	End If
	oRS.Close
	Set oRS = Nothing
	oConn.Close
	Set oConn = Nothing
End If
%>

</tbody>
</table>
</div>
</fieldset>
<SCRIPT LANGUAGE="VBScript">
<%	If nRecCount >= 0 Then %>
if Parent.frames("TOP").document.readyState = "complete" then
	curCount = <%=nRecCount%>
	if curCount = <%=MAXRECORDCOUNT%> then
		Parent.frames("TOP").UpdateStatus("<%=MSG_MAXRECORDS%>")
	else		
		Parent.frames("TOP").UpdateStatus("Record count is <%=nRecCount%>")
	end if		
end if
<%	End If %>
</SCRIPT>
</body>
</html>
