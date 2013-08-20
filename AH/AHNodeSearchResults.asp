<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"--> 
<%
Server.ScriptTimeout = "6000"
Function Swap(InData)
	If InData <> "" Then
		Swap = "'" & InData & "'"
	Else
		Swap = "null"
	End If

End Function
%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE></TITLE>
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT LANGUAGE=JavaScript>
<!--
function getahsid( objRow ) {
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	return "AHSID=" + objRow.getAttribute("AHSID")
}

function SelectRow() {
	lret = getselectedindex(document.all.tblResult )
	if (lret != -1)
	{
	return getahsid(document.all.tblResult.rows(lret))
	}
	else
	{
	return -1
	}
}


function CopyItem()
{

lret = getselectedindex(document.all.tblResult )
	if (lret == "-1")
	{
	return -1;
	}
ID = getahsid(document.all.tblResult.rows(lret))
if (ID == "-1")
	{
	return -1;
	}
else
{
	MakeCopy(ID)
}
}
//-->
</script>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onload

End Sub

function MakeCopy(ID)
	ClipboardAgent.ClearAllProperties()
	ClipboardAgent.AddProperty "AHSID", ID
	ClipboardAgent.SetPropertiesToClipboard()
End Function
-->
</SCRIPT>
<OBJECT ID="ClipboardAgent" 
<%GetClipboardCLSID()%>
width=1 height=1>
<PARAM NAME="MaxPropertiesStringLength" VALUE="1000">
<PARAM NAME="MaxPropertyNameLength" VALUE="50">
<PARAM NAME="MaxPropertyValueLength" VALUE="200">
<PARAM NAME="NameValueDelimiter" VALUE="#">
<PARAM NAME="PropertyItemDelimiter" VALUE="|">
<PARAM NAME="PrivateClipboardFormatName" VALUE="CF_FNSDESIGNER">
</OBJECT>

</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<DIV align="LEFT" id="Account_RESULTS" style="display:block;height:240;width:'100%';overflow:scroll">
<table cellPadding=2 cellSpacing=0 rules=all ID="tblResult" name="tblResult" width=100%>
<thead CLASS="ResultHeader">
<TR>
<TD CLASS=ResultHeader>AHSID</TD>
<TD CLASS=ResultHeader>Name</TD>
<TD CLASS=ResultHeader>City</TD>
<TD CLASS=ResultHeader>State</TD>
<TD CLASS=ResultHeader>Zip</TD>
</TR>
</THEAD>
	<tbody ID="TableRows">
<%
dim oRS
dim cSQL

RecCount = -1
If Request.Form <> "" Then
	RecCount = 0	' Needs to convert into 4 Single-Quotes for the Dynamic SQL in order to process (').
				NAME = Replace(Request.Form("NAME"), "'", "''''")
				LOCATION_CODE = Replace(Request.Form("LOCATION_CODE"), "'", "''''")
				'Added for PSUS-0008
				'Adding a new Search filter AHS ID under a client node.
				'Prashant Shekhar 04/16/2007
				ACCNT_HRCY_STEP_ID = Replace(Request.Form("ACCNT_HRCY_STEP_ID"), "'", "''''")
				CITY = Replace(Request.Form("CITY"), "'", "''''")
				STATE = Replace(Request.Form("STATE"), "'", "''''")
				ZIP = Replace(Request.Form("ZIP"), "'", "''''")
				ADDRESS_1 = Replace(Request.Form("ADDRESS_1"), "'", "''''")
				AT_AHSID = Request.Form("AT_AHSID")
				Select Case Request.Form("SEARCHTYPE")
					Case "B"
						SEARCHTYPE = 1
					Case "C"
						SEARCHTYPE = 2
					Case "E"
						SEARCHTYPE = 3
				End Select
	cSQL = ""
	If Request.Form("SEARCHDIRECTION") = "UP" Then
		cSQL = cSQL & "{call Designer_2.SrchAHSNodesUpTree("
	Else
		cSQL = cSQL & "{call Designer_2.SrchAHSNodesDownTree("
	End If
	
	USER_ID = Session("SecurityObj").m_UserID
	
	cSQL = cSQL & AT_AHSID & ", "
	cSQL = cSQL & "null, "
	cSQL = cSQL & Swap(NAME) & ", " 
	cSQL = cSQL & Swap(ADDRESS_1) & ", "
	cSQL = cSQL & Swap(CITY) & "," 
	cSQL = cSQL & Swap(STATE) & ","
	cSQL = cSQL & Swap(ZIP) & ","
	cSQL = cSQL & Swap(LOCATION_CODE) & ","  
	
	'Added for PSUS-0008
	'Adding a new Search filter AHS ID under a client node.
	'Prashant Shekhar 04/16/2007
	cSQL = cSQL & Swap(ACCNT_HRCY_STEP_ID) & ","
	cSQL = cSQL & USER_ID & ","  
	cSQL = cSQL & SEARCHTYPE 
	cSQL = cSQL & ",{resultset 1000, outAHSID, outAHStName, outCity, outState, outZip, outLevel})}"
	
	Set oRS = Server.CreateObject("ADODB.Recordset")
	oRS.MaxRecords = MAXRECORDCOUNT
	oRS.Open cSQL, CONNECT_STRING, adOpenStatic, adLockReadOnly, adCmdStoredProc
	If oRS.EOF AND oRS.BOF Then 
%>
<TR ID="FieldRow" CLASS=RESULTROW POLICYID='' AHSID=''>
<TD CLASS=LABEL COLSPAN=5>No Business Entity Found</TD>
</TR>
<%
	Else
	Do While Not oRS.EOF
%>
<TR ID="FieldRow" CLASS=RESULTROW OnClick="Javascript:multiselect(this);" AHSID='<%= oRS("outAHSID") %>'>
<TD CLASS=ResultCell><%= renderCell(oRS("outAHSID")) %></TD>
<TD CLASS=ResultCell><%= renderCell(oRS("outAHStName")) %></TD>
<TD CLASS=ResultCell><%= renderCell(oRS("outCity")) %></TD>
<TD CLASS=ResultCell><%= renderCell(oRS("outState")) %></TD>
<TD CLASS=ResultCell><%= renderCell(oRS("outZip")) %></TD>
</TR>
<% 
oRS.MoveNext
Loop
End If
%>
</tbody>
</TABLE>
</DIV>
<SCRIPT LANGUAGE=VBSCRIPT>
<% If oRS.RecordCount = MAXRECORDCOUNT Then %>
	Parent.frames("TOP").document.all.spanstatus.innerhtml = "<%= MSG_MAXRECORDS %>"
<% ElseIf oRS.RecordCount = 1000 Then %>
	Parent.frames("TOP").document.all.spanstatus.innerhtml = "Query Returned More Than 1000 Records, the First 1000 Records Are Listed."
<% Else %>
	Parent.frames("TOP").document.all.spanstatus.innerhtml = "Record count is <%= oRS.RecordCount %>"
<% End If  %>
</SCRIPT>
<% 
oRS.Close
set oRS = nothing
End If 
%>
</BODY>
</HTML>
