<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"--> 
<%
Response.Expires = 0
%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE></TITLE>
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT LANGUAGE=JavaScript>
<!--
function getpolicyid( objRow ) {
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	return "POLICY_ID=" + objRow.getAttribute("POLICYID") + "&AHSID=" + objRow.getAttribute("AHSID")
}

function SelectRow() {
	lret = getselectedindex(document.all.tblResult )
	if (lret != -1)
	{
	return getpolicyid(document.all.tblResult.rows(lret))
	}
	else
	{
	return -1
	}
}

function CopyItem()
{
ID = getselectedindex(document.all.tblResult )
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
	ClipboardAgent.AddProperty "POLICY_ID", ID
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
<TD NOWRAP CLASS=thd>AHSID</TD>
<TD NOWRAP CLASS=thd>Name</TD>
<TD NOWRAP CLASS=thd>Policy ID</TD>
<TD NOWRAP CLASS=thd>Policy Number</TD>
<TD NOWRAP CLASS=thd>Policy Description</TD>
<TD NOWRAP CLASS=thd>Effective Date</TD>
<TD NOWRAP CLASS=thd>Expiration Date</TD>
<TD NOWRAP CLASS=thd>LOB</TD>
</TR>
</THEAD>
	<tbody ID="TableRows">
<%
RecCount = -1
If Request.Form <> "" Then
	RecCount = 0		' Needs to convert into 4 Single-Quotes for the Dynamic SQL in order to process (').
				POLICY_NUMBER = "'" & Replace(Request.Form("POLICY_NUMBER"), "'", "''''") & "'"
				EFFECTIVE_DATE = "'" & Replace(Request.Form("EFFECTIVE_DATE"), "'", "''''") & "'"
				EXPIRATION_DATE = "'" & Replace(Request.Form("EXPIRATION_DATE"), "'", "''''") & "'"
				LOB_CD = "'" &  Replace(Request.Form("LOB_CD"), "'", "''''") & "'"
				POLICY_DESC = "'" & Replace(Request.Form("POLICY_DESC"), "'", "''''") & "'"
				POLICY_ID = "''"
				inAtAHSID = Request.Form("AT_AHSID")
				

	Select Case Request.Form("SearchType")
		Case "B"
			SEARCHTYPE = 1
		Case "C"
			SEARCHTYPE = 2
		Case "E"
			SEARCHTYPE = 3
		Case Else
			SEARCHTYPE = 1	
	End Select
	
	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
	QSQL = ""
	
	If Request.Form("SearchDirection") = "UP" Then
		QSQL = QSQL & "{call Designer_2.SrchPolicyUpTree( "
	Else
		QSQL = QSQL & "{call Designer_2.SrchPolicyDownTree( "
	End If


	USER_ID = "null"
	If Not IsEmpty(Session("ACCOUNT_SECURITY")) Then USER_ID = Session("SecurityObj").m_UserID

	QSQL = QSQL & inAtAHSID & ", "
	QSQL = QSQL & POLICY_ID & ", "
	QSQL = QSQL & LOB_CD & ", "
	QSQL = QSQL & POLICY_NUMBER & ", "
	QSQL = QSQL & POLICY_DESC & ", "
	QSQL = QSQL & EFFECTIVE_DATE & ", "
	QSQL = QSQL & EXPIRATION_DATE & ", "
	QSQL = QSQL & USER_ID & ", "
	QSQL = QSQL & SEARCHTYPE & ", "
	QSQL = QSQL & "{resultset 1000, outRelatedAHSID, outRelatedAcctName, outPID, outPolicyNum, outEffectiveDate, outExpireDate, outLOB, outPolicyDesc})}"

	Set RSSearch = Server.CreateObject("ADODB.RecordSet")
	RSSearch.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
	RSSearch.Open QSQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
	
	If RSSearch.EOF AND RSSearch.BOF Then 
%>
<TR ID="FieldRow" CLASS=RESULTROW POLICYID='' AHSID=''>
<TD CLASS=LABEL COLSPAN=9>No Policies Found</TD>
</TR>

<%
	Else
	Do While Not RSSearch.EOF
%>
<TR ID="FieldRow" CLASS=RESULTROW OnClick="Javascript:multiselect(this);" POLICYID='<%= RSSearch("outPID") %>' AHSID='<%= RSSearch("outRelatedAHSID") %>'>
<TD NOWRAP CLASS=ResultCell><%= renderCell(RSSearch("outRelatedAHSID")) %></TD>
<TD NOWRAP CLASS=ResultCell><%= renderCell(RSSearch("outRelatedAcctName")) %></TD>
<TD NOWRAP CLASS=ResultCell><%= renderCell(RSSearch("outPID")) %></TD>
<TD NOWRAP CLASS=ResultCell><%= renderCell(RSSearch("outPolicyNum")) %></TD>
<TD NOWRAP CLASS=ResultCell><%= renderCell(RSSearch("outPolicyDesc")) %></TD>
<TD NOWRAP CLASS=ResultCell><%= renderCell(RSSearch("outEffectiveDate")) %></TD>
<TD NOWRAP CLASS=ResultCell><%= renderCell(RSSearch("outExpireDate")) %></TD>
<TD NOWRAP CLASS=ResultCell><%= renderCell(RSSearch("outLOB")) %></TD>
</TR>
<% 
RSSearch.MoveNext
Loop
End If
%>
</tbody>
</TABLE>
</DIV>
<SCRIPT LANGUAGE=VBSCRIPT>
<% If RSSearch.RecordCount = MAXRECORDCOUNT Then %>
	Parent.frames("TOP").document.all.spanstatus.innerhtml = "<%= MSG_MAXRECORDS %>"
<% Elseif RSSearch.RecordCount = 1000 Then %>
	Parent.frames("TOP").document.all.spanstatus.innerhtml = "Query Returned More Than 1000 Records, the First 1000 Records Are Listed."
<% Else %>
	Parent.frames("TOP").document.all.spanstatus.innerhtml = "Record count is <%= RSSearch.RecordCount %>"
<% End If  %>
</SCRIPT>
<% 
RSSearch.close
End If 
%>
</BODY>
</HTML>
