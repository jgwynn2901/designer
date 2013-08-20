<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\AHSTree.inc"-->
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE></TITLE>
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT LANGUAGE=JavaScript>
<!--
function getahsid(objRow) 
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	return objRow.getAttribute("AHSID")
}

function getahsidName(objRow) 
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	return objRow.cells("NAME").innerText;
}

function ExeAddNode(inWhichList)
{
	var strName, strAHSID;
	lret = getselectedindex(document.all.tblResult);
	if (lret != -1)
	{
		curRow = document.all.tblResult.rows(lret);
		strAHSID = getahsid(curRow);
		strName =  getahsidName(curRow);
		parent.AddNode(strAHSID,strName,inWhichList);
	}
	else
		alert("Please select a row");

}
function ExeClear()
{
	
	for (i=0; i < document.all.tblResult.rows; i++)
		document.all.tblResult.rows(i).remove;
}

//-->
</script>
<SCRIPT LANGUAGE="JavaScript" FOR="FilterBtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
		case "ATTACHBUTTONCLICK":
				ExeAddNode("INCLUDE");
			break;
		case "REMOVEBUTTONCLICK":
				ExeAddNode("EXCLUDE");
			break;
		default:
			break;
	}
   
</SCRIPT>

</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'92%';width:'100%'">
<OBJECT data="../Scriptlets/ObjButtons.asp?REMOVECAPTION=Exclude&ATTACHCAPTION=Include&HIDEREFRESH=TRUE&HIDENEW=TRUE&HIDESEARCH=TRUE&HIDECOPY=TRUE&HIDEPASTE=TRUE&HIDEEDIT=TRUE" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id=FilterBtnControl type=text/x-scriptlet></OBJECT>
<DIV align="LEFT" id="Account_RESULTS" style="display:block;height:'80%';width:'100%';overflow:auto">
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
dim cID, nRecCount

If Request.Form <> "" then
	Select Case Request.Form("SEARCHTYPE")
		Case "B"
			NAME = Replace(Request.Form("NAME"), "'", "''") & "%"
			FNS_CLIENT_CD = Replace(Request.Form("FNS_CLIENT_CD"), "'", "''") & "%"
			CITY = Replace(Request.Form("CITY"), "'", "''") & "%"
			STATE = Replace(Request.Form("STATE"), "'", "''") & "%"
			ZIP = Replace(Request.Form("ZIP"), "'", "''") & "%"
			ADDRESS_1 = Replace(Request.Form("ADDRESS_1"), "'", "''") & "%"
			'Added for PSUS-0008
			'Adding a new Search filter AHS ID and Location Code under a client node.
			'Prashant Shekhar 04/17/2007
			LOCATION_CODE = Replace(Request.Form("LOCATION_CODE"), "'", "''") & "%"
			ACCNT_HRCY_STEP_ID = Replace(Request.Form("ACCNT_HRCY_STEP_ID"), "'", "''") & "%"
			
		Case "C"
			NAME = "%" & Replace(Request.Form("NAME"), "'", "''") & "%"
			FNS_CLIENT_CD = "%" & Replace(Request.Form("FNS_CLIENT_CD"), "'", "''") & "%"
			CITY = "%" & Replace(Request.Form("CITY"), "'", "''") & "%"
			STATE = "%" & Replace(Request.Form("STATE"), "'", "''") & "%"
			ZIP = "%" & Replace(Request.Form("ZIP"), "'", "''") & "%"
			ADDRESS_1 = "%" & Replace(Request.Form("ADDRESS_1"), "'", "''") & "%"
			'Added for PSUS-0008
			'Adding a new Search filter AHS ID and Location Code under a client node.
			'Prashant Shekhar 04/17/2007

			LOCATION_CODE = "%" & Replace(Request.Form("LOCATION_CODE"), "'", "''") & "%"
			ACCNT_HRCY_STEP_ID = "%" & Replace(Request.Form("ACCNT_HRCY_STEP_ID"), "'", "''") & "%"
			
		Case "E"
			NAME = Replace(Request.Form("NAME"), "'", "''")
			FNS_CLIENT_CD = Replace(Request.Form("FNS_CLIENT_CD"), "'", "''")
			CITY = Replace(Request.Form("CITY"), "'", "''")
			STATE = Replace(Request.Form("STATE"), "'", "''")
			ZIP = Replace(Request.Form("ZIP"), "'", "''")
			ADDRESS_1 = Replace(Request.Form("ADDRESS_1"), "'", "''")
			'Added for PSUS-0008
			'Adding a new Search filter AHS ID and Location Code under a client node.
			'Prashant Shekhar 04/17/2007

			LOCATION_CODE = Replace(Request.Form("LOCATION_CODE"), "'", "''")
			ACCNT_HRCY_STEP_ID = Replace(Request.Form("ACCNT_HRCY_STEP_ID"), "'", "''")
	End Select

	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
	SQLWHERE = ""
	
	SQL = "SELECT * FROM ACCOUNT_HIERARCHY_STEP WHERE PARENT_NODE_ID = " & Request.Form("AHSID") & " AND "
	
	If Request.Form("AHSID") = "1" Then
		If Not IsEmpty(Session("ACCOUNT_SECURITY")) Then SQL = "SELECT * FROM ACCOUNT_HIERARCHY_STEP AHS WHERE ACCNT_HRCY_STEP_ID IN(" & CStr(Session("ACCOUNT_SECURITY")) & ") AND " 
	End If

	If Request.Form("NAME") <> "" Then 
		SQLWHERE = SQLWHERE & "Upper(NAME) LIKE '" & Ucase(NAME) & "'"
	End If
	
	If Request.Form("FNS_CLIENT_CD") <> "" Then 
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " AND "
		End If
		SQLWHERE = SQLWHERE & "Upper(FNS_CLIENT_CD) LIKE '" & Ucase(FNS_CLIENT_CD) & "'"
	End If

	'Added for PSUS-0008
	'Adding a new Search filter AHS ID and Location Code under a client node.
	'Prashant Shekhar 04/17/2007

	If Request.Form("LOCATION_CODE") <> "" Then 
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " AND "
		End If
		SQLWHERE = SQLWHERE & "Upper(LOCATION_CODE) LIKE '" & Ucase(LOCATION_CODE) & "'"
	End If
	If Request.Form("ACCNT_HRCY_STEP_ID") <> "" Then 
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " AND "
		End If
		SQLWHERE = SQLWHERE & "Upper(ACCNT_HRCY_STEP_ID) LIKE '" & Ucase(ACCNT_HRCY_STEP_ID) & "'"
	End If
	
	If Request.Form("CITY") <> "" Then 
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " AND "
		End If
		SQLWHERE = SQLWHERE & "Upper(CITY) LIKE '" & Ucase(CITY) & "'"
	End If
	
	If Request.Form("STATE") <> "" Then 
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " AND "
		End If
		SQLWHERE = SQLWHERE & "Upper(STATE) LIKE '" & Ucase(STATE) & "'"
	End If
	
	If Request.Form("ZIP") <> "" Then 
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " AND "
		End If
		SQLWHERE = SQLWHERE & "Upper(ZIP) LIKE '" & Ucase(ZIP) & "'"
	End If

	If Request.Form("ADDRESS_1") <> "" Then 
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " AND "
		End If
		SQLWHERE = SQLWHERE & "Upper(ADDRESS_1) LIKE '" & Ucase(ADDRESS_1) & "'"
	End If
	
	SQL = SQL & SQLWHERE & " ORDER BY NAME"
	RS.Open SQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
	nRecCount = 0
	If RS.EOF AND RS.BOF Then 
%>
<TR ID="FieldRow" CLASS=RESULTROW POLICYID='' AHSID=''>
<TD CLASS=LABEL COLSPAN=5>No Business Entity Found</TD>
</TR>
<%
	Else
		Do While Not RS.EOF
		%>
			<TR ID="FieldRow" CLASS=RESULTROW OnClick="Javascript:multiselect(this);" AHSID='<%= RS("ACCNT_HRCY_STEP_ID") %>'>
			<TD CLASS=LABEL><%= renderCell(RS("ACCNT_HRCY_STEP_ID")) %></TD>
			<TD CLASS=LABEL ID="NAME"><%= renderCell(RS("NAME")) %></TD>
			<TD CLASS=LABEL><%= renderCell(RS("CITY")) %></TD>
			<TD CLASS=LABEL><%= renderCell(RS("STATE")) %></TD>
			<TD CLASS=LABEL><%= renderCell(RS("ZIP")) %></TD>
			</TR>
		<% 
			nRecCount = 	nRecCount + 1
			RS.MoveNext
		Loop
	end if
%>
</tbody>
</TABLE>
</DIV>
</fieldset>
<SCRIPT LANGUAGE=VBSCRIPT>
<% If RS.RecordCount = MAXRECORDCOUNT Then %>
	Parent.frames("TOP").document.all.spanstatus.innerhtml = "<%= MSG_MAXRECORDS %>"
<% Else %>
	Parent.frames("TOP").document.all.spanstatus.innerhtml = "Record count is <%= RS.RecordCount %>"
<% End If%>
</SCRIPT>
<% 
End If 
%>
</BODY>
</HTML>
