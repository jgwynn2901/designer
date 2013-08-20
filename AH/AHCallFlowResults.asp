<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"--> 
<%
Function Swap(Indata)
If InData = "" Then
	Swap = "null"
Else
	Swap = InData
End If
End Function

	dim RecCount
	RecCount = -1
If Request.Form <> "" Then
	RecCount = 0
		Select Case Request.Form("SEARCHTYPE")
			Case "B"
				SEARCHTYPE = 1
			Case "C"
				SEARCHTYPE = 2
			Case "E"
				SEARCHTYPE = 3
		End Select
	If Request.Form("SearchDirection") = "UP" Then
		SQL = SQL & "{call Designer_2.SrchCallFlowUpTree("
	Else
		SQL = SQL & "{call Designer_2.SrchCallFlowDownTree( "
	End If
	' Needs to convert into 4 Single-Quotes for the Dynamic SQL in order to process (').
	DESCRIPTION = "'" & Replace(Request.Form("DESCRIPTION"), "'" , "''''") & "'"
	NAME = "'" & Replace(Request.Form("NAME"), "'" , "''''") & "'"
	LOB_CD = "'" & Replace(Request.Form("LOB_CD"), "'" , "''''") & "'"
	CALLFLOW_ID = Replace(Request.Form("CALLFLOW_ID"), "'" , "''''")
	
	USER_ID = "null"
	If Not IsEmpty(Session("ACCOUNT_SECURITY")) Then USER_ID = Session("SecurityObj").m_UserID
	
	SQL = SQL & Request.Form("AT_AHSID") & ", "
	SQL = SQL & Swap(CALLFLOW_ID) & ", "
	SQL = SQL & LOB_CD & ", "
	SQL = SQL & NAME & ", "
	SQL = SQL & DESCRIPTION & ", "
	SQL = SQL & USER_ID & ", "	
	SQL = SQL & SEARCHTYPE & ", "
	SQL = SQL &"{resultset 10000, outRelatedAHSID, outRelatedAcctName, outCFID, outLOB, outName, outDesc})}"

	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
	RS.Open SQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
End If
%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="javascript">
<!--
var lastObject;
var currentObject;
var firsttime = 1;
var arraySelectedObjects= new Array( );
var currentRowIndex = 0;
var bHighlighted = 0;
function highlight( object )
{
	if( !firsttime && lastobject) 
	{
		lastobject.className = 'ResultRow';
	}

	if( object.rowIndex == currentRowIndex )
	{
		if( bHighlighted == 0)
		{
			object.className = 'ResultSelectRow';
			lastobject = object;
			firsttime = 0;
			currentRowIndex = object.rowIndex;
			bHighlighted = 1;
		}
		else
		{
			object.className = 'ResultRow';
			currentRowIndex = 1;
			bHighlighted = 0;
		}
	}
	else
	{
		object.className = 'ResultSelectRow';
		lastobject = object;
		firsttime = 0;
		currentRowIndex = object.rowIndex;
		bHighlighted = 1;
	}
	object.scrollIntoView(true);
}

function relativemultiselect( objTable, updown )
{
	var i;

	for(i=1;i<objTable.rows.length;i++)
	{
		if( objTable.rows[i].className == 'ResultSelectRow' )
		{
		
			if ((1 == updown) && (i < objTable.rows.length - 1))
			{
				multiselect( objTable.rows[i + 1] );
				obj = objTable.rows[i + 1];
				obj.scrollIntoView(true);
				window.event.keyCode = 0;
				window.event.returnValue = 0;
			}
			else if ((-1 == updown) && (i > 1))
			{
				multiselect( objTable.rows[i - 1] );
				obj = objTable.rows[i - 1];
				obj.scrollIntoView(true);
				window.event.keyCode = 0;
				window.event.returnValue = 0;
			}
			i = objTable.rows.length;
		}		
	}

}
function getselectedindex( objTable )
{

	for(i=1;i<objTable.rows.length;i++)
	{
		if( objTable.rows[i].className == 'ResultSelectRow' )
		{
			return dblhighlight(objTable.rows[i]);
			
		}		
	}
	return -1;
}

function getindex(objTable)
{
	for(i=1;i<objTable.rows.length;i++)
	{
		if( objTable.rows[i].className == 'ResultSelectRow' )
		{
			return i;
		}		
	}
	return -1;

}

function getmultipleindex( objTable )
{
var ids
ids = "";
	for(i=1;i<objTable.rows.length;i++)
	{
		if( objTable.rows[i].className == 'ResultSelectRow' )
		{
			if (ids != "") 
			{
			ids = ids + ",";
			}
			ids = ids + dblhighlight(objTable.rows[i]) 
		}		
	
	}
	return ids;
}


function selectall(objTable)
	{
	arraySelectedObjects.length=0;
	for(i=1;i<objTable.rows.length;i++)
		{
		objTable.rows[i].className = 'ResultSelectRow';
		arraySelectedObjects[i-1]=objTable.rows[i];
		}		
	}

function numselected()
	{
	return arraySelectedObjects.length;
	}

function clearselection()
	{
	for(i=0;i<arraySelectedObjects.length;i++)
		{
		arraySelectedObjects[i].className='ResultRow';
		}		
	arraySelectedObjects.length=0;
	}

function selectedrownum(rownum)
	{
	return arraySelectedObjects[rownum].rowIndex;
	}


function compareRows( a,b )
{
	return a.rowIndex - b.rowIndex;
	
}
// We will enter the rows into the array in order.
function multiselect( object )
	{
	if(window.event == null || window.event.ctrlKey==0)
		{
		clearselection();
		object.className='ResultSelectRow';
		arraySelectedObjects[0]=object;
		}
	else
		{
		if(object.className!='ResultSelectRow')
			{
			object.className='ResultSelectRow';
			
			arraySelectedObjects[arraySelectedObjects.length]=object;
			arraySelectedObjects.sort( compareRows );
			}
		}
	currentRowIndex = object.rowIndex;
	lastObject = object;
	}

function unhighlight( object )
{
	object.className = 'ResultRow';
}
function dblhighlight( objRow )
{
var qry
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	return objRow.getAttribute("RPID");
}

function SelectRow() {
	RPID = getselectedindex(document.all.tblResult )
	index = getindex(document.all.tblResult)
	if (RPID != "-1")
	{
		return "CFID=" + RPID + "&AHSID=" + document.all.tblResult.rows(index).getAttribute("AHSID")
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
-->
</script>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onload
<% If Request.Form <> "" Then %>
		if 0 < Document.all.tblResult.rows.length then
			call multiselect( Document.all.tblResult.rows(1))
		end if
<%End if %>
End Sub

Sub document_onkeydown
	select case window.event.keycode
		case 8:
			window.event.keyCode = 0
			window.event.returnValue = 0
		case 38:
			call relativemultiselect( Document.all.tblResult, -1 )
		case 40:
			call relativemultiselect( Document.all.tblResult, 1 )
		case 13:
			i = getselectedindex( Document.all.tblResult )
			if 0 < i then
				dblhighlight(Document.all.tblResult.rows(i))
			end if
		case else:
	end select
End Sub

function MakeCopy(ID)
If Not IsNull(ID) Then
	ClipboardAgent.ClearAllProperties()
	ClipboardAgent.AddProperty "CALLFLOW_ID", ID
	ClipboardAgent.SetPropertiesToClipboard()
End If
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
<BODY  bgcolor='<%= BODYBGCOLOR %>'  leftmargin=2 topmargin=2 rightmargin=0 bottommargin=2>
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblResult" name="tblResult" width="100%" RPID='X'>
<THEAD>
<TR>
<TD CLASS=ResultHeader>AHSID</TD>
<TD CLASS=ResultHeader>Name</TD>
<TD CLASS=ResultHeader>CFID</TD>
<TD CLASS=ResultHeader>Callflow Name</TD>
<TD CLASS=ResultHeader>Description</TD>
<TD CLASS=ResultHeader>LOB</TD>
</TR>
</THEAD>
<%
If Request.Form <> "" Then
	Do While Not RS.EOF
	RecCount = RecCount + 1
%>
<TR CLASS=RESULTROW OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);" RPID='<%= RS("outCFID") %>' AHSID='<%= RS("outRelatedAHSID") %>'>
<TD CLASS=ResultCell><NOBR><%= RenderCell(RS("outRelatedAHSID")) %></TD>
<TD CLASS=ResultCell><NOBR><%= RenderCell(RS("outRelatedAcctName")) %></TD>
<TD CLASS=ResultCell><NOBR><%= RenderCell(RS("outCFID")) %></TD>
<TD CLASS=ResultCell><NOBR><%= RenderCell(RS("outName")) %></TD>
<TD CLASS=ResultCell><NOBR><%= RenderCell(RS("outDesc")) %></TD>
<TD CLASS=ResultCell><NOBR><%= RenderCell(RS("outLOB")) %></TD>
</TR>
<%
RS.movenext
Loop
If RS.EOF AND RS.BOF Then 
%>
<TR>
<TD CLASS=LABEL COLSPAN=7>No Routing Plans Found</TD>
</TR>
<%
End If
RS.Close
End If 
%>
</TABLE>
</DIV>
</FIELDSET>
<% If Request.Form <> "" Then %>
<SCRIPT LANGUAGE="VBScript">
if Parent.frames("TOP").document.readyState = "complete" then
	curCount = <%=RecCount%>
	if curCount = <%=MAXRECORDCOUNT%> then
		Parent.frames("TOP").UpdateStatus("<%=MSG_MAXRECORDS%>")
	elseif curCount = -1 then
		Parent.frames("TOP").UpdateStatus("<%=MSG_PROMPT%>")
	else		
		Parent.frames("TOP").UpdateStatus("Record count is <%=RecCount%>")
	end if		
end if
</SCRIPT>
<% End If %>
</BODY>
</HTML>
