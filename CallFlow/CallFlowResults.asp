<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires=0
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
	
	DESCRIPTION = Replace(DESCRIPTION, "'" , "''")
	NAME = Replace(NAME, "'" , "''")
	LOB_CD = Replace(LOB_CD, "'" , "''")
	CALLFLOW_ID = Replace(CALLFLOW_ID, "'" , "''")
	FRAME_NAME = Replace(FRAME_NAME, "'" , "''")
	
''	Set Conn = Server.CreateObject("ADODB.Connection")
'	ConnectionString = CONNECT_STRING
''	Conn.Open ConnectionString
	
	SQLST = ""
	SQLWHERE = ""
	SQLST = SQLST & "SELECT CALLFLOW.* FROM CALLFLOW WHERE "
	
	If Request.Form("DESCRIPTION") = "" Then
		DESCRIPTION = "null"
	Else
		DESCRIPTION = "'" & Replace(Replace(Request.Form("DESCRIPTION"), "'" , ""), "%", "") & "'"
	End If
	If Request.Form("NAME") = "" Then
		NAME= "null"
	Else
		NAME =  "'" & Replace(Replace(Request.Form("NAME"), "'" , ""), "%", "") & "'" 
	End If
	If Request.Form("CALLFLOW_ID") = "" OR Not IsNumeric(Request.Form("CALLFLOW_ID")) Then
		CALLFLOW_ID = "null"
	Else
		CALLFLOW_ID = Replace(Replace(Request.Form("CALLFLOW_ID"), "'" , ""), "%", "")
		If CALLFLOW_ID = "" Then
			CALLFLOW_ID = "null"
		End IF
	End If
	If Request.Form("ahsid") = "" Then
		AHSID = "null"
	Else
		AHSID = Replace(Request.Form("ahsid") ,"''","'")
	End If
	If Request.Form("LOB_CD") = "" Then
		LOB_CD = "null"
	Else
		LOB_CD = "'" & Replace(Replace(Request.Form("LOB_CD"), "'" , "''"), "%", "") & "'"
	End If
	If Request.Form("FRAME_NAME") = "" Then
		FRAME_NAME = "null"
	Else
		FRAME_NAME = "'" & Replace(Replace(Request.Form("FRAME_NAME"), "'" , "''"), "%", "") & "'"
	End If

	QSQL = QSQL & "{call Designer_2.SrchCallFlow( "
	QSQL = QSQL & NAME &", "
	QSQL = QSQL & DESCRIPTION & ", " & CALLFLOW_ID & ", "
	QSQL = QSQL & LOB_CD & ", null, " & FRAME_NAME & ", "
	QSQL = QSQL & SEARCHTYPE & "," & MAXRECORDCOUNT & ",{resultset 10000, outCFName, outCFDescription, outCFID, outACFLOB})}"
	Set RS = Server.CreateObject("ADODB.RecordSet")
	ConnectionString = CONNECT_STRING
	RS.Open QSQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
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
	return "-0.1";
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
	return -0.1;
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
	if (RPID != "-0.1")
	{
		return "CFID=" + RPID + "&AHSID=" + document.all.tblResult.rows(index).getAttribute("AHSID")
	}
	else
	{
		return -0.1
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
-->
</SCRIPT>
</HEAD>
<BODY  bgcolor='<%= BODYBGCOLOR %>'  leftmargin=2 topmargin=2 rightmargin=0 bottommargin=2>
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblResult" name="tblResult" width="100%">
<THEAD>
<TR>
<TD CLASS=ResultHeader>CFID</TD>
<TD CLASS=ResultHeader>Name</TD>
<TD CLASS=ResultHeader>Description</TD>
<TD CLASS=ResultHeader>LOB</TD>
</TR>
</THEAD>
<%
If Request.Form <> "" Then
	Do While Not RS.EOF And RecCount <> MAXRECORDCOUNT
	RecCount = RecCount + 1
%>
<TR CLASS=RESULTROW OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);" RPID='<%= RS("outCFID") %>' AHSID=''>
<TD CLASS=ResultCell><NOBR><%= renderCell(RS("outCFID")) %></TD>
<TD CLASS=ResultCell><NOBR><%= renderCell(RS("outCFName")) %></TD>
<TD CLASS=ResultCell><NOBR><%= renderCell(RS("outCFDescription")) %></TD>
<TD CLASS=ResultCell><NOBR><%= renderCell(RS("outACFLOB")) %></TD>
</TR>
<%
RS.movenext
Loop
If RS.EOF AND RS.BOF Then 
%>
<TR>
<TD CLASS=LABEL COLSPAN=5>No Call Flows Found</TD>
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
