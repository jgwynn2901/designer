<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Function RenderCell( strValue )
	if IsNull(strValue) or Len(Trim(strValue)) <= 0 then
		RenderCell = "&nbsp;"
	else
		RenderCell = strValue
	end if
End Function

	dim RecCount
	RecCount = -1
If Request.Form("SEARCHTYPE") <> "" Then
	RecCount = 0
		Select Case Request.Form("SEARCHTYPE")
			Case "B"
				DESCRIPTION = Request.Form("DESCRIPTION") & "%"
				DESTINATION_TYPE = Request.Form("DESTINATION_TYPE") & "%"
				INPUT_SYSTEM_NAME = Request.Form("INPUT_SYSTEM_NAME") & "%"
				STATE = Request.Form("STATE") & "%"
				LOB_CD = Request.Form("LOB_CD") & "%"
				ROUTING_PLAN_ID = Request.Form("ROUTING_PLAN_ID") & "%"
			Case "C"
				DESCRIPTION = "%" & Request.Form("DESCRIPTION") & "%"
				DESTINATION_TYPE = "%" & Request.Form("DESTINATION_TYPE") & "%"
				INPUT_SYSTEM_NAME = "%" & Request.Form("INPUT_SYSTEM_NAME") & "%"
				STATE = "%" & Request.Form("STATE") & "%"
				LOB_CD = "%" & Request.Form("LOB_CD") & "%"
				ROUTING_PLAN_ID = "%" & Request.Form("ROUTING_PLAN_ID") & "%"
			Case "E"
				DESCRIPTION = Request.Form("DESCRIPTION")
				DESTINATION_TYPE = Request.Form("DESTINATION_TYPE")
				INPUT_SYSTEM_NAME = Request.Form("INPUT_SYSTEM_NAME")
				STATE = Request.Form("STATE")
				LOB_CD = Request.Form("LOB_CD")
				ROUTING_PLAN_ID = Request.Form("ROUTING_PLAN_ID")
		End Select
	
	DESCRIPTION = Replace(DESCRIPTION, "'" , "''")
	DESTINATION_TYPE = Replace(DESTINATION_TYPE, "'" , "''")
	INPUT_SYSTEM_NAME = Replace(INPUT_SYSTEM_NAME, "'" , "''")
	STATE = Replace(STATE, "'" , "''")
	LOB_CD = Replace(LOB_CD, "'" , "''")
	ROUTING_PLAN_ID = Replace(ROUTING_PLAN_ID, "'" , "''")
	
	'Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	'Conn.Open ConnectionString
	SQLST = ""
	SQLWHERE = ""
	SQLST = SQLST & "SELECT * FROM ROUTING_PLAN "
	If Request.Form("DESCRIPTION") <> "" Then
		SQLWHERE = SQLWHERE & "UPPER(DESCRIPTION) LIKE '" & UCASE(DESCRIPTION) & "' "
	End If
	If Request.Form("DESTINATION_TYPE") <> "" Then
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " AND " 
		End If
		SQLWHERE = SQLWHERE & "UPPER(DESTINATION_TYPE) LIKE '" & UCASE(DESTINATION_TYPE) & "' "
	End If
	If Request.Form("INPUT_SYSTEM_NAME") <> "" Then
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " AND " 
		End If
		SQLWHERE = SQLWHERE & "UPPER(INPUT_SYSTEM_NAME) LIKE '" & UCASE(INPUT_SYSTEM_NAME) & "' "
	End If
	If Request.Form("ROUTING_PLAN_ID") <> "" Then
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " AND " 
		End If
		SQLWHERE = SQLWHERE & "UPPER(ROUTING_PLAN_ID) LIKE '" & UCASE(ROUTING_PLAN_ID) & "' "
	End If
	If Request.Form("STATE") <> "" Then
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " AND " 
		End If
		SQLWHERE = SQLWHERE & "UPPER(STATE) LIKE '" & UCASE(STATE) & "' "
	End If
	If Request.Form("LOB_CD") <> "" Then
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " AND " 
		End If
		SQLWHERE = SQLWHERE & "UPPER(LOB_CD) LIKE '" & UCASE(LOB_CD) & "' "
	End If
	If Request.Form("ahsid") <> "" Then
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " AND " 
		End If
		SQLWHERE = SQLWHERE & "accnt_hrcy_step_id=" & Request.Form("ahsid") & " "
	End If
	If Request.Form("s_EnabledFlag") <> "" Then
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " AND " 
		End If
		SQLWHERE = SQLWHERE & "ENABLED_FLG='" & UCASE(Request.Form("s_EnabledFlag")) & "' " 
	End If
	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
	If SQLWHERE <> "" Then
		SQLST = SQLST & " WHERE " & SQLWHERE
	End If
	RS.Open SQLST, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
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

function GetMultipleIndexDesc( objTable )
{
var desc
desc = "";
	for(i=1;i<objTable.rows.length;i++)
	{
		if( objTable.rows[i].className == 'ResultSelectRow' )
		{
			if (desc != "") 
			{
			desc = desc + ",";
			}
			desc = desc + dblHighLightGetDesc(objTable.rows[i]) 
		}		
	
	}
	return desc;
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

function dblHighLightGetDesc( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	return objRow.getAttribute("RPDesc");
}


function SelectRow() {
	RPID = getselectedindex(document.all.tblResult )
	index = getindex(document.all.tblResult)
	if (RPID != "-1")
	{
		return "RPID=" + RPID + "&AHSID=" + document.all.tblResult.rows(index).getAttribute("AHSID")
	}
	else
	{
		return -1
	}
}

function getindex( objTable )
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


-->
</script>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub window_onload
<% If Request.Form <> "" Then %>
		if 0 < Document.all.tblResult.rows.length then
			call multiselect( Document.all.tblResult.rows(1))
			'call document.all.tblResult.focus() 'Doesnt like focus in the modal
		end if
<%End if %>
End Sub

Function GetRPID
	GetRPID = getmultipleindex(document.all.tblResult)
End Function

Function GetRPDesc
	GetRPDesc = getmultipleindexdesc(document.all.tblResult)
End Function

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
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblResult" name="tblResult" width="100%" RPID='X'>
<THEAD>
<TR>
<TD CLASS=ResultHeader>RPID</TD>
<TD CLASS=ResultHeader>Description</TD>
<TD CLASS=ResultHeader>LOB</TD>
<TD CLASS=ResultHeader>State</TD>
<TD CLASS=ResultHeader>Enabled</TD>
</TR>
</THEAD>
<%
If Request.Form <> "" Then
	Do While Not RS.EOF
	RecCount = RecCount + 1
%>
<TR CLASS=RESULTROW OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);" RPID='<%= RS("ROUTING_PLAN_ID") %>' RPDesc='<%= Mid(RS("DESCRIPTION"),1, 25) & "..." %>' AHSID='<%= RS("ACCNT_HRCY_STEP_ID") %>'>
<TD CLASS=ResultCell><%= RenderCell(RS("ROUTING_PLAN_ID")) %></TD>
<TD CLASS=ResultCell id="DESCRIPTION"><%= RenderCell(RS("DESCRIPTION")) %></TD>
<TD CLASS=ResultCell><%= RenderCell(RS("LOB_CD")) %></TD>
<TD CLASS=ResultCell><%= RenderCell(RS("STATE")) %></TD>
<TD CLASS=ResultCell><%= RenderCell(RS("ENABLED_FLG")) %></TD>
</TR>
<%
RS.movenext
Loop
If RS.EOF AND RS.BOF Then 
%>
<TR>
<TD CLASS=LABEL COLSPAN=5>No Routing Plans Found</TD>
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
