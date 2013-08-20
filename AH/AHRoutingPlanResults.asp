<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"--> 
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Function Swap(InData)
	If InData <> "" Then
		Swap = "'" & InData & "'"
	Else
		Swap = "null"
	End If

End Function

Function Swap2(InData)
	If InData <> "" Then
		Swap2 = InData
	Else
		Swap2 = "null"
	End If
End Function

	dim RecCount
	RecCount = -1
If Request.Form <> "" Then
	RecCount = 0		' Needs to convert into 4 Single-Quotes for the Dynamic SQL in order to process (').	
	DESCRIPTION = Replace(Request.Form("DESCRIPTION"), "'" , "''''")
	DESTINATION_TYPE = Replace(Request.Form("DESTINATION_TYPE"), "'" , "''''")
	INPUT_SYSTEM_NAME = Replace(Request.Form("INPUT_SYSTEM_NAME"), "'" , "''''")
	STATE = Request.Form("STATE")
	LOB_CD =  Request.Form("LOB_CD")
	ROUTING_PLAN_ID = Replace(Request.Form("ROUTING_PLAN_ID"), "'" , "''''")
	AT_AHSID = Request.Form("AT_AHSID")
	
	Select Case Request.Form("SEARCHTYPE")
		Case "B"
			SEARCHTYPE = 1
		Case "C"
			SEARCHTYPE = 2
		Case "E"
			SEARCHTYPE = 3
		Case Else
			SEARCHTYPE = 1
	End Select
	SQLST = ""
	If Request.Form("SEARCHDIRECTION") = "UP" Then
		SQLST = SQLST & "{call Designer_2.SrchRoutingPlanUpTree("
	Else
		SQLST = SQLST & "{call Designer_2.SrchRoutingPlanDownTree("
	End If
	
	USER_ID = "null"
	If Not IsEmpty(Session("ACCOUNT_SECURITY")) Then USER_ID = Session("SecurityObj").m_UserID
	
	SQLST = SQLST & AT_AHSID & ", "
	SQLST = SQLST & Swap2(ROUTING_PLAN_ID) & ", "
	SQLST = SQLST & Swap(LOB_CD) & ", "
	SQLST = SQLST & Swap(STATE) & ", "
	SQLST = SQLST & Swap(DESCRIPTION) & ", "
	SQLST = SQLST & Swap(DESTINATION_TYPE) & ", "
	SQLST = SQLST & Swap(INPUT_SYSTEM_NAME) & ", "
	SQLST = SQLST & USER_ID & ", " ' User ID
	SQLST = SQLST & SEARCHTYPE & ", "
	SQLST = SQLST & "{resultset 20000, outRelatedAHSID, outRelatedAcctName, outRPID, outDesc, outLOB, outState})}"
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open CONNECT_STRING
	
	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	RS.Open SQLST, Conn, adOpenStatic, adLockReadOnly, adCmdText
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
			'call document.all.tblResult.focus() 'Doesnt like focus in the modal
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
	ClipboardAgent.AddProperty "ROUTING_PLAN_ID", ID
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
<TD CLASS=ResultHeader>RPID</TD>
<TD CLASS=ResultHeader>Description</TD>
<TD CLASS=ResultHeader>LOB</TD>
<TD CLASS=ResultHeader>State</TD>
</TR>
</THEAD>
<%
If Request.Form <> "" Then
	Do While Not RS.EOF
	RecCount = RecCount + 1
%>
<TR CLASS=RESULTROW OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);" RPID='<%= RS("outRPID") %>' AHSID='<%= RS("outRelatedAHSID") %>'>
<TD CLASS=ResultCell><%= renderCell(RS("outRelatedAHSID")) %></TD>
<TD CLASS=ResultCell><%= renderCell(RS("outRelatedAcctName")) %></TD>
<TD CLASS=ResultCell><%= renderCell(RS("outRPID")) %></TD>
<TD CLASS=ResultCell><%= renderCell(RS("outDesc")) %></TD>
<TD CLASS=ResultCell><%= renderCell(RS("outLOB")) %></TD>
<TD CLASS=ResultCell><%= renderCell(RS("outState")) %></TD>
</TR>
<%
RS.movenext
Loop
If RS.EOF AND RS.BOF Then 
%>
<TR>
<TD CLASS=LABEL COLSPAN=7 >No Routing Plans Found</TD>
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
