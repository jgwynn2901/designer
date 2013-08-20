<!--#include file="..\lib\common.inc"-->
<%
Function RenderCell( strValue )
	if IsNull(strValue) or Len(Trim(strValue)) <= 0 then
		RenderCell = "&nbsp;"
	else
		RenderCell = strValue
	end if
End Function

%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE></TITLE>
<STYLE>
BODY { 
		background:#d6cfbd;
		Font-Family:Verdana;
		Font-Size:10 
		}
</STYLE>
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
			return i;
			
		}		
	}
	return -1;
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
	//NavCall(objRow.getAttribute("LOB"), objRow.getAttribute("CALLID"), objRow.getAttribute("AHSID"))
	alert ("here")
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
-->
</SCRIPT>
</HEAD>
<BODY>
<TABLE ID="tblResult" BORDER=1 BGCOLOR="WINDOW" CLASS=Label width = 100%>
<THEAD>
<TR>
<TD CLASS=ResultHeader>Policy Number</TD>
<TD CLASS=ResultHeader>Policy Description</TD>
<TD CLASS=ResultHeader>Effective Date</TD>
<TD CLASS=ResultHeader>LOB_CD</TD>
</TR>
</THEAD>
<%
If Request.Form <> "" Then
	'Set Conn = Server.CreateObject("ADODB.Connection")
	'ConnectionString = CONNECT_STRING
	'Conn.Open ConnectionString
	'SQLST = ""
	'SQLWhere = ""
	SQLST = SQLST & "SELECT * FROM POLICY WHERE "
	For Each x In Request.Form 
		If Request.Form(x) <> "" Then
				If SQLWhere <> "" Then
					SQLWhere = SQLWhere & " AND "
				End If
			SQLWhere = SQLWhere & "UPPER(" & x & ") LIKE '" & UCASE(Request.Form(x)) & "%'"
		End If
	Next	
	Response.Write(SQLST & SQLWhere)
	'Set RS = Conn.Execute(SQLST & SQLWhere)
	'Do While Not RS.EOF
%>
<TR CLASS=RESULTROW OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);">
<TD></TD>
<%
Else

End If 
%>
</BODY>
</HTML>
