<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"--> 
<%
Response.Expires = 0
Response.Buffer = true
on error resume next
	Function RenderCell( strValue )
		if IsNull(strValue) or Len(Trim(strValue)) <= 0 then
			RenderCell = "&nbsp;"
		else	
			RenderCell = strValue
		end if
	End Function

	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	SQLST = ""
	If Request.QueryString("SQL") = "FROM" Then
		SQLST = SQLST & "SELECT DISTINCT TABLE_NAME FROM USER_TABLES"
	End If
	Set RS = Conn.Execute(SQLST)
%>
<HTML>
<HEAD>
<TITLE>SQL Help</TITLE>
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
	parent.frames("RIGHT").location.href = "SQLHELPCOLUMNS.asp?TABLENAME=" + object.getAttribute("TABLE")
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
	return objRow.getAttribute("TABLE");
	//NavCall(objRow.getAttribute("LOB"), objRow.getAttribute("CALLID"), objRow.getAttribute("AHSID"))
	//alert(objRow.getAttribute("TABLE"))
	//window.dialogArguments.tablename = objRow.getAttribute("TABLE")
	//window.close()
			//lret = ClipboardAgent.ClearAllProperties();
			//lret = ClipboardAgent.PropertiesString = objRow.getAttribute("TABLE");
			//lret = ClipboardAgent.SetPropertiesToClipboard();
	parent.frames("RIGHT").location.href = "SQLHELPCOLUMNS.asp?TABLENAME=" + objRow.getAttribute("TABLE")
}

function GetTableName(  )
{
lret = getselectedindex( tblResult )
	if (lret != "")
	{
		return tblResult.rows(lret).getAttribute("TABLE")
	}
	else
	{
		return -1;
	}
}

-->
</script>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

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

Sub window_onload
	if 0 < Document.all.tblResult.rows.length then
		call multiselect( Document.all.tblResult.rows(1))
		call Document.all.tblResult.focus()
	end if
End Sub

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
<PARAM NAME="PrivateClipboardFormatName" VALUE="CF_TEXT">
</OBJECT>
</HEAD>
<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<TABLE ID="tblResult" BORDER=1 BGCOLOR="WINDOW" CLASS=Label width = 100%>
<THEAD>
<TR CLASS=RESULTHEADER>
<TD CLASS=LABEL><FONT COLOR="WHITE">Table Name</FONT></TD>
</TR>
<% Do While Not RS.EOF %>
<TR CLASS=RESULTROW OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblhighlight(this);" TABLE='<%= RS("TABLE_NAME") %>'  title="Click copy button to place name in clipboard.">
<TD CLASS=LABEL><%= RS("TABLE_NAME") %></TD>
</TR>
<%
RS.MoveNext
Loop
RS.Close
%>
</TABLE>
</BODY>
</HTML>
