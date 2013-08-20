<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"--> 
<!--#include file="..\lib\AHSTree.inc"--> 
<%
dim oConn, oRS, cSQL, ConnectionString
dim oRS1, lError, cSD, aDest
dim lFirst, nRecs

Response.Expires=0 
Response.Buffer = true
Response.AddHeader  "Pragma", "no-cache"

Set oConn = Server.CreateObject("ADODB.Connection")
ConnectionString = CONNECT_STRING
oConn.Open ConnectionString
nRecs = 0
If HasModifyPrivilege("FNSD_SPECIFIC_DESTINATION", SECURITYPRIV) <> True Then 
	cMODE = "RO"
else
	cMODE = "RW"
end if	
cSD = trim(Request.QueryString("SD"))
aDest = SplitNew(cSD)
cSQL = "SELECT * FROM SPECIFIC_DESTINATION WHERE SPECIFIC_DESTINATION_ID IN ("
lFirst = true
for x=lBound(aDest) to uBound(aDest)
	if lFirst then
		cSQL = cSQL & aDest(x)
		lFirst = false
	else
		cSQL = cSQL & "," & aDest(x)
	end if
next
cSQL = cSQL & ")"
if not lFirst then
	set oRS = oConn.Execute(cSQL)
end if	

function SplitNew(cInput)
dim x, nLen
dim cChar
dim aValues
dim cCell

' create an empty array
aValues = Array()
nLen = Len(cInput)
cCell = ""
x = 1
do while x <= nLen
	cChar = mid(cInput, x, 1)
	if cChar <> " " then
		cCell = cCell & cChar
		x = x + 1
	else
		redim preserve aValues(ubound(aValues) + 1)
		aValues(ubound(aValues)) = cCell
		cCell = ""
		x = x + 1
		do while mid(cInput, x, 1) = " " and x <= nLen
			x = x + 1
		loop
	end if
loop
if cCell <> "" then
	redim preserve aValues(ubound(aValues) + 1)
	aValues(ubound(aValues)) = cCell
end if
SplitNew = aValues
end function

%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<!--#include file="..\lib\tablecommon.inc"-->
<SCRIPT LANGUAGE="Javascript">
<!--

function dblclick( objRow )
{
EditClick()
}

function specDestObj()
{
	this.selectedID = "";
}

function newSpecDest()
{
	this.SDID = "";

}

var SearchObj = new specDestObj();
var oSD = new newSpecDest();

-->
</SCRIPT>

<SCRIPT LANGUAGE="JavaScript" FOR="CFBtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
	case "EDITBUTTONCLICK":
       	EditClick()
		break;

	case "NEWBUTTONCLICK":
		NewClick()
		break;

	case "REMOVEBUTTONCLICK":
		RemoveClick()
		break;
	
	case "REFRESHBUTTONCLICK":
		RefreshClick();
		break;
	case "ATTACHBUTTONCLICK":
		AttachClick();
		break;
	default:
		break;
}
   
</SCRIPT>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
sub EditClick()
dim lRet, cURL, i, cSD, cName, cAHSID

i = getselectedindex( Document.all.tblResult )
if 0 < i then
	cSD = Document.all.tblResult.rows(i).getAttribute("SDID")
	cName = "<%=server.URLEncode(Request.Querystring("NAME"))%>"
	cAHSID = "<%=Request.Querystring("AHSID")%>"
	cURL = "SpecificDestination.asp?MODE=<%=cMODE%>&SDID=" & cSD & "&NAME=" & cName & "&AHSID=" & cAHSID
	lRet = window.showModalDialog(cURL, oSD,"dialogWidth:625px;dialogHeight:650px;center")
	self.location.href = "AHSpecDestBott.asp?" & "<%=Request.Querystring%>"
else
	msgbox "Please select a Specific Destination to edit.",,"FNSDesigner"
end if
End sub

sub NewClick()
dim lRet, cURL, cSD
dim cAHSID, cName, cMode

cAHSID = "<%=Request.Querystring("AHSID")%>"
cName = "<%=server.URLEncode(Request.Querystring("NAME"))%>"
cURL = "SpecificDestination.asp?MODE=" & cMode & "&SDID=NEW&NAME=" & cName & "&AHSID=" & cAHSID & "&CLIENT_NODE=<%=Request.Querystring("CLIENT_NODE")%>"
oSD.SDID = ""
lRet = window.showModalDialog(cURL, oSD,"dialogWidth:625px;dialogHeight:650px;center")
cSD = "<%=trim(Request.QueryString("SD"))%>"
if oSD.SDID <> "" then
	cSD = cSD & " " & trim(oSD.SDID)
end if
self.location.href = "AHSpecDestBott.asp?SD=" & cSD & "&AHSID=" & cAHSID & "&NAME=" & cName & "&CLIENT_NODE=<%=Request.Querystring("CLIENT_NODE")%>&MODE=<%=cMODE%>"
End sub

sub RemoveClick()
dim i, nRet, cSDID, cSDs
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		cSDID = Document.all.tblResult.rows(i).getAttribute("SDID")
		cSDs = "<%=Request.QueryString("SD")%>"
		cSDs = removeSD(cSDs, cSDID)
		self.location.href = "AHSpecDestBott.asp?SD=" & cSDs & "&AHSID=<%=Request.Querystring("AHSID")%>&NAME=<%=server.URLEncode(Request.Querystring("NAME"))%>&CLIENT_NODE=<%=Request.QueryString("CLIENT_NODE")%>&MODE=<%=cMODE%>"
	else
		msgbox "Please select a Specific Destination to delete.",,"FNSDesigner"
	end if
End sub

sub RefreshClick()
	self.location.href = "AHSpecDestBott.asp?SD=<%=Request.QueryString("SD")%>&AHSID=<%=Request.Querystring("AHSID")%>&NAME=<%=server.URLEncode(Request.Querystring("NAME"))%>&CLIENT_NODE=<%=Request.QueryString("CLIENT_NODE")%>&MODE=<%=cMODE%>"
End sub


sub AttachClick()

dim lRet, cURL

SearchObj.selectedID = ""
cURL = "SpecDesStep.asp?CLIENT_NODE=<%=Request.QueryString("CLIENT_NODE")%>"
lRet = window.showModalDialog(cURL  ,SearchObj ,"dialogWidth:625px;dialogHeight:550px;center")
if SearchObj.selectedID <> "" Then
	self.location.href = "AHSpecDestBott.asp?SD=<%=trim(Request.QueryString("SD"))%> " & SearchObj.selectedID & "&AHSID=<%=Request.Querystring("AHSID")%>&NAME=<%=server.URLEncode(Request.Querystring("NAME"))%>&CLIENT_NODE=<%=Request.QueryString("CLIENT_NODE")%>&MODE=<%=cMODE%>"
End If
End sub

function removeSD(cString, cIDtoRemove)
dim aFav, x
dim lFirst

removeSD = ""
lFirst = true
aFav = split(cString, " ")
for x=0 to ubound(aFav)
	if aFav(x) <> cIDtoRemove then
		if lFirst then
			removeSD = aFav(x)
			lFirst = false		
		else
			removeSD = removeSD & " " & aFav(x)
		end if
	end if
next	
End function

-->
</SCRIPT>
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<FIELDSET STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699">
<%
	dim cParams
	
	cParams = "&HIDESEARCH=TRUE&HIDEPASTE=TRUE&HIDECOPY=TRUE"
	if cMODE = "RO" then
		cParams = cParams & "&HIDENEW=TRUE&HIDEATTACH=TRUE&HIDEREMOVE=TRUE"
	end if
%>
<OBJECT id=CFBtnControl style="LEFT: 0px; WIDTH: 100%; HEIGHT: 23px" type=text/x-scriptlet data="../Scriptlets/ObjButtons.asp?REMOVECAPTION=Remove<%=cParams%>">
	</OBJECT>
<SPAN  STYLE="CURSOR:HAND" TITLE="Clear Filter" LANGUAGE="JScript" ONCLICK="return FilterSpan_OnClick()" align=right ID=FilterSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<SPAN ID=StatusSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<DIV align="LEFT" id="Account_RESULTS" style="display:block;height:230;width:'100%';overflow:scroll">
<table cellPadding=2 cellSpacing=0 rules=all ID="tblResult" name="tblResult" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div><NOBR>SD ID</div></td>
			<td class=thd><div><NOBR>Last Name</div></td>
			<td class=thd><div><NOBR>First Name</div></td>
			<td class=thd><div><NOBR>Title</div></td>
			<td class=thd><div><NOBR>Address</div></td>
			<td class=thd><div><NOBR>City</div></td>
			<td class=thd><div><NOBR>State</div></td>
			<td class=thd><div><NOBR>Zip Code</div></td>
			<td class=thd><div><NOBR>Phone</div></td>			
		</tr>
	</thead>
	<tbody ID="TableRows">
<% 
if not lFirst then
	Do While Not oRS.EOF
		nRecs = nRecs + 1 %>
		<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblclick(this);" SDID="<%=oRS.fields("SPECIFIC_DESTINATION_ID")%>">
		<td NOWRAP CLASS="ResultCell"><%= renderCell(oRS.fields("SPECIFIC_DESTINATION_ID")) %></td>		
		<td NOWRAP CLASS="ResultCell"><%= renderCell(oRS.fields("NAME_LAST")) %></td>
		<td NOWRAP CLASS="ResultCell"><%= renderCell(oRS.fields("NAME_FIRST")) %></td>
		<td NOWRAP CLASS="ResultCell"><%= renderCell(oRS.fields("TITLE")) %></td>
		<td NOWRAP CLASS="ResultCell"><%= renderCell(oRS.fields("ADDRESS1")) %></td>
		<td NOWRAP CLASS="ResultCell"><%= renderCell(oRS.fields("CITY")) %></td>
		<td NOWRAP CLASS="ResultCell"><%= renderCell(oRS.fields("STATE")) %></td>
		<td NOWRAP CLASS="ResultCell"><%= renderCell(oRS.fields("ZIP")) %></td>
		<td NOWRAP CLASS="ResultCell"><%= renderCell(oRS.fields("PHONE")) %></td>
		</TR>
		<%
		oRS.MoveNext
	Loop
	oRS.Close
	set oRS = nothing
	oConn.Close
	set oConn = nothing
end if	
%>
</tbody>
</table>
</DIV>
</FIELDSET>
<SCRIPT LANGUAGE="vbscript">
Sub window_onload
	<% If nRecs >= MAXRECORDCOUNT Then %>
			StatusSpan.innerHTML = "<%= MSG_MAXRECORDS %>"
	<% Else %>
			StatusSpan.innerHTML = "Record Count is <%=nRecs%>"
	<% End If %>	

End Sub

</SCRIPT>
</BODY>
</HTML>

