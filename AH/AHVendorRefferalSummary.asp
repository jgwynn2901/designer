<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\security.inc"-->
<!--#include file="..\lib\commonError.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ClipboardCLSID.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->  
<!--#include file="..\lib\AHSTree.inc"--> 
<% Response.Expires=0 
   Response.Buffer = True
   On Error Resume Next
%>
<%
Function NextPkey( TableName, ColName )
	NextSQL = "{call Designer.GetValidSeq('" & TableName & "', '" & ColName &"', {resultset 1, outResult})}"
	Set NextRS = Conn.Execute(NextSQL)
	NextPkey = NextRS("NextID") 
End Function

	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString

If Request.QueryString("INSERT") <> "" Then

	SQLINS = "INSERT INTO VENDOR_REFERRAL_TYPE (VENDOR_REFERRAL_TYPE_ID,ACCNT_HRCY_STEP_ID) VALUES (" 
	SQLINS = SQLINS & NextPkey("VENDOR_REFERRAL_TYPE","VENDOR_REFERRAL_TYPE_ID") & "," & Request.QueryString("INSERT") & ")"
	Set RS = Conn.Execute(SQLINS)
	Set RS = Nothing
	Conn.Close
	Set Conn = Nothing
	Response.Redirect "AHVendorRefferalSummary.asp?AHSID=" & Request.QueryString("AHSID") 
End If
If Request.QueryString("DELETED") <> "" Then
  
      SQLDEL=  "DELETE FROM VENDOR_REFERRAL_RULE WHERE VENDOR_REFERRAL_TYPE_ID = " & Request.QueryString("DELETED")
          Set RS1 = Conn.Execute(SQLDEL)
      SQLDEL1=  "DELETE FROM VENDOR_REFERRAL_TYPE WHERE VENDOR_REFERRAL_TYPE_ID  = " & Request.QueryString("DELETED")
	  Set RS = Conn.Execute(SQLDEL1)
     'SQLDEL = "{call Designer3.SP_DELETE_VENDORREFERRALTYPE(" & Request.QueryString("DELETED") & ")}"
	  'Set RS = Conn.Execute(SQLDEL)
	strError = CheckADOErrors(Conn,"DELETE" )

	Set RS = Nothing
	Conn.Close
	Set Conn = Nothing
	
	If strError = "" Then 
		Response.Redirect "AHVendorRefferalSummary.asp?AHSID=" & Request.QueryString("AHSID")
	End If

End If


If Request.QueryString("CLEARFILTER") <> "" Then
	RemoveFilter "AHSID=" & Request.QueryString("AHSID"),"DESIGNER_BAFILTER"
	Response.Redirect "AHVendorRefferalSummary.asp?AHSID=" & Request.QueryString("AHSID")
End If

	Set RS = Server.CreateObject("ADODB.RecordSet")
	RS.MaxRecords = MAXRECORDCOUNT
	ConnectionString = CONNECT_STRING
	
	
			
	SQL = "SELECT * FROM VENDOR_REFERRAL_TYPE, RULES  WHERE VENDOR_REFERRAL_TYPE.RULE_ID = RULES.RULE_ID(+) AND " &_
		 "VENDOR_REFERRAL_TYPE.ACCNT_HRCY_STEP_ID=" & Request.QueryString("AHSID")
	
	strInclude = GetSpecificFilter("AHSID=" & Request.QueryString("AHSID"), "DESIGNER_BAFILTER", "MUSTINCLUDE")	
	
	If Request.QueryString("MultiSelected") <> "" Then
		If strInclude <> "" Then
			strInclude = strInclude & ", " &  Request.QueryString("multiselected")
		Else
			strInclude  = Request.QueryString("multiselected")
		End If
	SetFilterByName "AHSID=" & Request.QueryString("AHSID"), "DESIGNER_BAFILTER", "MUSTINCLUDE", strInclude
	End If
	
	if strInclude <> "" then SQL = SQL & " AND VENDOR_REFERRAL_TYPE.VENDOR_REFERRAL_TYPE_ID  IN (" & strInclude & ") "

	RS.Open SQL, ConnectionString, adOpenStatic, adLockReadOnly, adCmdText
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
function dblhighlight( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("BATID");
}
function getname( objRow )
{
	currentRowIndex = objRow.rowIndex;
	currentobjRow = objRow;
	// call the vbscript event handler function
	return objRow.getAttribute("BRNAME");
}
function FilterSpan_OnClick()
{
	lret = confirm("Are you sure you want to clear the filter?");
	if (lret == true)
		self.location.href = "AHVendorRefferalSummary.asp?<%= Request.Querystring %>" + "&CLEARFILTER=TRUE"
}

function CBranchAssignTypeSearchObj()
{
	this.BATID = "";
	this.Selected = "";
}
var BranchAssignTypeSearchObj = new CBranchAssignTypeSearchObj();
-->
</SCRIPT>
<!-- #include file="..\lib\BRBtnControl.inc" -->
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Function EditClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then
		parent.frames.window.location = "../VendorReferal/VendorReferalMaintenance.asp?CONTEXTTYPE=DRILLIN&CONTAINERTYPE=FRAMEWORK&DETAILONLY=TRUE&BATID=" & dblhighlight(Document.all.tblResult.rows(i)) & "&AHSID=<%= Request.QueryString("AHSID") %>"
	end if
End Function

Function NewClick()
    parent.frames.window.location = "../VendorReferal/VendorReferalMaintenance.asp?CONTEXTTYPE=DRILLIN&CONTAINERTYPE=FRAMEWORK&BATID=NEW&DETAILONLY=TRUE&AHSID=<%= Request.QueryString("AHSID") %>"
End Function

Function RemoveClick()
	i = getselectedindex( Document.all.tblResult )
	if 0 < i then

		BATID = dblhighlight(Document.all.tblResult.rows(i))
		lret = MsgBox ("Are you sure you want to delete Vendor referral Type ID:" & BATID & " for Account ID:<%=Request.QueryString("AHSID")%>?", 1, "FNSDesigner")
		If lret = 1 Then
	       self.location.href = "AHVendorRefferalSummary.asp?DELETED=" & BATID & "&AHSID=<%= Request.QueryString("AHSID") %>"
		End If
		
	end if
End Function

Function SearchClick()
	lret = ""
	strURL = ""
	BranchAssignTypeSearchObj.BATID = ""
	strURL = "../VendorReferal/VendorReferalMaintenance.asp?CONTAINERTYPE=MODAL&MODE=RO&SearchAHSID=<%= Request.QueryString("AHSID") %>"
	lret = window.showModalDialog(strURL, BranchAssignTypeSearchObj ,"center")
	if BranchAssignTypeSearchObj.BATID <> "" Then
		multi = Replace(BranchAssignTypeSearchObj.BATID,"||",",")
		self.location.href = "AHVendorRefferalSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>&MultiSelected=" & multi
	End If
End Function


Function RefreshClick()
	self.location.href = "AHVendorRefferalSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
End Function

Sub window_onload
<% If RS.RecordCount = MAXRECORDCOUNT Then %>
	StatusSpan.innerHTML = "<%= MSG_MAXRECORDS %>"
<% Else %>
	StatusSpan.innerHTML = "Record Count is <%= RS.RecordCount %>"
<% End If %>	
	ClipboardAgent.GetPropertiesFromClipboard

<% If strInclude <> "" Then %>
	FilterSpan.innerHTML = "<IMG SRC='..\images\filter2.gif'></IMG>"
<%	Else %>	
	FilterSpan.innerHTML = ""
<%	End If%>


<%If strError <> "" Then %>
	MsgBox ("<%=strError%>")
<% End If %>


End Sub
-->
</SCRIPT>
</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<FIELDSET STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699">
<%
PARAMS = ""
PARAMS = PARAMS & "&HIDECOPY=TRUE&HIDEPASTE=TRUE&HIDEREFRESH=TRUE&HIDESEARCH=TRUE"
%>
<OBJECT data="../Scriptlets/ObjButtons.asp?HIDEATTACH=TRUE<%=PARAMS%>" STYLE="WIDTH:100%;HEIGHT:23;LEFT:0" id="BRBtnControl" type=text/x-scriptlet></OBJECT>
<SPAN  STYLE="CURSOR:HAND" TITLE="Clear Filter" LANGUAGE="JScript" ONCLICK="return FilterSpan_OnClick()" align=right ID=FilterSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<SPAN ID=StatusSpan STYLE="COLOR:#006699" CLASS=LABEL></SPAN>
<DIV align="LEFT" id="Branch_RESULTS" style="display:block;height:145;width:'100%';overflow:scroll">
<table cellPadding=2 cellSpacing=0  rules=all ID="tblResult" name="tblResult" width=100%>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class=thd><div><NOBR>V.R.T.  ID</div></td>
			<td class=thd><div><NOBR>Description</div></td>
			<td class=thd><div><NOBR>Rule Text</div></td>
			<td class=thd><div><NOBR>Rule ID</div></td>			
		</tr>
	</thead>
	<tbody ID="TableRows">
<% Do While Not RS.EOF %>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" OnDblClick="Javascript:dblclick(this);" BRNAME="<%= Mid(RS("DESCRIPTION"),1,25) %>" BATID='<%= RS("VENDOR_REFERRAL_TYPE_ID") %>'>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("VENDOR_REFERRAL_TYPE_ID")) %></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("DESCRIPTION")) %></td>
			<td  TITLE="<%=ReplaceQuotesInText(renderCell(RS("RULE_TEXT")))%>" NOWRAP CLASS="ResultCell"><%=TruncateText(renderCell(RS("RULE_TEXT")),25)%></td>
			<td NOWRAP CLASS=ResultCell><%= renderCell(RS("RULE_ID")) %></td>			
	</tr>
<% 
RS.MoveNext
Loop
RS.Close
Set RS = Nothing
Conn.Close
Set Conn = Nothing
%>
</tbody>
</table>
</DIV>
</FIELDSET>
</SCRIPT>

<OBJECT ID="ClipboardAgent" VIEWASTEXT
<%GetClipboardCLSID()%>
width=1 height=1>
<PARAM NAME="MaxPropertiesStringLength" VALUE="1000">
<PARAM NAME="MaxPropertyNameLength" VALUE="50">
<PARAM NAME="MaxPropertyValueLength" VALUE="200">
<PARAM NAME="NameValueDelimiter" VALUE="#">
<PARAM NAME="PropertyItemDelimiter" VALUE="|">
<PARAM NAME="PrivateClipboardFormatName" VALUE="CF_FNSDESIGNER">
</OBJECT>
</BODY>
</HTML>
