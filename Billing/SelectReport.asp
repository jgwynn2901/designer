<!--#include file="..\..\lib\common.inc"-->
<!--#include file="..\lib\genericSQL.asp"-->
<%
dim cAHS, cStartDate, oRS, lNewQuery, nResult, cSQL, cCustName, cCustCode

cAHS = Request.QueryString("AHS")
cStartDate = Request.QueryString("DATE")
cCustName = Request.QueryString("CUSTNAME")
cCustCode = Request.QueryString("CUSTCODE")

cSQL = "Select * from BILLING_HISTORY Where MMM_YYYY='" & UCase(cStartDate) & "' and AHS_ID='" & cAHS & "'"
Set oRS = Conn.Execute(cSQL)
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!--#include file="..\..\lib\tablecommon.inc"-->
<link rel="stylesheet" type="text/css" href="..\..\FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub BtnSelect_onclick
dim nIndex, cFileID, cFilePath

nIndex = getselectedindex(document.all.tblFields)
if nIndex < 0 then
	msgbox "Please select a Report.",,"FNSDesigner"
else
	cFileID = document.all.tblFields.Rows(nIndex).GetAttribute("FileID")
	cFilePath = document.all.tblFields.Rows(nIndex).GetAttribute("FilePath")
	self.location.href = "getReport.asp?FILEID=" & cFileID & "&FILEPATH=" & cFilePath
end if	
End Sub

-->
</SCRIPT>
</HEAD>
<BODY topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">

<div align="LEFT">
<TABLE width="32%">
	<tr>
	<td CLASS="GrpLabel" WIDTH="50" HEIGHT="12"><font face="Verdana, Helvetica, Arial"><nobr>&nbsp;» Existing reports:</font></td>	
	</tr>
</table></div>
<div align="left">
<table border="1" rules="all" ID="tblFields" name="tblFields" width="64%" >
	<tr>&nbsp;</tr>
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><div id><nobr>Created on</div></td>
			<td class="thd"><div id><nobr>Created by</div></td>
		</tr>
	</thead>
	<tbody ID="TableRows">
<% 
Do While not oRS.EOF 
%>
	<tr ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);" FileID='<%=oRS.Fields("FILENAME").value%>' FilePath='<%=oRS.Fields("FILE_PATH").value%>'>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS.Fields("CREATED_ON").value)%></td>
	<td NOWRAP CLASS="ResultCell"><%=renderCell(oRS.Fields("CREATED_BY").value)%></td>
	</tr>
<%
	oRS.MoveNext
Loop
oRS.CLose
set oRS = nothing
%>
</TABLE>
<TABLE width="64%">
<tr>&nbsp;</tr>
<TR>
<TD CLASS=LABEL align="right"><BUTTON NAME=BtnSelect CLASS=STDBUTTON ACCESSKEY="S"><U>S</U>elect</BUTTON></TD>
</TR>
</TABLE>
</div>
</BODY>
</HTML>
