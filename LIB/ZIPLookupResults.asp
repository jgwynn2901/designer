<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\tablecommon.inc"-->
<%
'	main code
'option explicit

dim oConn, cSQL, ConnectionString
dim cCity, cZip, cState
dim cWhere, oRS


if vartype(Request.QueryString("HEADER")) = vbEmpty then
	cCity = "%" & Request.QueryString("CITY") & "%"
	cState = Request.QueryString("STATE")
	cZip = UCase(Request.QueryString("ZIP"))
	
	if cZip <> "" and IsNumeric(cZip) and Len(cZip) > 5 then
		cZip = Left(cZip, 5)
	End if
	
	cCity = Replace(cCity, "'" , "''")
	
	cWhere = ""
	If Request.QueryString("ZIP") <> "" Then
		cWhere = "ZIP like '" & cZIP & "%'"
	End If
	If Request.QueryString("CITY") <> "" Then
		If cWhere <> "" Then 
			cWhere = cWhere & " AND "
		End If
		cWhere = cWhere & "UPPER(CITY) LIKE '" & UCase(cCity) & "'"
	End If

	If Request.QueryString("STATE") <> "" Then
		If cWhere <> "" Then 
			cWhere = cWhere & " AND "
		End If
		cWhere = cWhere & "STATE = '" & UCase(cState) & "'"
	End If
	
	Set oConn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	oConn.Open ConnectionString

	cSQL = "Select * From LOCATION_CODE Where (" & cWhere & ")" 
	cSQL = cSQL & " Order By CITY" 
	Set oRS = Server.CreateObject("ADODB.Recordset")
	oRS.MaxRecords = MAXRECORDCOUNT
	oRS.Open cSQL, oConn, adOpenStatic, adLockReadOnly, adCmdText
end if	
%>
<HTML>
<HEAD>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
</HEAD>

<BODY  bgcolor='<%= BODYBGCOLOR %>'  leftmargin=10 topmargin=2 rightmargin=0 bottommargin=2>
<fieldset STYLE="BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#006699;height:'100%';width:'100%'">
<div align="LEFT" style="display:block;height:'100%';width:'100%';overflow:auto">
<table cellPadding="2" cellSpacing="0" rules="all" ID="tblResult" name="tblResult" width="100%">
<thead CLASS="ResultHeader">
<TR>
<td class="thd"><div id><nobr>City</div></td>
<td class="thd"><div id><nobr>State/Province</div></td>
<td class="thd"><div id><nobr>Zip</div></td>
<td class="thd"><div id><nobr>County</div></td>
<td class="thd"><div id><nobr>FIPS</div></td>
<td class="thd"><div id><nobr>Country</div></td>
</TR>
</THEAD>
<tbody ID="TableRows">
<%
if vartype(Request.QueryString("HEADER")) = vbEmpty then
	Do While Not oRS.EOF 
		%>
		<TR ID="FieldRow" CLASS="ResultRow" OnClick="Javascript:multiselect(this);">
		<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("CITY")) %></TD>
		<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("STATE")) %></TD>
		<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("ZIP")) %></TD>
		<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("COUNTY")) %></TD>
		<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("FIPS")) %></TD>
		<td NOWRAP CLASS="ResultCell" ><%= renderCell(oRS.Fields("COUNTRY")) %></TD>
		</TR>
		<%
		oRS.movenext
	Loop
	If oRS.EOF AND oRS.BOF Then 
	%>
		<TR>
		<TD CLASS=LABEL COLSPAN=10 ALIGN="MIDDLE">No Zip/Postal Code Found</TD>
		</TR>
	<%
	End If
	oRS.Close
	set oRS = nothing

	oConn.Close
	set oConn = nothing
end if
%>
</tbody>
</TABLE>
</DIV>
</FIELDSET>
<SCRIPT LANGUAGE="VBScript">
Sub window_onload
End Sub
</SCRIPT>

</BODY>
</HTML>
