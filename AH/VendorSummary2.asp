<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\commonError.inc"-->
<% 
Response.Expires=0 

dim cAccVendorID, oConn, oRS, cSQL
dim nVID, nNID, nSEQ, nAHSID, cLOB, nST, nCM
dim cError, lUpdateOK
	
nVID = Request.QueryString("VID")
nNID = Request.QueryString("NID")
nSEQ = Request.QueryString("SEQ")	
nAHSID = Request.QueryString("AHSID")
cLOB = Request.QueryString("LOB")
nST = Request.QueryString("ST")
nCM = Request.QueryString("CM")
nUpdateStatus = 0
	
if not isEmpty(Request.QueryString("NEW")) then
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	nNewID = CLng(NextPkey("ACCOUNT_VENDOR","ACCOUNT_VENDOR_ID"))
	If nNewID > 0 Then
		cSQL = "Insert into ACCOUNT_VENDOR values (" & nNewID & "," & nAHSID & "," & nNID & "," & nVID & "," & nSEQ & "," & nST & ",'" & cLOB & "'," & nCM & ")"
		oConn.Execute(cSQL)
		cError = CheckADOErrors(oConn,"Account Vendor: ADD VENDOR")
		If cError = "" Then 
			nUpdateStatus = 1
		else
			nUpdateStatus = -1
		end if
	Else
		nUpdateStatus = -1
	End If
	oConn.Close
	Response.Write ("<script>window.close();</script>")
	Response.End 
End If

cAccVendorID = trim(Request.QueryString("AVID"))
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open CONNECT_STRING

'cSQL = "SELECT AV.*, ST.TYPE, CM.NAME FROM ACCOUNT_VENDOR AV, SERVICE_TYPE ST, CONTACT_METHOD CM " & _
'		"WHERE AV.ACCOUNT_VENDOR_ID = " & cAccVendorID & _
'		" AND AV.SERVICE_TYPE_ID = ST.SERVICE_TYPE_ID " & _
'		"AND AV.CONTACT_METHOD_ID = CM.CONTACT_METHOD_ID"
'Set oRS = oConn.Execute(cSQL)
'If Not oRS.EOF then
'	RSLOB_CD = oRS("LOB")
'	RSSERV_TYPE = oRS("TYPE")
'	RSCONT_METH = oRS("NAME")
'end if
'oRS.close
'oConn.close
'set oRS = nothing
'set oConn = nothing

Function NextPkey( TableName, ColName )
	NextSQL = NextSQL & "{call Designer.GetValidSeq('" & TableName & "', '" & ColName & "', {resultset 1, outResult})}"
	Set NextRS = oConn.Execute(NextSQL)
	NextPkey = NextRS("outResult") 
End Function
	
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<STYLE TYPE="text/css">
HTML {width: 270pt; height: 150pt}
</STYLE>
<SCRIPT LANGUAGE="Javascript">

function AVSearchObj()
{
	this.Sequence = "";
	this.VendorID = "";
	this.NetworkID = "";
	this.Selected = false;	
}

function VendorSearch()
{
	this.VID = "";
	this.Selected = false;
}

function NetworkSearch()
{
	this.NetID = "";
	this.Selected = false;
}

var VendorSearchObj = new VendorSearch();
var NetObj = new NetworkSearch();
var AccVendObj = new AVSearchObj();
var lAlreadySaved = false;

</SCRIPT>

<script LANGUAGE="vbScript" FOR="window" EVENT="onload">
<%
if nUpdateStatus = 1 then
%>
	document.all.SpanStatus.innerHTML = "Update successful."
	lAlreadySaved = true
<%
elseif nUpdateStatus = -1 then
%>
	document.all.SpanStatus.innerHTML = "<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful."
<%	
end if
%>	
</script>


<SCRIPT LANGUAGE=vbscript>
<!--
sub VBStop()
STOP
end sub

sub AttachVendor()

	VendorSearchObj.Selected = false
	strURL = "..\Vendors\VendorMaintenance.asp?SECURITYPRIV=FNSD_NETWORKS&CONTAINERTYPE=MODAL"
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	showModalDialog strURL, VendorSearchObj, "center"
	If VendorSearchObj.Selected then
		document.all.Vendor_ID.value = VendorSearchObj.VID
	end if
end sub

sub AttachNet()
	NetObj.Selected = false
	strURL = "networkSearch-f.asp"
	showModalDialog  strURL, NetObj, "dialogWidth=450px; dialogHeight=550px; center=yes"
	If NetObj.Selected then
		document.all.Network_ID.value = NetObj.NetID
	end if
end sub
-->
</SCRIPT>
<script LANGUAGE="vbScript" FOR="BtnSave" EVENT="onclick">
	if not lAlreadySaved then
		if document.all.TxtSequence.value = "" then
			cErrMsg = "Sequence is a required field." & chr(10) & chr(13)
		end if
		if cint(document.all.Network_ID.value) = 0 and cint(document.all.Vendor_ID.value) = 0 then
			cErrMsg = cErrMsg & "A Vendor ID or a Network ID must be supplied."
		end if
		if document.all.TxtContMeth.value = "" then
			cErrMsg = "Contact method is a required field." & chr(10) & chr(13)
		end if
		if len(cErrMsg) <> 0 then
			msgbox cErrMsg 
		else
			self.location.href = "VendorSummary2.asp?NEW=&VID=" & document.all.Vendor_ID.value & "&NID=" & document.all.Network_ID.value & "&SEQ=" & document.all.TxtSequence.value & "&CM=" & document.all.TxtContMeth.value & "&<%=Request.Querystring%>"
			window.close 
		end if
	end if
</script>

<script LANGUAGE="vbscript" FOR="BtnClose" EVENT="onclick">
	window.close
</script>

</HEAD>
<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<table CELLPADDING=0 LANGUAGE="JScript" ONDRAGSTART="return false;" class="Label">
<tr><td colspan=2>&nbsp</td></tr>
<tr>

<td VALIGN="CENTER" WIDTH="5">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER" ALT="View Status Report">
</td>
<td width="485">
:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>
<table CELLPADDING=0 class="LABEL">
<tr>
<td><b>LOB:</b>&nbsp;<span id="spanAHSID"><%=cLOB%></span></td>
<td>&nbsp</td>
<td><b>Serv. Type:</b>&nbsp;<span id="spanAHSID"><%=nST%></span></td>
</tr> 
<tr>
<td>&nbsp</td>
</tr>
<tr>
&nbsp
</tr>
</table>
<table>
<tr>
<TD CLASS=LABEL>Sequence:<BR><INPUT TYPE=TEXT SIZE=4 NAME=TxtSequence CLASS=LABEL></TD>
<td CLASS="LABEL">Vendor ID:<br>
<img SRC="../Images/Attach.gif" align=absmiddle ID="BtnATTACHVendor" STYLE="CURSOR:HAND" ALT="Attach Vendor" OnClick="AttachVendor()" WIDTH="16" HEIGHT="16">&nbsp
<input TYPE="TEXT" READONLY ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" STYLE="BACKGROUND-COLOR:SILVER" NAME="Vendor_ID" CLASS="LABEL" SIZE="10" MAXLENGTH="10" VALUE="0"></td>
<td CLASS="LABEL">&nbsp</td>
<td CLASS="LABEL">OR</td>
<td CLASS="LABEL">&nbsp</td>
<td CLASS="LABEL">Network ID:<br>
<img SRC="../Images/Attach.gif"  align=absmiddle ID="BtnATTACHNetwork" STYLE="CURSOR:HAND" ALT="Attach Network" OnClick="AttachNet()" WIDTH="16" HEIGHT="16">
<input TYPE="TEXT" READONLY ONKEYPRESS="VBScript::Control_OnChange" ONCHANGE="VBScript::Control_OnChange" STYLE="BACKGROUND-COLOR:SILVER" NAME="Network_ID" CLASS="LABEL" SIZE="10" MAXLENGTH="10" VALUE="0"></td>
<td CLASS="LABEL">&nbsp</td>
<td CLASS="LABEL">Contact method:<br>
	<select ID="TxtContMeth" CLASS="LABEL" ScrnBtn="TRUE">
	<option VALUE>
	<%
	cSQL = "SELECT * FROM CONTACT_METHOD WHERE NAME IS NOT NULL"
	Set oRS = oConn.Execute(cSQL)
	Do While Not oRS.EOF
	%>
	<option VALUE="<%=oRS("CONTACT_METHOD_ID")%>"><%= oRS("NAME") %>
	<%
	oRS.MoveNext
	Loop
	oRS.CLose
	oConn.close
	set oConn = nothing
	set oRS = nothing
	%>
</td>
</tr>
</table>
<table>
<tr>
&nbsp
</tr>
<tr>
<td>&nbsp</td>
<td CLASS="LABEL">
<button CLASS="StdButton" NAME="BtnSave" ACCESSKEY="N"><u>S</u>ave</button></td>
<td CLASS="LABEL">
<button CLASS="StdButton" NAME="BtnClose" ACCESSKEY="N"><u>C</u>ancel</button></td>
</tr>
</table>

</BODY>
</HTML>

