<!--#include file="..\lib\common.inc"-->
<%	
Response.Expires = 0

dim cSQL, oRS, oConn, cAHSID
dim nST, cLOB
dim nVendorID, nNetworkID
dim cNID, cVID
dim cVendorName, cNetworkName
dim cContactMethod

cAHSID = trim(Request.QueryString("AHSID"))
nST = Request.QueryString("ST")
cLOB = Request.QueryString("LOB")
cAVID = Request.QueryString("AVID")

if cAVID <> "NEW" then
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open CONNECT_STRING
	cSQL = "SELECT * FROM ACCOUNT_VENDOR WHERE ACCNT_HRCY_STEP_ID = " & cAHSID & " AND SERVICE_TYPE_ID=" & nST & " AND LOB='" & cLOB & "' AND CONTACT_METHOD_ID <> -1 ORDER BY SEQUENCE"
	Set oRS = oConn.Execute(cSQL)
end if	
%>
<!--#include file="..\lib\RenderTextinc.asp"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Vendors Details</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
Function GetSelectedAVID
	dim idx
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedAVID = document.all.tblFields.rows(idx).getAttribute("AVID")
	Else
		GetSelectedAVID = ""
	End If
End Function
</script>
<!--#include file="..\lib\tablecommon.inc"-->
</head>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0" >
<div align="LEFT" style="height:'100%';width:'90%'">
<table cellPadding="2" rules=all  cellSpacing="0" scrolling="auto" ID="tblFields" name="tblFields" width="30%">
	<thead CLASS="ResultHeader">
		<tr align="left">
			<td class="thd"><nobr>Sequence</td>
			<td class="thd"><nobr>VID</td>			
			<td class="thd"><nobr>Vendor Name</td>			
			<td class="thd"><nobr>NID</td>			
			<td class="thd"><nobr>Network Name</td>
			<td class="thd"><nobr>Contact Method</td>						
		</tr>
	</thead>
	<tbody ID="TableRows">
<%
if cAVID <> "NEW" then
	Do While Not oRS.EOF And Not oRS.BOF
		nVendorID = 0
		if not isNull(oRS("VENDOR_ID")) then
			nVendorID = Cint(oRS("VENDOR_ID"))
		end if
		nNetworkID = 0
		if not isNull(oRS("NETWORK_ID")) then
			nNetworkID = Cint(oRS("NETWORK_ID"))
		end if
		cVendorName = ""
		cNetworkName = ""
		cNID = ""
		cVID = ""
		cSQL = "SELECT NAME FROM CONTACT_METHOD WHERE CONTACT_METHOD_ID=" & oRS("CONTACT_METHOD_ID")
		Set oRS0 = oConn.Execute(cSQL)
		cContactMethod = oRS0("NAME")
		oRS0.close
		if nVendorID <> 0 then
			cSQL = "SELECT * FROM VENDOR WHERE VENDOR_ID = " & nVendorID
			Set oRS0 = oConn.Execute(cSQL)
			cVendorName = oRS0("NAME")
			cVID = CStr(nVendorID)
		else
			cSQL = "SELECT * FROM NETWORK WHERE NETWORK_ID = " & nNetworkID
			Set oRS0 = oConn.Execute(cSQL)
			cNetworkName = oRS0("NAME")
			cNID = CStr(nNetworkID)
		end if
		oRS0.close
		Set oRS0 = Nothing
%>	 
	<tr ID="FieldRow" CLASS="ResultRow"  DYNKEY="1" OnClick="Javascript:multiselect(this);" AVID=<%=oRS("ACCOUNT_VENDOR_ID")%>>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(oRS("SEQUENCE"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(cVID)%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(cVendorName)%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(cNID)%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(cNetworkName)%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(cContactMethod)%></td>	
	</tr>
<%
		oRS.MoveNext
	Loop
	oRS.Close
	Set RS = Nothing
	oConn.Close
	Set oConn = Nothing
end if	
%>

</tbody>
</table>
</div>
</BODY>
</HTML>


