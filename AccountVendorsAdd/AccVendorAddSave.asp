<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<HTML>
<HEAD>
<SCRIPT>
<%
On Error Resume Next
dim nAccVendorID, oConn, oRS, cSQL
dim nVID, nNID, nSEQ, nAHSID, cLOB, nST, nCM
dim cError, lUpdateOK, lIsEdit, nAVID

nVID = Cint(Request.Form("VID"))
nNID = Cint(Request.Form("NID"))
nSEQ = Request.Form("SEQ")	
nAHSID = Request.Form("AHSID")
cLOB = Request.Form("LOB")
nST = Request.Form("ST")
nCM = Request.Form("CM")
lIsEdit = Request.Form("ACTION") = "EDIT"
nAVID = Request.Form("AVID")
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open CONNECT_STRING
if lIsEdit then
	cSQL = "UPDATE ACCOUNT_VENDOR SET CONTACT_METHOD_ID = " & nCM & ", SEQUENCE = " & nSEQ & ", VENDOR_ID = " & nVID & ", NETWORK_ID = " & nNID & " WHERE ACCOUNT_VENDOR_ID = " & nAVID
	oConn.Execute(cSQL)
	cError = CheckADOErrors(oConn,"Account Vendor: UPD VENDOR")
else
	nNewID = CLng(NextPkey("ACCOUNT_VENDOR","ACCOUNT_VENDOR_ID"))
	If nNewID > 0 Then
		cSQL = "Insert into ACCOUNT_VENDOR values (" & nNewID & "," & nAHSID & "," & nNID & "," & nVID & "," & nSEQ & "," & nST & ",'" & cLOB & "'," & nCM & ",'')"
		oConn.Execute(cSQL)
		cError = CheckADOErrors(oConn,"Account Vendor: ADD VENDOR")
	Else
		cError = "Unable to obtain next primary key for ACCOUNT_VENDOR table."
	End If
end if
oConn.Close
	
If cError <> "" Then
	LogStatusGroupBegin
	LogStatus S_ERROR, strError, "ACCOUNT_VENDOR", "", 0, ""
	LogStatusGroupEnd %>
	parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.");
	parent.frames("WORKAREA").SetDirty();
	parent.frames('WORKAREA').SetStatusInfoAvailableFlag(true);		
<%		
Else
	LogStatusGroupBegin
	LogStatusGroupEnd %>
	parent.frames("WORKAREA").UpdateStatus("Update successful.");
	parent.frames("WORKAREA").ClearDirty();
	parent.frames('WORKAREA').SetStatusInfoAvailableFlag(false);				
<%		
End If

Function NextPkey( TableName, ColName )
	NextSQL = NextSQL & "{call Designer.GetValidSeq('" & TableName & "', '" & ColName & "', {resultset 1, outResult})}"
	Set NextRS = oConn.Execute(NextSQL)
	NextPkey = NextRS("outResult") 
End Function
%>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
