<%	Response.Expires = 0
	On Error Resume Next

%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptInc.asp"-->
<!--#include file="..\lib\commonError.inc"-->
<!--#include file="..\lib\CheckLogicExpression.inc"-->

<HTML>
<HEAD>
<SCRIPT LANGUAGE="VBScript">

<%
dim lErrors, lValid
dim cRIDText
dim oConn, oRS
dim cSQL
dim cError

lErrors = True
		

cRIDText    = Request.Form("TxtDictText")
cRID        = Request.Form("RID")
cDeleteFlag = Request.Form("DeleteFlag")	

'cRIDText = Replace(cRIDText,"'","''")
'cRIDText = Replace(cRIDText,CHR(13)," ")
'cRIDText = Replace(cRIDText,CHR(10)," ")
	

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open CONNECT_STRING

If cRID = "NEW" Then ' ADD
	cSQL = "{call Designer_3.SP_ADD_WORD(" &_
					"'" & cRIDText & "', {resultset  1, outResult})}"
Else	' UPDATE or delete
    if cDeleteFlag = "Y" then
       cSQL = "{call Designer_3.SP_DELETE_WORD(" &_
       	            "'" & cRIDText & "', {resultset  1, outResult})}"
					
    else
	   cSQL = "{call Designer_3.SP_UPDATE_WORD(" &_
									"'" & cRIDText & "',"  &_
					"'" & cRID & "', {resultset  1, outResult})}"
	end if
End If
				
Set oRS = Server.CreateObject("ADODB.Recordset")
oRS.Open cSQL, oConn, adOpenStatic, adLockReadOnly, adCmdText
if oRS("outResult")="0" then
	lErrors = true
else
	If cRID = "" Then 
		cRID = oRS("outResult") %>
		parent.frames("WORKAREA").SetRID(<%=cRID%>)
	<%end if
	lErrors = false
end if
cError = CheckADOErrors(oConn,"Word Save" )
oRS.Close
Set oRS = Nothing
oConn.Close
Set oConn = Nothing
If lErrors Or cError <> "" Then 
		LogStatusGroupBegin
		LogStatus S_ERROR, cError, "WORD", "", 0, ""
		LogStatusGroupEnd
%>		
parent.frames("WORKAREA").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.")	
parent.frames("WORKAREA").SetDirty
parent.frames("WORKAREA").SetStatusInfoAvailableFlag(true)		
<%
	Else
		LogStatusGroupBegin
		LogStatusGroupEnd
%>
parent.frames("WORKAREA").UpdateStatus("Update successful.")			
parent.frames("WORKAREA").ClearDirty
parent.frames("WORKAREA").SetStatusInfoAvailableFlag(false)

<%	End If 
			
%>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
</HTML>


