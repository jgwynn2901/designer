<!--#include file="..\lib\common.inc"-->
<%
Function NextPkey( TableName, ColName )
	NextSQL = ""
	'NextSQL = NextSQL & "SELECT " & Trim(TableName) & "_SEQ.NextVal As NextID FROM DUAL"
	NextSQL = NextSQL & "{call Designer.GetValidSeq('" & TableName & "', '" & ColName &"', {resultset 1, outResult})}"
	Set NextRS = Conn.Execute(NextSQL)
	NextPkey = NextRS("outResult") 
End Function


Function Split(UsrString, Delimeter)
Dim posd, i, Tempstr, lngstr
Dim RArray(200)

i = 1
posd = 1
lngstr = Len(UsrString)
Tempstr = Trim(UsrString)

Do While posd <> 0
    posd = InStr(Tempstr, Delimeter)
    If posd <> 0 Then
        Rarray(i) = Trim(Left(Tempstr, posd - 1))
    Else
        Rarray(i) = Trim(Tempstr)
    End If
    Tempstr = Right(UsrString, lngstr - posd)
    lngstr = Len(Tempstr)
    i = i + 1
Loop
Split = Rarray
End Function

	'Set X = Server.CreateObject("IDGen.IDGen.1")
	'X.TableName = "OUTPUT_FIELD"
	SQLUPDATE = Split(Request.Form("TxtUpdateData") , "|")
	SQLINSERT = Split(Request.Form("TxtInsertData") , "|")
	SQLDELETE = Split(Request.Form("TxtDeleteData") , "|")
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnectionString = CONNECT_STRING
	Conn.Open ConnectionString
	For a = 1 to Cint(Request.Form("INCOUNT")) Step 1
		SQLINSERT(a) = SQLINSERT(a) & NextPkey("OUTPUT_FIELD", "OUTPUT_FLD_ID") & ")"
		Set RS = Conn.Execute(SQLINSERT(a))
	Next
	
	For i = 1 to Cint(Request.Form("UPCOUNT")) Step 1
		Set RS = Conn.Execute(SQLUPDATE(i))
	Next
	
	For z = 1 to Cint(Request.Form("DELCOUNT")) Step 1
		Set RS = Conn.Execute(SQLDELETE(z))
	Next
	Status = "Saved"
 %>
<HTML>
<HEAD>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
	'top.frames("WORKAREA").location.href = "RP_TabWindow.asp?OPID=<%= Request.QueryString("OPID") %>"
End Sub

-->
</SCRIPT>
</HEAD>
<BODY>

</BODY>
