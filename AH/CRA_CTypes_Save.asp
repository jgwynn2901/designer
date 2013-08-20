<!--#include file="..\lib\genericSQL.asp"-->
<HTML>
<HEAD>
<%
dim cSQL, oRS, cCovCodeID
dim cAHSID, cCovCode, cLOB, cStates
dim cRespString, aStates, lIsEdit, cSC

cAHSID = Request.Querystring("AHSID")
cCovCode = Request.Querystring("CC")
cCovCodeID = Request.Querystring("CC_ID")
cLOB = Request.Querystring("LOB")
cStates = Request.Querystring("STATES")
aStates = split(cStates, "||")
cSC = Request.Querystring("SC")

lIsEdit = cCovCodeID <> "" 
if lIsEdit then
	'	delete existing records
	removeStates cCovCodeID
	addStates cCovCodeID, aStates
	removeSC cCovCodeID
	if len(cSC) <> 0 then
		addSC cSC, cCovCodeID
	end if
else
	cSQL = "{call Designer.GetValidSeq('COVERAGE_CODE', 'COVERAGE_CODE_ID', {resultset 1, outResult})}"
	Set oRS = Conn.Execute(cSQL)
	cCovCodeID = oRS("outResult")
	cSQL = "Insert Into COVERAGE_CODE VALUES ("
	cSQL = cSQL & cCovCodeID & ", "
	cSQL = cSQL & cAHSID & ", '"
	cSQL = cSQL & cLOB & "', '"
	cSQL = cSQL & cCovCode & "')"
	oRS.close
	Conn.Execute cSQL
	addStates cCovCodeID, aStates
	if len(cSC) <> 0 then	
		addSC cSC, cCovCodeID
	end if
end if	
Conn.Close
set Conn = nothing
cRespString = "<script language='vbscript'>window.close</script>"
response.write cRespString 

sub removeStates(nCovCodeID)
dim cSQL

cSQL = "Delete From BENEFIT_STATE Where COVERAGE_CODE_ID=" & nCovCodeID
Conn.execute cSQL
end sub

sub addStates(nCovCodeID, aStates)
dim cSQL, nTop, x

nTop = ubound(aStates)
for x=0 to nTop
	cSQL = "Insert Into BENEFIT_STATE Values(" & nCovCodeID & ",'" & aStates(x) & "')"
	Conn.execute cSQL
next
end sub

sub removeSC(nCovCodeID)
dim cSQL

cSQL = "Delete From SPECIAL_COMMENTS Where COVERAGE_CODE_ID=" & nCovCodeID
Conn.execute cSQL
end sub

sub addSC(cSC, nCovCodeID)
dim cSQL

cSQL = "Insert Into SPECIAL_COMMENTS Values(" & nCovCodeID & ",'Y','" & cSC & "')"
Conn.execute cSQL
end sub
%>
</HEAD>
<BODY>
</BODY>
</HTML>