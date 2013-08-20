<!--#include file="..\lib\genericSQL.asp"-->
<HTML>
<HEAD>
<%
dim cSQL, oRS
dim cAHSID, cCovCodeID, cPriTerrCodeID, cSecTerrCodeID, cZIPCodeLocation
dim cRespString, cBranchOvrRuleID

cAHSID = Request.Querystring("AHSID")
cCovCodeID = Request.Querystring("Cov_Code")
cPriTerrCodeID = Request.Querystring("PrTerr_Code")
cSecTerrCodeID = Request.Querystring("SecTerr_Code")
cZIPCodeLocation = Request.Querystring("ZIPLoc_Code")
cBranchOvrRuleID = Request.Querystring("BOR")
if cBranchOvrRuleID = "0" then
	cBranchOvrRuleID = "null"
end if	

if len(cSecTerrCodeID) = 0 then	
	cSecTerrCodeID = "null"
end if
if Request.Querystring("BT_ID") <> "" then
	' edit mode
	cSQL = "UPDATE CRA_BRANCH_TYPES SET COVERAGE_CODE_ID=" & cCovCodeID & ",PRIMARY_TERRITORY_ID=" & cPriTerrCodeID & ",SECONDARY_TERRITORY_ID=" & cSecTerrCodeID & ",ZIPCODE_LOCATION='" & cZIPCodeLocation & "',BRANCH_OVERRIDE_RULE_ID=" & cBranchOvrRuleID
	cSQL = cSQL & " WHERE CRA_BRANCH_TYPES_ID=" & Request.Querystring("BT_ID")
else
	cSQL = "{call Designer.GetValidSeq('CRA_BRANCH_TYPES', 'CRA_BRANCH_TYPES_ID', {resultset 1, outResult})}"
	Set oRS = Conn.Execute(cSQL)
	cSQL = "INSERT INTO CRA_BRANCH_TYPES VALUES ("
	cSQL = cSQL & oRS("outResult") & ", "
	cSQL = cSQL & cAHSID & ", "
	cSQL = cSQL & cCovCodeID & ","
	cSQL = cSQL & cPriTerrCodeID & ","
	cSQL = cSQL & cSecTerrCodeID & ",'"
	cSQL = cSQL & cZIPCodeLocation & "'," 
	cSQL = cSQL & cBranchOvrRuleID & ")"
	oRS.close
end if	
Conn.Execute(cSQL)
Conn.Close
set Conn = nothing
cRespString = "<script language='vbscript'>window.close</script>"
response.write cRespString 
%>
</HEAD>
<BODY>
</BODY>
</HTML>