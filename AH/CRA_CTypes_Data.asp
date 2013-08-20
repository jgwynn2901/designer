<!--#include file="..\lib\genericSQL.asp"-->
<%
Response.Expires = 0
dim cAHSID, cCovTypeID, cLOB, cCC, cStates, cSC, cLOBExcl

cAHSID = Request.QueryString("AHSID") 
cCovCodeID = Request.QueryString("EDIT")
if len( cCovCodeID ) <> 0 then
	'	edit mode
	cLOB = Request.QueryString("LOB")
	cCC = Request.QueryString("CC")
	cStates = Request.QueryString("STATES")
	cSC = Request.QueryString("SC")
	cLOBExcl = ""
else	
	cLOBExcl = Request.QueryString("EXCL")
end if	
if not isEmpty(Request.QueryString("SAVE")) then
	'	save
	response.Redirect "CRA_CTypes_Save.asp?" & Request.QueryString
end if
%>
<HTML>
<HEAD>
	<META name="VI60_defaultClientScript" content="VBScript">
	<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language=jscript>
function SelectOption(objSelect, strValue)
{
	var i, iRetVal=-1;

	for (i=0; i < objSelect.length; i ++)
	{
		if (strValue == objSelect(i).value)
		{
			objSelect(i).selected = true;
			return;
		}
	}
}

</script>

<SCRIPT ID="clientEventHandlersVBS" LANGUAGE="vbscript">
dim aLOBExcl

sub LOB_OnChange
dim oOpts, oOption, aCC

set oOpts = document.all.Cov_Code.options
for each oOption in oOpts
	oOpts.remove(0)
next

select case document.all.LOB_CD.value
	case "CAU"
		aCC = Array("AL", "HE", "VC", "VD")
		addOptions oOpts, chkCCS(document.all.LOB_CD.value, aCC)
		'AL, HE, VC, VD
	case "CLI"
		aCC = Array("GL", "SU")
		addOptions oOpts, chkCCS(document.all.LOB_CD.value, aCC)
		'GL, SU
	case "CPR"
		aCC = Array("PL")
		addOptions oOpts, chkCCS(document.all.LOB_CD.value, aCC)
		'PL
	case "DIS"
		aCC = Array("DI", "DM")
		addOptions oOpts, chkCCS(document.all.LOB_CD.value, aCC)
		'DI, DM
	case "INF"
		aCC = Array("GI", "GM")
		addOptions oOpts, chkCCS(document.all.LOB_CD.value, aCC)
		'GI, GM
	case "WOR"
		aCC = Array("WC")
		addOptions oOpts, chkCCS(document.all.LOB_CD.value, aCC)
		'WC
end select
end sub

sub addOptions(oColl, aCovCodes)
dim oOption, nTop, x

nTop = uBound(aCovCodes)
for x=0 to nTop
	set oOption = document.createElement("OPTION")
	oColl.add(oOption)
	oOption.innerText = aCovCodes(x)
	oOption.value = aCovCodes(x)
next
end sub

function chkCCS(cLOB, aCC)
dim nTop, x, lFirst, aResult

aResult = Array()
nTop = ubound(aCC)
lFirst = true
for x=0 to nTop
	if AScan(aLOBExcl, cLOB & "|" & aCC(x)) = -1 then
		if lFirst then
			redim aResult(0)
			lFirst = false
		else
			redim preserve aResult(ubound(aResult)+1)
		end if
		aResult(ubound(aResult)) = aCC(x)
	end if
next
chkCCS = aResult
end function

Function AScan(aValues, cSearch)
Dim x
Dim nTop

AScan = -1
nTop = UBound(aValues)
For x = 0 To nTop
	If aValues(x) = cSearch Then
		AScan = x
		Exit For
	End If
Next
End Function
    
Sub ExeSave()
dim cHRef, cMsg

'msgbox getStates
if len(document.all.LOB_CD.value) = 0 then
	cMsg = "LOB is a required field." & vbcrlf
end if
if len(document.all.Cov_Code.value) = 0 then
	cMsg = cMsg & "Coverage Code is a required field." & vbcrlf
end if
if len(document.all.States.value) = 0 then
	cMsg = cMsg & "State is a required field." & vbcrlf
end if
if len(cMsg) <> 0 then
	msgbox cMsg
else	
	cHRef= "CRA_CTypes_Data.asp?AHSID=<%=cAHSID%>&CC_ID=<%=cCovCodeID%>&SAVE="
	document.location.href = cHRef & "&LOB=" & document.all.LOB_CD.value & "&CC=" & document.all.Cov_Code.value & "&STATES=" & getStates & "&SC=" & document.all.txtSpecComm.value
end if
End Sub

function getStates()
dim oOpts, oOption 

getStates = ""
set oOpts = document.all.States.options
for each oOption in oOpts
	if oOption.selected then
		if len(getStates) <> 0 then
			getStates = getStates & "||"
		end if
		getStates = getStates & oOption.value
	end if
next
end function

function setStates(aList)
dim oOpts, oOption, nTop, x

nTop = ubound(aList)
set oOpts = document.all.States.options
for x=0 to nTop
	for each oOption in oOpts
		if oOption.value = aList(x) then
			oOption.selected = true
			exit for
		end if
	next
next
end function
</SCRIPT>
</HEAD>
	<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="#d6cfbd" ScreenDirty="NO" ScreenMode="RW">
			<table class="Label" ID="Table3">
				<tr>
					<td VALIGN="CENTER" WIDTH="5">
						<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER" ALT="View Status Report">
					</td>
					<td width="485">
						:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
					</td>
				</tr>
			</table>
			<table CLASS="LABEL" CELLPADDING="0" CELLSPACING="0" id="TblControls" width="495">
				<tr>
					<td width="560">
						<table class="LABEL" ID="Table4">
							<tr>
								<td>A.H.S. ID:&nbsp;<span id="spanAHSID">81</span></td>
							</tr>
						</table>
						&nbsp
	<table ID="Table1">
	<tr>
	<td CLASS="LABEL" width="30" valign="top"><font size="1">LOB:</font></td>
	<td CLASS="LABEL" width="70" valign="top">
	<select NAME="LOB_CD" CLASS="LABEL" ScrnBtn="TRUE" ONCHANGE="VBScript::LOB_OnChange" ID="Select1" style="font-family: Verdana; font-size: 9pt">
	<option VALUE>
	<%
	cSQL = "SELECT * FROM LOB WHERE LOB_CD IS NOT NULL"
	Set oRS = Conn.Execute(cSQL)
	Do While Not oRS.EOF
	%>
	<option VALUE="<%= oRS("LOB_CD") %>"><%= oRS("LOB_CD") %>
	<%
	oRS.MoveNext
	Loop
	oRS.CLose
	Conn.close
	%>
	</select></td>
						
								<td CLASS="LABEL" width="111" valign="top"><font size="1">Coverage Code:</font></td>
								<td CLASS="LABEL"  width="111" valign="top">
									<select NAME="Cov_Code" CLASS="LABEL" STYLE="WIDTH:100" ScrnBtn="TRUE" ID="Select3" style="font-family: Verdana; font-size: 9pt" >
									
									</select>
								</td>
								
<td CLASS="LABEL" width="50" valign="top"><font size="1">State:</font></td>
<td CLASS="LABEL">
			<select NAME="States" CLASS="LABEL" tabindex="8" size="12" multiple style="font-family: Verdana; font-size: 9pt" ID="Select2">
				<OPTION VALUE='AB'>AB
				<OPTION VALUE='AK'>AK
				<OPTION VALUE='AL'>AL
				<OPTION VALUE='AR'>AR
				<OPTION VALUE='AZ'>AZ
				<OPTION VALUE='BC'>BC
				<OPTION VALUE='CA'>CA
				<OPTION VALUE='CN'>CN
				<OPTION VALUE='CO'>CO
				<OPTION VALUE='CT'>CT
				<OPTION VALUE='DC'>DC
				<OPTION VALUE='DE'>DE
				<OPTION VALUE='FL'>FL
				<OPTION VALUE='FM'>FM
				<OPTION VALUE='GA'>GA
				<OPTION VALUE='GU'>GU
				<OPTION VALUE='HI'>HI
				<OPTION VALUE='IA'>IA
				<OPTION VALUE='ID'>ID
				<OPTION VALUE='IL'>IL
				<OPTION VALUE='IN'>IN
				<OPTION VALUE='KS'>KS
				<OPTION VALUE='KY'>KY
				<OPTION VALUE='LA'>LA
				<OPTION VALUE='MA'>MA
				<OPTION VALUE='MB'>MB
				<OPTION VALUE='MD'>MD
				<OPTION VALUE='ME'>ME
				<OPTION VALUE='MH'>MH
				<OPTION VALUE='MI'>MI
				<OPTION VALUE='MN'>MN
				<OPTION VALUE='MO'>MO
				<OPTION VALUE='MS'>MS
				<OPTION VALUE='MT'>MT
				<OPTION VALUE='NB'>NB
				<OPTION VALUE='NC'>NC
				<OPTION VALUE='ND'>ND
				<OPTION VALUE='NE'>NE
				<OPTION VALUE='NH'>NH
				<OPTION VALUE='NJ'>NJ
				<OPTION VALUE='NL'>NL
				<OPTION VALUE='NM'>NM
				<OPTION VALUE='NS'>NS
				<OPTION VALUE='NT'>NT
				<OPTION VALUE='NU'>NU
				<OPTION VALUE='NV'>NV
				<OPTION VALUE='NY'>NY
				<OPTION VALUE='OH'>OH
				<OPTION VALUE='OK'>OK
				<OPTION VALUE='ON'>ON
				<OPTION VALUE='OR'>OR
				<OPTION VALUE='PA'>PA
				<OPTION VALUE='PE'>PE
				<OPTION VALUE='PR'>PR
				<OPTION VALUE='PW'>PW
				<OPTION VALUE='QC'>QC
				<OPTION VALUE='RI'>RI
				<OPTION VALUE='SC'>SC
				<OPTION VALUE='SD'>SD
				<OPTION VALUE='SK'>SK
				<OPTION VALUE='TN'>TN
				<OPTION VALUE='TX'>TX
				<OPTION VALUE='UT'>UT
				<OPTION VALUE='VA'>VA
				<OPTION VALUE='VI'>VI
				<OPTION VALUE='VT'>VT
				<OPTION VALUE='WA'>WA
				<OPTION VALUE='WI'>WI
				<OPTION VALUE='WV'>WV
				<OPTION VALUE='WY'>WY 
				<OPTION VALUE='YT'>YT
			</select>

&nbsp;</TD>
								
							</tr>
						</table>
						&nbsp
						<table ID="Table2" width="560">
							<tr>
							<td CLASS="LABEL"><font size="1">Special Comments</font></td>
</tr><tr>
								<td CLASS="LABEL" width="500" >
                                  <textarea rows="4" name="txtSpecComm" cols="42" ID="Textarea1"></textarea>
                                </td>
							</tr>
						</table>
&nbsp							
<script language=vbscript>
Sub window_onload
dim aStates

<%
if len(cCovCodeID) <> 0 then
%>
	aLOBExcl = Array()
	aStates = split("<%=cStates%>", "||")
	document.all.LOB_CD.disabled = true
	document.all.Cov_Code.disabled = true
	SelectOption document.all.LOB_CD,"<%=cLOB%>"
	LOB_OnChange
	SelectOption document.all.Cov_Code,"<%=cCC%>"
	setStates aStates
	document.all.txtSpecComm.value = "<%=cSC%>"
<%
else
%>	
	aLOBExcl = split("<%=cLOBExcl%>", ",")
<%
end if
%>	
End Sub
</script>
</body>
</HTML>
