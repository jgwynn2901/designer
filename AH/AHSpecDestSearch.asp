<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<html>

<head>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Search</title>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--

<!--#include file="..\lib\Help.asp"-->

Sub window_onload
	document.all.SearchType(0).checked = True
End Sub

Sub BtnClear_OnClick
	document.all.CITY.Value = ""
	document.all.LNAME.Value = ""
	document.all.FNAME.Value = ""    
	document.all.MI.Value = ""  
	document.all.SD.Value = ""  
	document.all.STATE.Value = "" 
	document.all.PHONE.Value = ""
	document.all.ZIP.Value = ""
End Sub

Sub UpdateStatus(inStatus)
	document.all.SpanStatus.innerHTML = inStatus
End Sub

Sub PostTo(strURL)
	dim cSDID, aTemp
	
	cSDID = Parent.frames("WORKAREA").GetSDID
	aTemp = Split(cSDID, "||")
	If UBound(aTemp) >= 0 Then 
		document.all.SDID.value = aTemp(0)
	Else		
		document.all.SDID.value = ""
	End If
	FrmSearch.action = "SpecDestDetails-f.asp"
	FrmSearch.method = "GET"	
	FrmSearch.target = "_parent"	
	FrmSearch.submit
End Sub

-->
</script>
<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function FrmSearch_onsubmit() {
document.all.SpanStatus.innerHTML = "<%=MSG_SEARCH%>";
}

//-->
</SCRIPT>
</head>

<body rightmargin="0" leftmargin="0" bottommargin="0" topmargin="0"
BGCOLOR="<%=BODYBGCOLOR%>">

<form NAME="FrmSearch" ACTION="AHSpecDestResults.asp" METHOD="GET" TARGET="WORKAREA" LANGUAGE="javascript" onsubmit="return FrmSearch_onsubmit();">
  <table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
    <tr>
      <td colspan="2" HEIGHT="4"></td>
    </tr>
    <tr>
      <td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Specific Destination
      Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img
      SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help"
      OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></nobr></td>
      <td HEIGHT="5" ALIGN="LEFT"><table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
        <tr>
          <td WIDTH="3" HEIGHT="4"></td>
          <td WIDTH="300" HEIGHT="4"></td>
        </tr>
        <tr>
          <td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
          <td WIDTH="300" HEIGHT="8"></td>
        </tr>
      </table>
      </td>
    </tr>
    <tr>
      <td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td>
    </tr>
    <tr>
      <td colspan="2" HEIGHT="1"></td>
    </tr>
  </table>
  <table LANGUAGE="JScript" ONDRAGSTART="return false;" class="Label" cellspacing="0"
  cellpadding="0">
    <tr>
      <td VALIGN="CENTER" WIDTH="5"><img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16"
      height="16" VALIGN="CENTER" Title="View Status Report"> </td>
      <td width="485">:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
      </td>
    </tr>
  </table>
  <table CELLPADDING="0" CELLSPACING="0" BORDER="0" WIDTH="100%">
    <tr>
		<td width="40%">
			<table border="0" CELLPADDING="3" CELLSPACING="3" WIDTH="70%">
				<tr>
					<td CLASS="LABEL" VALIGN="BOTTOM" colspan="2" width="40%">Spec. Dest. ID:<br>
					<input CLASS="LABEL" TYPE="TEXT" NAME="SD" SIZE="10"></td>
					<td CLASS="LABEL" align="right"  width="5%"><button type="submit" CLASS="STDBUTTON" ACCESSKEY="C" NAME="BtnSearch">Sear<u>c</u>h</button></td>
					<td CLASS="LABEL" align="left"  width="50%"><button CLASS="STDBUTTON" ACCESSKEY="L" NAME="BtnClear">C<u>l</u>ear</button></td>
					<td CLASS="LABEL" align="left" width="5%">&nbsp;</td>
				</tr>
				<tr>
					<td CLASS="LABEL" VALIGN="BOTTOM" WIDTH="20%">Last Name:<br>
					<input CLASS="LABEL" TYPE="TEXT" NAME="LNAME" SIZE="38"></td>
					<td CLASS="LABEL" VALIGN="BOTTOM" width="20%">M.I.:<br>
					<input CLASS="LABEL" TYPE="TEXT" NAME="MI" SIZE="3"></td>
					<td CLASS="LABEL" colspan="3" VALIGN="BOTTOM" WIDTH="60%">First Name:<br>
					<input CLASS="LABEL" TYPE="TEXT" NAME="FNAME" SIZE="38"></td>
				</tr>
				<tr>
					<td CLASS="LABEL" VALIGN="BOTTOM" width="20%">City:<br>
					<input CLASS="LABEL" TYPE="TEXT" NAME="CITY" SIZE="38"></td>
					<td CLASS="LABEL" VALIGN="BOTTOM" width="20%">Zip Code:<br>
					<input CLASS="LABEL" TYPE="TEXT" NAME="ZIP" SIZE="10"></td>
					<td CLASS="LABEL" VALIGN="BOTTOM" width="10%">State:<br>
					<input CLASS="LABEL" TYPE="TEXT" NAME="STATE" SIZE="8"></td>
					<td CLASS="LABEL" VALIGN="BOTTOM" width="50%" colspan="2">Phone:<br>
					<input CLASS="LABEL" TYPE="TEXT" NAME="PHONE" SIZE="18"></td>
					<input TYPE=HIDDEN NAME="CLIENT_NODE" VALUE="<%=Request.QueryString("CLIENT_NODE")%>">
					<input TYPE=HIDDEN NAME="SDID">
					<input TYPE=HIDDEN NAME="MODE" VALUE="<%=Request.QueryString("MODE")%>">
					<%if not isempty(Request.QueryString("FULLSEARCH")) then%>
						<input TYPE=HIDDEN NAME="FULLSEARCH" VALUE="">
					<%end if%>
				</tr>
			</table>
			<table border="0">
				<tr>
					<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B"
					CLASS="LABEL">Begins With</td>
					<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C"
					CLASS="LABEL">Contains</td>
					<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E"
					CLASS="LABEL">Exact</td>
				</tr>
			</table>
		</td>
    </tr>
  </table>
</form>
</body>
</html>
