<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Filter</title>
<style TYPE="text/css">
HTML {width: 600pt; height:500pt}
</style>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
function BtnClear_onclick() 
{
	if (document.frames("TabFrame").document.readyState == "complete")
	{
		document.frames("TabFrame").document.frames("TOP").document.all.SPANSTATUS.innerHTML = "<%= MSG_FILTER_OFF %>"
		URL = "AHFilterModify.asp?AHSID=<%=Request.QueryString("KEY")%>&ACTION=REMOVE";
		document.frames("TabFrame").document.frames("hiddenPage").document.location = URL;
		document.frames("TabFrame").ExeClear();
	}
}


function BtnClose_onclick() 
{
	if (document.frames("TabFrame").document.readyState == "complete")
		window.close()
}
</SCRIPT>
<SCRIPT SRC="..\Lib\ValidateSearchString.js"></SCRIPT>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE="VBScript">
sub BtnApply_OnClick
	If document.frames("TabFrame").document.readyState = "complete" Then
		document.frames("TabFrame").document.frames("TOP").document.all.SPANSTATUS.innerHTML = "<%= MSG_FILTER_ON %>"
		URL = "AHFilterModify.asp?AHSID=<%=Request.QueryString("KEY")%>&ACTION=ADD"

		WHERECLAUSE = document.frames("TabFrame").document.frames("TOP").GetWhereClause
		USEWHERECLAUSE = TabFrame.document.frames("NODEFILTER").GetUseWhereClause()
		MUSTINCLUDE = TabFrame.document.frames("NODEFILTER").GetNodes(",","INCLUDE")
		MUSTEXCLUDE = TabFrame.document.frames("NODEFILTER").GetNodes(",","EXCLUDE")

		ACTION = "ADD"
		
		WHERECLAUSE = Replace(WHERECLAUSE, "%", "|")
		WHERECLAUSE = f_EncodeURLString(WHERECLAUSE)
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		''+' is a Special Character in URL to Explorer itself so replace it with '[[--]]'
		''ILOG ISSUE : JMCA-0128
		''Dated : 22/02/2007
		if (instr(WHERECLAUSE,"+") > 0) then
			WHERECLAUSE = Replace(WHERECLAUSE,"+","[[--]]")
		end if 
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		
		URL = URL & "&USEWHERECLAUSE=" & USEWHERECLAUSE
		URL = URL & "&WHERECLAUSE=" & WHERECLAUSE
		URL = URL & "&MUSTINCLUDE=" & MUSTINCLUDE 
		URL = URL & "&MUSTEXCLUDE=" & MUSTEXCLUDE
		URL = URL & "&NODEDELIM=," 
		document.frames("TabFrame").document.frames("hiddenPage").document.location = URL
	End If
end sub

sub changeBtnStatus(lDisable)
	document.all.BtnApply.disabled = lDisable
end sub
</SCRIPT>


</head>
<body LEFTMARGIN="0" TOPMARGIN="0" bgcolor='<%= BODYBGCOLOR %>' bottommargin=0 rightmargin=0>
<!--OBJECT VIEWASTEXT data="..\Scriptlets\TabScriptlet.htm" id=TabsControl style="LEFT: 0px; TOP: 0px" type=text/x-scriptlet></OBJECT>
<iframe FRAMEBORDER="0" ID="TabFrame" WIDTH="1" HEIGHT="1"></iframe-->
<iframe FRAMEBORDER="0" ID="TabFrame" WIDTH="100%" HEIGHT="90%" SRC="AHFilter-f.asp?AHSID=<%=Request.QueryString("KEY")%>"></iframe>


<BR>
<TABLE>
<TR>
<td CLASS="LABEL"><button CLASS="StdButton" style="width:100" ID="BtnApply" NAME="BtnApply" ACCESSKEY="A" ><u>A</u>pply Filter</button></td>
<td CLASS="LABEL"><button CLASS="StdButton" style="width:100" NAME="BtnClear" ACCESSKEY="F" LANGUAGE=javascript onclick="return BtnClear_onclick()">Clear <U>F</U>ilter</button></td>
<td CLASS="LABEL"><button CLASS="StdButton" style="width:100" NAME="BtnClose" ACCESSKEY="O"  LANGUAGE=javascript onclick="return BtnClose_onclick()">Cl<U>o</U>se</button></td>
</TR>
</table>
</body>
</html>