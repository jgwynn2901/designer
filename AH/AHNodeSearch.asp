<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE></TITLE>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub ClearSearch()
	document.all.NAME.value = ""
	document.all.CITY.value = ""
	document.all.STATE.value = ""
	document.all.ZIP.value = ""
	document.all.ADDRESS_1.value = ""
	document.all.LOCATION_CODE.value = ""
	'Added for PSUS-0008
	'Adding a new Search filter AHS ID under a client node.
	'Prashant Shekhar 04/16/2007
	document.all.ACCNT_HRCY_STEP_ID.value = ""
End Sub

 Sub ExeSearch()
		'Added for PSUS-0008
		'Adding a new Search filter AHS ID under a client node.
		'Prashant Shekhar 04/16/2007
	If document.all.NAME.value = "" AND document.all.CITY.value = "" AND document.all.STATE.value = "" AND document.all.ZIP.value = "" AND document.all.ADDRESS_1.value = "" AND document.all.LOCATION_CODE.value = "" and document.all.ACCNT_HRCY_STEP_ID.value = "" Then
		MsgBox "Please enter search criteria!", 0 , "FNSNetDesigner"
	Else
		SPANSTATUS.innerHTML = "<%= MSG_SEARCH %>"
		FrmSearch.submit
	End If
End Sub

Sub window_onload

End Sub

Sub BtnSearch_OnClick
	Call ExeSearch()
End Sub

Sub BtnClear_OnCLick
	Call ClearSearch()
End Sub
-->
</SCRIPT>
<!-- Javascript added for PSUS-0008 to enable only numbers to be entered
	 for the AHS ID text field. 
	Author : Prashant Shekhar 
	Date : 04/05/2007   -->
<script id = clientEventHandlersJS language = javascript>
function onlyNumbers()
 {
  if(event.keyCode < 48 || event.keyCode > 57) 
   event.returnValue = false;
  else if(event.which < 48 || event.which > 57) 
   return false;
 }
</script>
</HEAD>

<BODY  topmargin=0 leftmargin=0 bgcolor='<%= BODYBGCOLOR %>' bottommargin=0 rightmargin=0>
<FORM Name="FrmSearch" TARGET="WORKAREA" METHOD=POST ACTION="AHNodeSearchResults.asp">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Business Entity Search</TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME=AT_AHSID VALUE="<%= Request.QueryString("AHSID") %>">
<TABLE  cellspacing=0 cellpadding=0>
<TR>
<TD CLASS=LABEL><img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" Title=""></TD>
<TD CLASS=LABEL><SPAN ID=SPANSTATUS STYLE="COLOR:#006699" CLASS=LABEL>: Ready</SPAN></TD>
</TR>
</TABLE>
<TABLE WIDTH="100%"><TR><TD VALIGN=TOP ALIGN=LEFT>
<TABLE>
<TR>
<TD CLASS=LABEL>Name:<BR><INPUT TYPE=TEXT SIZE=45 NAME=NAME CLASS=LABEL></TD>
<TD CLASS=LABEL>Location Code:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=LOCATION_CODE MAXLENGTH=30 SIZE=20></TD>
<TD CLASS=LABEL colspan = 2>AHS ID:<BR><INPUT TYPE = TEXT CLASS=LABEL NAME = ACCNT_HRCY_STEP_ID onkeypress="javascript:onlyNumbers()" MAXLENGTH=10 SIZE=17></TD> 
</TR>
<TR>
<TD CLASS=LABEL>Address:<BR><INPUT TYPE=TEXT SIZE=45 NAME=ADDRESS_1 CLASS=LABEL></TD>
<TD CLASS=LABEL>City:<BR><INPUT TYPE=TEXT NAME=CITY SIZE=20 CLASS=LABEL></TD>
<TD CLASS=LABEL>State:<BR>
<SELECT NAME=STATE CLASS=LABEL>
<OPTION VALUE="">
<!--#include file="..\lib\states.asp"-->
</SELECT>
</TD>
<TD CLASS=LABEL>Zip:<BR><INPUT TYPE=TEXT NAME=ZIP SIZE=7 CLASS=LABEL></TD>
</TR>
</TABLE>
<TABLE CELLPADDING=0 CELLSPACING=0>
<tr>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL" CHECKED>Begins With</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
<td CLASS="LABEL"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
<td width=75>
<TD ALIGN=RIGHT CLASS=LABEL> Direction:
<input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchDirection" VALUE="UP" CLASS="LABEL">Up
<input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchDirection" VALUE="Down" CLASS="LABEL" CHECKED>Down
</td>
</tr>
</TABLE>
</TD><TD ALIGN=RIGHT VALIGN=TOP>
<TABLE>
<TR>
<td ALIGN=RIGHT CLASS="LABEL"><button CLASS="StdButton" NAME="BtnSearch" ACCESSKEY="C">Sear<u>c</u>h</button></td>
</TR>
<TR>
<td ALIGN=RIGHT CLASS="LABEL"><button CLASS="StdButton" NAME="BtnClear" ACCESSKEY="L">C<U>l</U>ear</button></td>
</TR>
</TABLE>
</TD></TR></TABLE>
</FORM>
</BODY>
</HTML>
