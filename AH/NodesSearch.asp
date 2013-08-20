<!--#include file="..\lib\common.inc"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<TITLE></TITLE>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<STYLE>
BODY { background:#d6cfbd;
			Font-Family:Verdana;
			Font-Size=8;
		 }
</STYLE>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub ClearSearch()
	document.all.SearchName.value = ""
	document.all.SearchLocCode.value = ""
	document.all.SearchPolicy.value = ""
	document.all.SearchAddress.value = ""
	document.all.SearchCity.value = ""
	document.all.SearchState.value = ""
	document.all.SearchZip.value = ""
End Sub

Sub ExeSearch()

End Sub

Sub window_onload
	'document.all.SearchName.focus ' Timing Problem
End Sub

-->
</SCRIPT>
</HEAD>
<BODY  topmargin=0 leftmargin=0>
<FORM Name="FrmSearch" METHOD=POST ACTION="PolicySearchResults.asp" TARGET="WORKAREA">
<TABLE>
<TR>
<TD></TD>
<TD CLASS=LABEL>Company Name<BR><INPUT CLASS=LABEL TYPE=TEXT NAME="SearchName"></TD>
<TD CLASS=LABEL>Loc Code<BR><INPUT CLASS=LABEL TYPE=TEXT NAME="SearchLocCode"></TD>
<TD CLASS=LABEL COLSPAN=2>Policy #<BR><INPUT CLASS=LABEL TYPE=TEXT NAME="SearchPolicy"></TD>
</tr>
<tr>
<td colspan=1></td>
<TD CLASS=LABEL>Address<BR><INPUT CLASS=LABEL TYPE=TEXT NAME="SearchAddress"></TD>
<TD CLASS=LABEL>City<BR><INPUT CLASS=LABEL TYPE=TEXT NAME="SearchCity"></TD>
<TD CLASS=LABEL>State<BR><INPUT CLASS=LABEL size = 3 TYPE=TEXT NAME="SearchState"></TD>
<TD CLASS=LABEL>Zip<BR><INPUT CLASS=LABEL TYPE=TEXT NAME="SearchZip"></TD>
</TR>
</TABLE>
</FORM>
</BODY>
</HTML>
