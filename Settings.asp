<%	Response.Expires = 0 %>
<!--#include file="lib\common.inc"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="FNSDESIGN.css">
<TITLE></TITLE>
<script>
var g_StatusInfoAvailable = false;

function HasSelection()
{
	if ((document.all.SAVEFILTER.checked == true) || (document.all.SAVEFAVORITES.checked == true) || (document.all.SAVEMAXRECORDS.value != "") || (document.all.TREELEVELS.value != "") || (document.all.TREECOUNT.value != ""))
		return true;
	else
		return false;
}
function ValidNumber(CNumber)
{
	if (isNaN(CNumber) == false)
	{
		if (CNumber < 1 )
		{
			return false;
		}
		else
			return true;
	}
	else
		return false;
}

</script>
<script language="VBScript">
Sub UpdateStatus(inStatus)
	document.all.SpanStatus.innerHTML = inStatus
End Sub

Sub SetStatusInfoAvailableFlag(bAvailable)
	g_StatusInfoAvailable = bAvailable
	If bAvailable = true Then 
		document.all.StatusRpt.style.cursor = "HAND"
	Else
		document.all.StatusRpt.style.cursor = "DEFAULT"
	End If
End Sub

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If		
End Sub

</script>
</HEAD>
<BODY  topmargin=0 leftmargin=0 bgcolor='<%= BODYBGCOLOR %>' bottommargin=0 rightmargin=0>
<FORM Name="FrmSettings" TARGET="hiddenPage" METHOD=POST ACTION="SettingsSave.asp">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 FNSNet Designer System Settings</TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
<input type="hidden" name="ACTION">
<TABLE  cellspacing=0 cellpadding=0 >
<TR>
<TD CLASS=LABEL><img ID="StatusRpt" SRC="images\StatusRpt.gif" width="16" height="16" Title=""></TD>
<TD CLASS=LABEL><SPAN ID=SPANSTATUS STYLE="COLOR:#006699" CLASS=LABEL>: Ready</SPAN></TD>
</TR>
</TABLE>
<TABLE  CLASS=LABEL WIDTH="100%"  cellspacing=0 cellpadding=0 BORDER=0>
<TR>
<TD><input type=checkbox name="SAVEFILTER" LANGUAGE="JScript">Filters</select></td>
<TD VALIGN=BOTTOM CLASS=LABEL><INPUT CLASS=LABEL TYPE=TEXT NAME=TREELEVELS SIZE=3 VALUE="<%= Session("USERTREELEVELS") %>" MAXLENGTH=3> - # of AH Tree levels shown</TD>
</TR>
<TR>
<TD><input type=checkbox name="SAVEFAVORITES" LANGUAGE="JScript">Favorites</select></td>
<TD VALIGN=BOTTOM CLASS=LABEL><INPUT CLASS=LABEL TYPE=TEXT NAME=TREECOUNT SIZE=3 VALUE="<%= Session("USERTREECOUNT") %>" MAXLENGTH=3> - Max # of AH Tree Child Nodes Shown</TD>
</TR><TR>
<TD></TD>
<TD VALIGN=BOTTOM CLASS=LABEL><INPUT CLASS=LABEL TYPE=TEXT NAME=SAVEMAXRECORDS SIZE=3 VALUE="<%= Session("USERMAXRECORDS") %>" MAXLENGTH=3> - Search Result Max Records</TD>
</tr>
</TABLE>
<table width=100%>
<tr>
<td class=label width=25% valign=bottom>
<strong>Visible Area for Frame</strong>
</td>
<td class=label width=8% valign=bottom>Width:</td>
<td>
<input class=label type=text size=4 name=LCWidth value=<%=Session("LayoutCtlWidth")%>>
</td>
</tr>
<tr>
<td></td>
<td class=label width=8% valign=bottom>Height:</td>
<td>
<input class=label type=text size=4 name=LCHeight value=<%=Session("LayoutCtlHeight")%>>
</td>
</tr>
</table>
</FORM>
</BODY>
</HTML>
