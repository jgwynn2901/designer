<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\ControlData.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Contact Search</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script>
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
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
Sub BtnClear_onclick()
	document.all.SearchName.value = ""
	document.all.SearchAHSID.value = ""
	document.all.SearchPhoneNumber.value = ""
	document.all.SearchDescription.value = ""
	document.all.SearchGreeting.value = ""
End Sub

Sub BtnSearch_onclick()
	'If document.all.SearchIBID.value = "" And document.all.SearchACCNT_HRCY_STEP_ID.value = "" And _
	'document.all.SearchLOB_CD.value = "" Then
	'		MsgBox "Please enter search criteria", 0, "FNSNetDesigner"
	'Else
		document.all.SpanStatus.innerHTML = "<%=MSG_SEARCH%>"
		FrmSearch.submit
	'End If
End Sub

Sub window_onload
	document.all.SearchType(0).checked = True
	UpdateStatus("Ready")	
	
<%	If Request.QueryString <> "" Then %>
<%		If CStr(Request.QueryString("SearchType")) = "B" Then	%>
			document.all.SearchType(0).checked = True
<%		ElseIf CStr(Request.QueryString("SearchType")) = "C" Then	%>
			document.all.SearchType(1).checked = True
<%		ElseIf CStr(Request.QueryString("SearchType")) = "E" Then	%>
			document.all.SearchType(2).checked = True
<%		End If

		If CStr(Request.QueryString("SearchInputType")) <> "" Then	%>
			SelectOption document.all.SearchInputType,"<%=CStr(Request.QueryString("SearchInputType"))%>"
<%		End If 

	End If %>
	
'Search Controls		
<% If Request.QueryString("SearchAHSID") <> "" Then %>
	document.all.SearchAHSID.value = "<%= Request.QueryString("SearchAHSID") %>"
<% End If %>
<% If Request.QueryString("SearchName") <> "" Then %>
	document.all.SearchName.value = "<%= Request.QueryString("SearchName") %>"
<% End If %>
<% If Request.QueryString("SearchPhoneNumber") <> "" Then %>
	document.all.SearchPhoneNumber.value = "<%= Request.QueryString("SearchPhoneNumber") %>"
<% End If %>
<% If Request.QueryString("SearchDescriptiont") <> "" Then %>
	document.all.SearchDescriptiont.value = "<%= Request.QueryString("SearchDescriptiont") %>"
<% End If %>
<% If Request.QueryString("SearchGreeting") <> "" Then %>
	document.all.SearchGreeting.value = "<%= Request.QueryString("SearchGreeting") %>"
<% End If %>

If document.all.SearchAHSID <> "" And document.all.SearchName.value <> "" And _
   document.all.SearchPhoneNumber.value <> "" and document.all.SearchDescriptiont.value <> ""  and  document.all.SearchGreeting.value <> "" Then
     UpdateStatus("<%=MSG_PROMPT%>")	
End If

End Sub

Sub PostTo(strURL)

	curIBID = Parent.frames("WORKAREA").GetIBID
	temp = Split(curIBID, "||")
	If UBound(temp) >= 0 Then 
		document.all.IBID.value = temp(0)
	Else		
		document.all.IBID.value = ""
	End If
	FrmSearch.action = "GreetingDetails-f.asp"
	FrmSearch.method = "GET"	
	FrmSearch.target = "_parent"	
	FrmSearch.submit
End Sub

Sub UpdateStatus(inStatus)
	document.all.SpanStatus.innerHTML = inStatus
End Sub
Sub StatusRpt_OnClick
	MsgBox "No other detail status reported.",0,"FNSNetDesigner"		
End Sub
<!--#include file="..\lib\Help.asp"-->
</script>
</head>
<BODY  topmargin=0 leftmargin=0  rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Greeting Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>
<table style="{position:absolute;top:20;}" class="Label" >
<tr>
<td VALIGN="CENTER" WIDTH="5" >
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER"  ALT="View Status Report">
</td>
<td width="485">
:<SPAN VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS=LABEL>Ready</SPAN>
</td>
</tr>
</table>
<form Name="FrmSearch" METHOD="GET" ACTION="GreetingSearchResults.asp" TARGET="WORKAREA">
 <input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" >
 <input type="hidden" NAME="IBID" value="<%=Request.QueryString("IBID")%>">
 <table width=100% CELLPADDING="0" CELLSPACING="0">
 <tr>
 <td><table>
	<tr></tr>
	<tr></tr>
	<tr nowrap>
	<td CLASS="LABEL">A.H. Step ID:<br>
	<input CLASS="LABEL" 
	       size="10"
	       MAXLENGTH=10 
	       tabindex=1 
	       TYPE="TEXT" 
	       NAME="SearchAHSID" 
	       VALUE="<%=Request.QueryString("SearchAHSID")%>"></td>
	<td CLASS="LABEL">Name:<br>
	<input CLASS="LABEL"
	       size="25" 
	       MAXLENGTH=80 
	       tabindex=2 
	       TYPE="TEXT" 
	       NAME="SearchName" 
	       VALUE="<%=Request.QueryString("SearchName")%>"></td>
	<td CLASS="LABEL">Phone Number<br>
	<input CLASS="LABEL" 
	        size="27" 
	        MAXLENGTH=40 
	        tabindex=3 TYPE="TEXT" 
	        NAME="SearchPhoneNumber" 
	        VALUE="<%=Request.QueryString("SearchPhoneNumber")%>"></td>
	</tr>
					
	<tr>
	<td CLASS="LABEL" colspan=3>Description:<br>
	<input size="73"
	       MAXLENGTH=225  
	       CLASS="LABEL" 
	       tabindex=4
	       TYPE="TEXT"
	       NAME="SearchDescriptiont"
	       VALUE="<%=Request.QueryString("SearchDescriptiont")%>"></td>
	
	<tr>
	<td CLASS="LABEL"  colspan=3>Greeting:<br>
    <input CLASS="LABEL"
           tabindex=5  
           size="73"
           MAXLENGTH=2000 
           TYPE="TEXT"
           NAME="SearchGreeting" 
           VALUE="<%=Request.QueryString("SearchGreeting")%>"></td>
    </tr> 
 </table>
</td>

<td VALIGN=TOP rowspan="3" >
	<TABLE>
	<TR><TD CLASS=LABEL><BUTTON CLASS=StdButton tabindex=9 NAME=BtnSearch TYPE="BUTTON" ACCESSKEY="H">Searc<U>h</U></BUTTON></TD></TR>
	<TR><TD CLASS=LABEL><BUTTON CLASS=StdButton tabindex=10 NAME=BtnClear ACCESSKEY="L">C<U>l</U>ear</BUTTON></TD></TR>
	</TABLE>
</td>	
</tr>
<tr>
<td>
	<table>
	<tr>
	<td CLASS="LABEL">
	 <input TYPE="RADIO"
	        tabindex=6 
	        STYLE="CURSOR:HAND" 
	        NAME="SearchType"
	        VALUE="B" 
	        CLASS="LABEL">Begins With</td>
	<td CLASS="LABEL">
	<input TYPE="RADIO" 
	       tabindex=7 
	       STYLE="CURSOR:HAND" 
	       NAME="SearchType"
	       VALUE="C" 
	       CLASS="LABEL">Contains</td>
	<td CLASS="LABEL">
	<input TYPE="RADIO" 
	       tabindex=8
	       STYLE="CURSOR:HAND"
	       NAME="SearchType"
	       VALUE="E" 
	       CLASS="LABEL">Exact</td>
	</tr>
	</table>
 </td>
 </tr>
 </table>
</form>
</body>
</html>
