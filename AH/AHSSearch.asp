<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<%
Response.Expires = 0
Response.AddHeader  "Pragma", "no-cache"
'response.write(Request.QueryString)


%>
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>AHSID Search</title>
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
<!--#include file="..\lib\Help.asp"-->
Sub BtnClear_onclick()
	document.all.SearchAHSID.value = ""
	document.all.SearchName.value = ""
	document.all.SearchFNS_CLIENT_CD.value = ""
	document.all.SearchNATURE_OF_BUSINESS.value = ""

	document.all.Search_TYPE.value = ""
	document.all.SearchUPLOAD_KEY.value = ""
	<% if Request.QueryString("ORIGIN") <> "USERS" then %> 
	document.all.SearchLOCATION_CODE.value = ""
	<%END IF %>
	document.all.SearchSUID.value = ""

End Sub

Sub BtnSearch_onclick()
	'If document.all.SearchAHSID.value = "" And document.all.SearchName.value = "" And _
	'document.all.SearchFNS_CLIENT_CD.value = "" And document.all.SearchNATURE_OF_BUSINESS.value = "" Then
	'		MsgBox "Please enter search criteria", 0, "FNSNetDesigner"
	'Else
		document.all.SpanStatus.innerHTML = "<%=MSG_SEARCH%>"
		FrmSearch.submit
	'End If
End Sub

Sub window_onload
	'document.all.SearchName.focus ' Timing Problem
	document.all.SearchType(0).checked = True
	UpdateStatus("Ready")	
<% 'response.write(request.querystring) %>
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

	If document.all.SearchAHSID.value <> "" And document.all.SearchName.value <> "" And _
	document.all.SearchFNS_CLIENT_CD.value <> "" And document.all.SearchNATURE_OF_BUSINESS.value <> "" Then
		UpdateStatus("<%=MSG_PROMPT%>")	
	End If

End Sub

Sub PostTo(strURL)
	curAHSID = Parent.frames("WORKAREA").GetAHSID
	temp = Split(curAHSID, "||")
	If UBound(temp) >= 0 Then 
		document.all.AHSID.value = temp(0)
		document.all.CLIENTNODE.value = Parent.frames("WORKAREA").getCientNode
	Else		
		document.all.AHSID.value = ""
	End If
	FrmSearch.action = "AHSDetails-f.asp"
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

sub enable_exact()
   document.all.SearchType(2).checked  = true 
   document.all.SearchType(0).disabled  = true 
   document.all.SearchType(1).disabled  = true 
end sub

sub enable_begin()
   document.all.SearchType(0).checked  = true 
   document.all.SearchType(0).disabled  = false 
   document.all.SearchType(1).disabled  = false 
   
end sub
</script>


</head>

<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» AHS Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" OnClick="LaunchHelp(&quot;Welcome.htm&quot;)" WIDTH="7" HEIGHT="8"></td>
<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label">
<tr>
<td VALIGN="CENTER" WIDTH="5">
<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" VALIGN="CENTER" ALT="View Status Report">
</td>
<td width="485">
:<span VALIGN="CENTER" ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>

<form Name="FrmSearch" METHOD="GET" ACTION="AHSSearchResults.asp" TARGET="WORKAREA">
  <input type="hidden" name="MODE" value="<%=Request.QueryString("MODE")%>"><input type="hidden" name="AHSID" value="<%=Request.QueryString("AHSID")%>">
  <input type="hidden" name="CLIENTNODE" value="">
  <table width="100%" CELLPADDING="0" CELLSPACING="0">
    <tr>
      <td><table CLASS="LABEL" style="width:300" align="left">
        <tr>
          <td CLASS="LABEL" COLSPAN="2">Name:<br>
          <input CLASS="LABEL" tabindex="1" TYPE="TEXT" NAME="SearchName" size="30" VALUE="<%=Request.QueryString("SearchName")%>"></td>
          <td CLASS="LABEL">Nature of Business:<br>
          <input CLASS="LABEL" tabindex="2" TYPE="TEXT" NAME="SearchNATURE_OF_BUSINESS" size="30" VALUE="<%=Request.QueryString("SearchNATURE_OF_BUSINESS")%>" style='text-transform:uppercase'></td>
        </tr>
        <tr>
          <td CLASS="LABEL">A.H.S. ID:<br>
          <input size="12" tabindex="5" CLASS="LABEL" TYPE="TEXT" NAME="SearchAHSID" VALUE="<%=Request.QueryString("SearchAHSID")%>" onfocus="enable_exact()" onBlur="enable_begin()"  ></td>
          <td CLASS="LABEL" style='text-transform:uppercase'><nobr>FNS Client Code:<br>
          <input size="12" tabindex="6" CLASS="LABEL" TYPE="TEXT" NAME="SearchFNS_CLIENT_CD" VALUE="<%=Request.QueryString("SearchFNS_CLIENT_CD")%>" style='text-transform:uppercase'></nobr></td>
          <td CLASS="LABEL"><nobr>Upload key:<br>
          <input size="30" tabindex="7" CLASS="LABEL" TYPE="TEXT" NAME="SearchUPLOAD_KEY" VALUE="<%=Request.QueryString("SearchUPLOAD_KEY")%>" style='text-transform:uppercase'></nobr></td>
 	</tr>
	<tr>
          <td CLASS="LABEL"><nobr>SUID:<br>
          <input size="12" tabindex="9" CLASS="LABEL" TYPE="TEXT" NAME="SearchSUID" VALUE="<%=Request.QueryString("SearchSUID")%>" ></nobr></td>
          <% if Request.QueryString("ORIGIN") = "USERS" then %> 
          <td CLASS="LABEL"><nobr>Parent Node ID:<br>
          <input CLASS="LABEL" tabindex="8" TYPE="TEXT" NAME="SearchPARENT_NODE_ID" size="1" VALUE="1" readonly></td>
          <%else%>
          <td CLASS="LABEL"><nobr>Location Code:<br>
          <input size="12" tabindex="8" CLASS="LABEL" TYPE="TEXT" NAME="SearchLOCATION_CODE" VALUE="<%=Request.QueryString("SearchLOCATION_CODE")%>" style='text-transform:uppercase'></nobr></td>
          <%end if%>
          <td CLASS="LABEL">Type:<br>
          <input CLASS="LABEL" tabindex="9" TYPE="TEXT" NAME="Search_TYPE" size="30" VALUE="<%=Request.QueryString("Search_TYPE")%>" style='text-transform:uppercase'></td>
    </tr>
      </table>
      </td>
      <td VALIGN="TOP" rowspan="3"><table>
        <tr>
          <td CLASS="LABEL"><button CLASS="StdButton" tabindex="14" NAME="BtnSearch" TYPE="BUTTON" ACCESSKEY="H">Searc<u>h</u></button></td>
        </tr>
        <tr>
          <td CLASS="LABEL"><button CLASS="StdButton" tabindex="15" NAME="BtnClear" ACCESSKEY="L">C<u>l</u>ear</button></td>
        </tr>
      </table>
      </td>
    </tr>
    <tr>
      <td><table>
        <tr>
          <td CLASS="LABEL"><input TYPE="RADIO" tabindex="11" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL">Begins With</td>
          <td CLASS="LABEL"><input TYPE="RADIO" tabindex="12" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
          <td CLASS="LABEL"><input TYPE="RADIO" tabindex="13" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
         
        </tr>
      </table>
      </td>
    </tr>
  </table>
</form>
</body>
</html>
