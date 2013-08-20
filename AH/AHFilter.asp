<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\AHSTree.inc"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE></TITLE>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=jscript>

/*
function showAll_onclick()
{
var ALLNOTUSED = "0";
var ALLCHECKED = "1";
var ALLUNCHECKED = "2";

if (document.all.ShowAll.checked == true)
	{
	clearFields();
	document.all.SHOWALLStatus.value = ALLCHECKED;
	FrmSearch.submit();
	}
else
	{
	document.all.SHOWALLStatus.value = ALLUNCHECKED;
	FrmSearch.submit();
	}
}
*/
</SCRIPT>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
Sub ClearSearch()
'	const ALLUNCHECKED = "2"
'	if document.all.ShowAll.checked then
'		document.all.ShowAll.checked = false
'		document.all.SHOWALLStatus.value = ALLUNCHECKED
'		FrmSearch.submit
'	end if
clearFields
End Sub

sub clearFields
	document.all.NAME.value = ""
	document.all.CITY.value = ""
	document.all.STATE.value = ""
	document.all.ZIP.value = ""
	document.all.ADDRESS_1.value = ""
	document.all.FNS_CLIENT_CD.value = ""
	'Added for PSUS-0008
	'Adding a new Search filter AHS ID and Location Code under a client node.
	'Prashant Shekhar 04/17/2007
	document.all.LOCATION_CODE.value = ""
	document.all.ACCNT_HRCY_STEP_ID.value = ""
 end sub

Sub ExeSearch()
	If document.all.NAME.value = "" AND document.all.CITY.value = "" AND document.all.STATE.value = "" AND	document.all.ZIP.value = "" AND document.all.ADDRESS_1.value = "" AND document.all.FNS_CLIENT_CD.value = "" AND 	document.all.LOCATION_CODE.value = "" AND  document.all.ACCNT_HRCY_STEP_ID.value = ""	   Then
	MsgBox "Please enter search criteria!", 0 , "FNSNetDesigner"
	Else
		SPANSTATUS.innerHTML = "<%= MSG_SEARCH %>"
		FrmSearch.submit
	End If
End Sub

Sub BtnSearch_OnClick
'	const ALLNOTUSED = "0"

'	document.all.SHOWALLStatus.value = ALLNOTUSED
	Call ExeSearch()
End Sub

Sub BtnClear_OnCLick
	clearFields
End Sub

function GetWhereClause()
	Dim SQLWHERE, SearchType
	SQLWHERE = ""
	SearchType = "B"

	If document.all.SearchType(0).checked = true Then SearchType = "B"
	If document.all.SearchType(1).checked = true Then SearchType = "C"
	If document.all.SearchType(2).checked = true Then SearchType = "E"

	select Case(SearchType)
		case "B"
			NAME = Replace(document.all.NAME.value, "'", "''") & "%"
			FNS_CLIENT_CD = Replace(document.all.FNS_CLIENT_CD.value, "'", "''") & "%"
			CITY = Replace(document.all.CITY.value, "'", "''") & "%"
			STATE = Replace(document.all.STATE.value, "'", "''") & "%"
			ZIP = Replace(document.all.ZIP.value, "'", "''") & "%"
			ADDRESS_1 = Replace(document.all.ADDRESS_1.value, "'", "''") & "%"
			'Added for PSUS-0008
			'Adding a new Search filter AHS ID and Location Code under a client node.
			'Prashant Shekhar 04/17/2007
			LOCATION_CODE = Replace(document.all.LOCATION_CODE.value, "'", "''") & "%"
			ACCNT_HRCY_STEP_ID	= 	 Replace(document.all.ACCNT_HRCY_STEP_ID.value, "'", "''") & "%"
		case "C"
			NAME = "%" & Replace(document.all.NAME.value, "'", "''") & "%"
			FNS_CLIENT_CD = "%" & Replace(document.all.FNS_CLIENT_CD.value, "'", "''") & "%"
			CITY = "%" & Replace(document.all.CITY.value, "'", "''") & "%"
			STATE = "%" & Replace(document.all.STATE.value, "'", "''") & "%"
			ZIP = "%" & Replace(document.all.ZIP.value, "'", "''") & "%"
			ADDRESS_1 = "%" & Replace(document.all.ADDRESS_1.value, "'", "''") & "%"
			'Added for PSUS-0008
			'Adding a new Search filter AHS ID and Location Code under a client node.
			'Prashant Shekhar 04/17/2007

			LOCATION_CODE = "%" & Replace(document.all.LOCATION_CODE.value, "'", "''") & "%"
			ACCNT_HRCY_STEP_ID = "%" & Replace(document.all.ACCNT_HRCY_STEP_ID.value, "'", "''") & "%"
		case "E"
			NAME = Replace(document.all.NAME.value, "'", "''")
			FNS_CLIENT_CD = Replace(document.all.FNS_CLIENT_CD.value, "'", "''")
			CITY = Replace(document.all.CITY.value, "'", "''")
			STATE = Replace(document.all.STATE.value, "'", "''")
			ZIP = Replace(document.all.ZIP.value, "'", "''")
			ADDRESS_1 = Replace(document.all.ADDRESS_1.value, "'", "''")
			'Added for PSUS-0008
			'Adding a new Search filter AHS ID and Location Code under a client node.
			'Prashant Shekhar 04/17/2007
			LOCATION_CODE = Replace(document.all.LOCATION_CODE.value, "'", "''")
			ACCNT_HRCY_STEP_ID = Replace(document.all.ACCNT_HRCY_STEP_ID.value, "'", "''")
	End select
	
	If document.all.NAME.value <> "" Then 
		SQLWHERE = SQLWHERE & "Upper(NAME) LIKE '" & Ucase(NAME) & "'"
	End If
	
	If document.all.FNS_CLIENT_CD.value <> "" Then 
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " #AND# "
		End If
		SQLWHERE = SQLWHERE & "Upper(FNS_CLIENT_CD) LIKE '" & Ucase(FNS_CLIENT_CD) & "'"
	End If
	'Added for PSUS-0008
	'Adding a new Search filter AHS ID and Location Code under a client node.
	'Prashant Shekhar 04/17/2007

	 If document.all.LOCATION_CODE.value <> "" Then 
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " #AND# "
		End If
		SQLWHERE = SQLWHERE & "Upper(LOCATION_CODE) LIKE '" & Ucase(LOCATION_CODE) & "'"
	End If
	If document.all.ACCNT_HRCY_STEP_ID.value <> "" Then 
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " #AND# "
		End If
		SQLWHERE = SQLWHERE & "Upper(ACCNT_HRCY_STEP_ID) LIKE '" & Ucase(ACCNT_HRCY_STEP_ID) & "'"
	End If
	
	If document.all.CITY.value <> "" Then 
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " #AND# "
		End If
		SQLWHERE = SQLWHERE & "Upper(CITY) LIKE '" & Ucase(CITY) & "'"
	End If
	
	If document.all.STATE.value <> "" Then 
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " #AND# "
		End If
		SQLWHERE = SQLWHERE & "STATE LIKE '" & STATE & "'"
	End If
	
	If document.all.ZIP.value <> "" Then 
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " #AND# "
		End If
		SQLWHERE = SQLWHERE & "ZIP LIKE '" & ZIP & "'"
	End If

	If document.all.ADDRESS_1.value <> "" Then 
		If SQLWHERE <> "" Then
			SQLWHERE = SQLWHERE & " #AND# "
		End If
		SQLWHERE = SQLWHERE & "Upper(ADDRESS_1) LIKE '" & Ucase(ADDRESS_1) & "'"
	End If

'	if document.all.ShowAll.checked then
'		SQLWHERE = "NAME LIKE '%'"
'	end if
	GetWhereClause = SQLWHERE
end function

sub window_onload
<%
'const ALLNOTUSED = "0"
'const ALLCHECKED = "1"
'const ALLUNCHECKED = "2"

AHSID = CStr(Request.QueryString("AHSID"))
strWhereClause = ""
   
if AHSID <> "" then 
	If HasSpecificFilter("AHSID=" & AHSID , "DESIGNER_AHSFILTER") = true Then
		strWhereClause = GetSpecificFilter("AHSID=" & AHSID,"DESIGNER_AHSFILTER", "WHERECLAUSE")
	End If
End If
'if Session("AHSTreeShowAllNodes").Item("AHSID=" & AHSID) <> "" then
'	%>
'	document.all.ShowAll.checked = true
'	document.all.SHOWALLStatus.value = ALLCHECKED
'	FrmSearch.submit
'	<%
'else
	if strWhereClause <> "" then
	   'Get back the "+" sign in the Name Field of the 
	   'Filter Criteria iff Presents
	   'ILOG Issue : JMCA-0128
	   'Dated: 22/02/2007
	   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	   if (instr(strWhereClause,"[[--]]")>0) Then 
			  strWhereClause=Replace(strWhereClause,"[[--]]","+")
	   end if 
	   '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		parts = Split(strWhereClause,"#AND#")
		for i = 0 to UBound(parts)
			fields = Split(parts(i)," LIKE ")
			strField  = Replace(fields(0),"Upper(","")
			strField  = Replace(strField,")","")
			strValue = Replace(fields(1),"Ucase(","")
			strValue = Replace(strValue,")","")
			strValue = Replace(strValue,"'","")

			if Right(strValue,1) = "%" Then
				if Left(strValue,1) = "%" Then %>
	document.all.SearchType(1).checked = true
<%				else %>				
	document.all.SearchType(0).checked = true
<%				end if	
			else %>				
	document.all.SearchType(2).checked = true
<%			end if			

			strValue = Replace(strValue,"%","")			
%>
	document.all.<%=strField%>.value = "<%=trim(strValue)%>"

<%			
		next
%>		
		ExeSearch		
<%	end if
'end if
%>		

end sub

-->
</SCRIPT>
</HEAD>
<BODY  topmargin=0 leftmargin=0 bgcolor='<%= BODYBGCOLOR %>' bottommargin=0 rightmargin=0>
<FORM Name="FrmSearch" TARGET="WORKAREA" METHOD=POST ACTION="AHFilterSearchResults.asp">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
<TR><TD colspan=2 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 Business Entity Filter</TD>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=300 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
<TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<input type="hidden" name="AHSID" value="<%=Request.QueryString("AHSID")%>">
<input type="hidden" name="SHOWALLStatus">
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
<TD CLASS=LABEL>Client Code:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=FNS_CLIENT_CD MAXLENGTH=10 SIZE=20></TD>
 <!-- 'Added for PSUS-0008
	  'Adding a new Search filter AHS ID and Location Code under a client node.
	  'Prashant Shekhar 04/17/2007	-->

 <TD CLASS=LABEL>Location Code:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=LOCATION_CODE MAXLENGTH=10 SIZE=20></TD>
 <TD CLASS=LABEL>AHS ID:<BR><INPUT TYPE=TEXT CLASS=LABEL NAME=ACCNT_HRCY_STEP_ID MAXLENGTH=10 SIZE=10></TD>
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
<TABLE CELLPADDING=0 CELLSPACING=0 width="80%">
<tr>
<td CLASS="LABEL" width="22%"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="B" CLASS="LABEL" CHECKED>Begins With</td>
<td CLASS="LABEL" width="18%"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="C" CLASS="LABEL">Contains</td>
<td CLASS="LABEL" width="18%"><input TYPE="RADIO" STYLE="CURSOR:HAND" NAME="SearchType" VALUE="E" CLASS="LABEL">Exact</td>
<td CLASS="LABEL" width="17%"></td>
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
