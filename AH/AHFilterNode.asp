<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\SearchMsg.inc"-->
<!--#include file="..\lib\AHSTree.inc"-->
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=VBScript>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<TITLE></TITLE>

<script>
function AddNode(AHSID, NAME, inWhichList)
{
	var curList = null;
	
	if (inWhichList == "INCLUDE") curList = document.all.lstInclude;
	else if (inWhichList == "EXCLUDE") curList = document.all.lstExclude;
	
	if (curList != null)
	{
		if (CheckForNode(AHSID) == false)
		{
			var newItem = document.createElement("OPTION");
			newItem.text = NAME;
			newItem.value = AHSID;
			curList.add(newItem);
		}		
	}
}


function CheckForNode(AHSID)
{	
	for (i=0; i < document.all.lstExclude.length; i++)
	{	
		if (document.all.lstExclude(i).text == AHSID)
		{
			alert("Selected node is already present in the exclude list.");
			return true;
		}
	}
	for (i=0; i < document.all.lstInclude.length; i++)
	{	
		if (document.all.lstInclude(i).text == AHSID)
		{
			alert("Selected node is already present in the include list.");
			return true;
		}
	}
	return false;
}

function GetNodes(inDelim, inWhichList)
{
	var strItems = "", curList = null;
	
	if (inWhichList == "INCLUDE") curList = document.all.lstInclude;
	else if (inWhichList == "EXCLUDE") curList = document.all.lstExclude;
	
	if (curList != null)
	{
		for(i=0; i < curList.length; i++)
		{
			if (strItems != "") strItems = strItems + inDelim;
			strItems = strItems + curList(i).value;
		}
	}
	
	return strItems;
}

function GetUseWhereClause()
{
	if (document.all.USEWHERECLAUSE.checked == true)
		return "TRUE";
	else 
		return "FALSE"; 
}

function ExeClear()
{
	ClearNodes(document.all.lstInclude);
	ClearNodes(document.all.lstExclude);
	document.all.USEWHERECLAUSE.checked = true;	
	
}
function ClearNodes(curList)
{
	if (curList != null)
	{
		for (i=curList.length; i >= 0; i--)
			curList.remove(i);
	}
}
</script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=JScript>
function USEWHERECLAUSE_onclick()
{
	if ((GetNodes(",", "INCLUDE") == "") && (GetNodes(",", "INCLUDE") == ""))
	{
		alert("You must select at least one node to include before selecting this option.");
		document.all.USEWHERECLAUSE.checked = true;	
	}
	else
		parent.frames("TOP").ClearSearch();

}
</SCRIPT>
</HEAD>

<BODY BGCOLOR='<%= BODYBGCOLOR %>' rightmargin=0 bottommargin=0 leftmargin=0 topmargin=0>
<FORM Name="FrmSearch" TARGET="WORKAREA" METHOD=POST ACTION="AHFilterSearchResults.asp">

<table class="LABEL">
<%
    AHSID = CStr(Request.QueryString("AHSID"))
    strMustInclude = ""
    strMustExclude = ""
    
    strUseWhereClause  = "CHECKED"
	if AHSID <> "" then 
		If HasSpecificFilter("AHSID=" & AHSID , "DESIGNER_AHSFILTER") = true Then
			Set RS = Server.CreateObject("ADODB.RecordSet")
			strMustInclude = GetSpecificFilter("AHSID=" & AHSID,"DESIGNER_AHSFILTER", "MUSTINCLUDE")
			strMustExclude = GetSpecificFilter("AHSID=" & AHSID, "DESIGNER_AHSFILTER","MUSTEXCLUDE")
			strUseWhereClause = GetSpecificFilter("AHSID=" & AHSID, "DESIGNER_AHSFILTER","USEWHERECLAUSE")			
			
			if strUseWhereClause = "TRUE" then 
				strUseWhereClause = "CHECKED"
			else
				strUseWhereClause = ""
			end if
				
		End If
	End If
%>		
<tr><td><input type=checkbox name="USEWHERECLAUSE" LANGUAGE="JScript" onclick="return USEWHERECLAUSE_onclick();" <%=strUseWhereClause%>>Use filter fields?</select></td>
</tr>

<td>Must exclude:<br>

<SELECT class="LABEL" name="lstExclude" style="width:325" SIZE=4 >
<%
		
	If strMustExclude <> "" Then
		SQL = "SELECT ACCNT_HRCY_STEP_ID, NAME FROM ACCOUNT_HIERARCHY_STEP WHERE ACCNT_HRCY_STEP_ID IN( " & strMustExclude & ")"
		RS.Open SQL, CONNECT_STRING, adOpenStatic, adLockReadOnly, adCmdText
		If Not RS.EOF AND Not RS.BOF Then 
			Do While Not RS.EOF %>
			
		<OPTION VALUE="<%=RS("ACCNT_HRCY_STEP_ID")%>"><%=RS("NAME")%></OPTION>	 

<%			RS.MoveNext
			Loop
		End If
			
		RS.Close
	End If 
%>		
</SELECT></td>
<td>Must include:<br>
<SELECT class="LABEL" name="lstInclude" style="width:325" SIZE=4 >
<%		
	If strMustInclude <> "" Then
		SQL = "SELECT ACCNT_HRCY_STEP_ID, NAME FROM ACCOUNT_HIERARCHY_STEP WHERE ACCNT_HRCY_STEP_ID IN( " & strMustInclude & ")"
		RS.Open SQL, CONNECT_STRING, adOpenStatic, adLockReadOnly, adCmdText
		If Not RS.EOF AND Not RS.BOF Then 
			Do While Not RS.EOF
%>
		<OPTION VALUE="<%=RS("ACCNT_HRCY_STEP_ID")%>"><%=RS("NAME")%></OPTION>
<% 
			RS.MoveNext
			Loop
		End If
		RS.Close
	End If
		
	Set RS=Nothing
%>

</SELECT></td></table>

</FORM>
</BODY>
</HTML>
