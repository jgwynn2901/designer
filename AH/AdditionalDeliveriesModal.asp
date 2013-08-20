<!--#include file="..\lib\common.inc"-->
<%
	dim bShowSave, bShowClose

	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"

	if Request.QueryString("MODE") <> "" then
		MODE = Request.QueryString("MODE")
	end if
	
	bShowClose = true
	
	Select Case MODE
		Case "RO"
			bShowSave = false
		Case "RW"
			bShowSave = true 
	End Select		 
%>  
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<title>Additional Deliveries Maintenance</title>
<STYLE TYPE="text/css">
HTML {width: 280pt; height: 140pt}
</STYLE>

<SCRIPT LANGUAGE="JScript">
function CAdditionalDeliveriesObj()
{
	this.strAdditional = "";
}

var AdditionalDeliveriesObj = new CAdditionalDeliveriesObj();
</SCRIPT>

<script LANGUAGE="JavaScript" FOR="BtnAdd" EVENT="onclick">
	var oOption, nIndex;
	
	nIndex = document.FrmDetails.AVAILABLE_LIST.selectedIndex;
	if( (document.FrmDetails.AVAILABLE_LIST.length > 0) && (nIndex != -1) )
	{
		var newItem = document.createElement("OPTION");
		newItem.text = document.FrmDetails.AVAILABLE_LIST.item(nIndex).text;
		newItem.value = document.FrmDetails.AVAILABLE_LIST.item(nIndex).value;
			
		document.FrmDetails.INUSE_LIST.options.add(newItem);
		document.FrmDetails.AVAILABLE_LIST.remove(nIndex);		
		
		document.body.ScreenDirty = "YES";
	}
		
</script>

<script LANGUAGE="JavaScript" FOR="BtnRemove" EVENT="onclick">
	var oOption, nIndex;
	
	nIndex = document.FrmDetails.INUSE_LIST.selectedIndex;
	if( (document.FrmDetails.INUSE_LIST.length > 0) && (nIndex != -1) )
	{
		var newItem = document.createElement("OPTION");
		newItem.text = document.FrmDetails.INUSE_LIST.item(nIndex).text;
		newItem.value = document.FrmDetails.INUSE_LIST.item(nIndex).value;

		document.FrmDetails.AVAILABLE_LIST.options.add(newItem);
		document.FrmDetails.INUSE_LIST.remove(nIndex);		
		
		document.body.ScreenDirty = "YES";
	}	
</script>

<script LANGUAGE="JavaScript" FOR="BtnSave" EVENT="onclick">
	var strList, i, nCount, oTemp;
	
	strList = " ";
	nCount = document.FrmDetails.INUSE_LIST.length;
	
	//We want a space at the end for easier field searching.	
	for(i = 0; i < nCount; i++)
	{
		oTemp = document.FrmDetails.INUSE_LIST.item(i);
		strList = strList + oTemp.value + " ";
	}
	
	AdditionalDeliveriesObj.strAdditional = strList;
	
	window.close();
</script>

<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
	var s, strList, i, j;
	var oOption;
	AdditionalDeliveriesObj = window.dialogArguments;
	
	s=AdditionalDeliveriesObj.strAdditional;
	
	if(s.length > 0)
	{
		strList = s.split(" ");
	
		for(i = 0; i < strList.length; i++)
		{
	
			for(j=0; j < document.FrmDetails.AVAILABLE_LIST.length; j++)
			{
				oOption = document.FrmDetails.AVAILABLE_LIST.item(j);
				
				if( strList[i] == oOption.value)
				{
					document.FrmDetails.AVAILABLE_LIST.remove(j);
					document.FrmDetails.INUSE_LIST.options.add(oOption);
					break;
				}
			}
		}
	}
	
</script>

<script LANGUAGE="JavaScript" FOR="BtnClose" EVENT="onclick">
	var lRet;
	if ( document.body.ScreenDirty == "YES")
	{
		lRet = confirm("Are you sure you want to discard your changes?");
		if(lRet == true)
			window.close();
	}
	else
	{
		window.close();
	}
</script>

</head>
<BODY  topmargin=20 leftmargin=20 rightmargin=0 BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» Additional Deliveries</td>
<TD HEIGHT=5 ALIGN=LEFT>
<TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
<TR><TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=100 HEIGHT=4></TD></TR>
<TR><TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
<TD WIDTH=100 HEIGHT=8></TD></TR>
</TABLE></TD></TR>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1" WIDTH="70%"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>
<form Name="FrmDetails">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0" align="center">
	<TR>
		<TD><SELECT CLASS=LABEL NAME=AVAILABLE_LIST SIZE=5 STYLE="WIDTH:130;" ><!--#include file="..\lib\AdditionalDeliveryCodes.asp"--></SELECT></TD>
		<% If bShowSave = true then %>
		<TD><TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
			<TR><TD><BUTTON NAME="BtnAdd" STYLE="CURSOR:HAND;WIDTH:80" CLASS=STDBUTTON ACCESSKEY="C">Add >></BUTTON></TD></TR>
			<TR><TD><BUTTON NAME="BtnRemove" STYLE="CURSOR:HAND;WIDTH:80" CLASS=STDBUTTON ACCESSKEY="C"><< Remove</BUTTON></TD></TR></TABLE>
		</TD>
		<% End If %>
		<TD><SELECT CLASS=LABEL NAME=INUSE_LIST SIZE=5 STYLE="WIDTH:130;"></SELECT></TD>
	</TR>
</TABLE>
</form>
<TABLE align="left">
<TR>
<%	if bShowSave = true then %>
<TD CLASS=LABEL><BUTTON CLASS=StdButton NAME=BtnSave ACCESSKEY="S"><U>S</U>elect</BUTTON></TD>
<%	end if 
	if bShowClose = true then %>
<TD CLASS=LABEL><BUTTON CLASS=StdButton NAME=BtnClose >Close</BUTTON></TD>
<%	end if %>
</TR>
</TABLE>
</body>
</html>
