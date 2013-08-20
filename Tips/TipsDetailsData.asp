<%	Response.Expires = 0
	Response.Buffer = true
	On Error Resume Next

	BranchTextLen = 25
	RuleTextLen = 25
	
%>
<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Account Tips List</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
<script language="VBScript">
dim cTips
function f_LastBARuleRecord
	if document.all.tblFields.Rows.Length <= 2 Then
		f_LastBARuleRecord = true
	else
		f_LastBARuleRecord = false
	end if
end Function

Function GetSelectedATLID
	dim idx	
	idx = CInt(getselectedindex(document.all.tblFields))
	If idx <> -1 Then
		GetSelectedATLID = document.all.tblFields.rows(idx).getAttribute("ATLID")
	Else
		GetSelectedATLID = ""
	End If
End Function
</script>
<!--#include file="..\lib\tablecommon.inc"-->
</head>
<BODY BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" leftmargin="0" rightmargin="0" >
<input type="hidden" name="TipsCount" id="TipsCount" >
<div align="LEFT" style="height:'100%';width:'100%'">
<table cellPadding="2" rules=all  cellSpacing="0" scrolling="auto" ID="tblFields" name="tblFields" width="100%">
	<thead CLASS="ResultHeader">
		<tr align="left">
		    <td class="thd"><div id><nobr>Tip Sequence</div></td>
			<td class="thd"><div id><nobr>Tip Description</div></td>
			<td class="thd"><div id><nobr>Enabled</div></td>			
		</tr>
	</thead>
	<tbody ID="TableRows">
	
<%	
	cTips=0
	ATID = CStr(Request.QueryString("ATID"))
	If ATID <> "NEW" And ATID <> "" Then
		Set oConn = Server.CreateObject("ADODB.Connection")
		oConn.Open CONNECT_STRING
		SQLST = "SELECT * FROM Account_Tip_List "
		SQLST = SQLST & "WHERE ACCOUNT_TIP_ID = "& ATID 
		SQLST = SQLST & " ORDER BY Tip_Sequence ASC" 
		Set oRS = oConn.Execute(SQLST)
		Do While Not oRS.EOF
%>
    <tr ID="FieldRow" CLASS="ResultRow"  DYNKEY="1" OnClick="Javascript:multiselect(this);" ATLID="<%=oRS("Account_Tip_List_ID")%>">
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(oRS("Tip_Sequence"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(oRS("Tip_Description"))%></td>
	<td NOWRAP CLASS="ResultCell" ><%=renderCell(oRS("Enabled_Flg"))%></td>
	</tr>
	<%
		cTips=cTips+1
			oRS.MoveNext
		Loop
		oRS.Close
		Set oRS = Nothing
		oConn.Close
		Set oConn = Nothing
	End If
	Response.write "<span STYLE='COLOR:#006699' class='label'> Tips Count is "& cTips  & "<br>"	
	set document.getElementById("TipsCount").value = cTips
%>
</tbody>
</table>
<script type="text/javascript" language="javascript">	
	window.parent.FrmDetails.TIPSCOUNT.value = "<%=cTips%>";
	//document.parent.frame.cTips = "<%=cTips%>";
	//alert(document.getElementById("TipsCount").value);
	
</script>
</div>
</BODY>
</HTML>