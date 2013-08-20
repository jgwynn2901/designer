<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\Security.inc"-->
<!--#include file="..\lib\ValidValues.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<% Response.Expires=0%>

<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<link rel="stylesheet" type="text/css" href="../FNSDESIGN.css">
<title>Routing Plan</title>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--
<!--#include file="..\lib\Help.asp"-->
Sub BtnGrfxBack_Onclick()
	self.location.href = "../AH/NodeSummary.asp?AHSID=<%= Request.QueryString("AHSID") %>"
End Sub

Sub Window_Onload
	document.all.ACCNT_HRCY_STEP_ID.value = "<%= Request.QueryString("AHSID") %>"
	document.all.DESCRIPTION.value = "WOR " & "<%= Request.QueryString("CLIENT") %>" & " Fax to Branch"
	for i = 0 to document.all.Frmsave.state.length -1
        document.all.Frmsave.state.options(i).selected = "False"
     next
     document.all.Frmsave.state.options(0).selected = "True"
     document.all.SEQUENCE2.disabled = true
	 document.all.OUTPUTDEF_ID2.disabled = true
	 
	 'Since the disable feature is not valid in browsers below 5.5, implement that by hiding the image
	 'document.all.BtnFindOD2.disabled = true
	 'document.all.BtnFindOD3.disabled = true
	 
	 document.all.BtnFindOD2.width = 0
	 document.all.SEQUENCE3.disabled = true
	 document.all.OUTPUTDEF_ID3.disabled = true
	 document.all.BtnFindOD3.width = 0
	 document.all.BtnDetachOD3.width = 0
	 document.all.BtnDetachOD2.width = 0
	 document.all.BtnDetachOD1.width = 0
End Sub

Function AttachRule (ID, SPANID)

RID = ID.value
RuleSearchObj.RID = RID
RuleSearchObj.RIDText = SPANID.innerhtml
RuleSearchObj.Selected = false
MODE = document.body.getAttribute("ScreenMode")

   If RID = "" Then RID = "NEW"
   If RID = "NEW" And MODE = "RO" Then
		MsgBox "No rule currently attached.",0,"FNSNetDesigner"
		Exit Function
   End If
   strURL = "..\Rules\RuleMaintenance.asp?SECURITYPRIV=FNSD_CALLFLOW&CONTAINERTYPE=MODAL&RID=" & RID
   showModalDialog  strURL  ,RuleSearchObj ,"dialogWidth:450px;dialogHeight:450px;center"
	
   If RuleSearchObj.Selected = true Then
      If RuleSearchObj.RID <> ID.value then
		ID.value = RuleSearchObj.RID
	  end if
	  SPANID.innerhtml = RuleSearchObj.RIDText
	  SPANID.Title = RuleSearchObj.RIDText
   ElseIf ID.value = RuleSearchObj.RID And RuleSearchObj.RID<> "" Then
	  SPANID.innerhtml = RuleSearchObj.RIDText
	  SPANID.Title = RuleSearchObj.RIDText
   End If
End Function

Function DetachRule(ID, SPANID)
<% If MODE="RO" Then Response.Write(" Exit Function ") %>
	ID.value = ""
	SPANID.innerhtml = ""
	SPANID.Title = ""
End Function

Function AttachNode(ID)
    msgbox "WARNING: Changing the A.H.S. ID will attach the routing plan to a different account.", 0 , "FNSDesigner"
	AHSID = ID.value
	MODE = document.body.getAttribute("ScreenMode")

	NodeSearchObj.AHSID = AHSID
	NodeSearchObj.Selected = false

	If AHSID = "" Then AHSID = "NEW"
	
	If AHSID = "NEW" And MODE = "RO" Then
		MsgBox "No AHS currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_ACCOUNT_HIERARCHY_STEP&SELECTONLY=TRUE&AHSID=" & AHSID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,NodeSearchObj ,"dialogWidth=650px; dialogHeight=700px; center=yes"
	If NodeSearchObj.AHSID <> ID.value then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.value = NodeSearchObj.AHSID
	end if
End Function

sub setTransmissionValues()
   dim strDestType, strTransmissionType
    document.FrmSave.destination_string.value     = ""
    document.FrmSave.alt_destination_string.value = ""
    document.FrmSave.sequence.value               = ""
    document.FrmSave.retry_count.value            = ""
    document.FrmSave.retry_wait_time.value        = ""
        
    strTransmissionType = document.FrmSave.transmission_type_id.value
    strDestType         = document.FrmSave.destination_type.value
    
    if strTransmissionType = 1 then ' "Fax" then
      
      document.FrmSave.alt_destination_string.value = "8009659825"
      document.FrmSave.sequence.value               = "1"
      document.FrmSave.retry_count.value            = "3"
      document.FrmSave.retry_wait_time.value        = "180"
      
      
      select case strDestType   
       case "Branch"
          document.FrmSave.destination_string.value     = "~CLAIM:BRANCH:PHONE_FAX~"
          document.all.DESCRIPTION.value  = "WOR " & "<%= Request.QueryString("CLIENT") %>" & " Fax to Branch"
      
       Case "Caller"
          document.FrmSave.destination_string.value     = "~CLAIM:CALLER:PHONE_FAX~"
          document.all.DESCRIPTION.value  = "WOR " & "<%= Request.QueryString("CLIENT") %>" & " Fax to Caller"
      
       Case "Risk Location"
          document.FrmSave.destination_string.value     = "~CLAIM:RISK_LOCATION:PHONE_FAX~"
          document.all.DESCRIPTION.value  = "WOR " & "<%= Request.QueryString("CLIENT") %>" & " Fax to Risk Location"
      
       Case "Insured"
          document.FrmSave.destination_string.value     = "~CLAIM:INSURED:PHONE_FAX~"
          document.all.DESCRIPTION.value  = "WOR " & "<%= Request.QueryString("CLIENT") %>" & " Fax to Insured"
       Case "Special"
       Case "State"
     end select
   end if
   
   if strTransmissionType = 2 then '"Print" then
       document.FrmSave.destination_string.value     = "\\cha0s00t\oper_hp5si_b"
       document.FrmSave.alt_destination_string.value = "\\cha0s2t\sqa_hp4050"
       document.FrmSave.retry_count.value            = "1"
       document.FrmSave.retry_wait_time.value        = "180"
       document.all.DESCRIPTION.value  = "WOR " & "<%= Request.QueryString("CLIENT") %>" & " Print to " & strDestType
      
   end if
   
   if strTransmissionType = 6 then ' "Email" then
       
       document.FrmSave.retry_count.value            = "3"
       document.FrmSave.retry_wait_time.value        = "5"
      
   end if
end sub

Sub BtnFindOD1_onclick
<% If MODE="RO" Then Response.write(" Exit Sub ") %>
	lret = window.showModalDialog( "../OutputDefiniton/OutputDefinitionMaintenance.asp?CONTAINERTYPE=MODAL"  ,DefinitionObj1 ,"dialogWidth:450px;dialogHeight:450px;center")
    if DefinitionObj1.ODID <> "" Then
		document.all.OUTPUTDEF_ID1.value = DefinitionObj1.ODID
		if document.all.SEQUENCE1.value = "" then
			document.all.SEQUENCE1.value = "10"
		end if
	    document.all.SEQUENCE2.disabled = false
		document.all.OUTPUTDEF_ID2.disabled = false
		'display image by setting width back to original
		document.all.BtnFindOD2.width = 16
		document.all.BtnDetachOD1.width = 16
		
	end if
End Sub

Sub BtnFindOD2_onclick
<% If MODE="RO" Then Response.write(" Exit Sub ") %>
	lret = window.showModalDialog( "../OutputDefiniton/OutputDefinitionMaintenance.asp?CONTAINERTYPE=MODAL"  ,DefinitionObj2 ,"dialogWidth:450px;dialogHeight:450px;center")
	if DefinitionObj2.ODID <> "" Then
		document.all.OUTPUTDEF_ID2.value = DefinitionObj2.ODID
		if document.all.SEQUENCE2.value = "" then
			document.all.SEQUENCE2.value = "20"
		end if
	    document.all.SEQUENCE3.disabled = false
		document.all.OUTPUTDEF_ID3.disabled = false
		document.all.BtnDetachOD2.width = 16
		document.all.BtnFindOD3.width = 16
		
	end if
End Sub

Sub BtnFindOD3_onclick
<% If MODE="RO" Then Response.write(" Exit Sub ") %>
	lret = window.showModalDialog( "../OutputDefiniton/OutputDefinitionMaintenance.asp?CONTAINERTYPE=MODAL"  ,DefinitionObj3 ,"dialogWidth:450px;dialogHeight:450px;center")
	if DefinitionObj3.ODID <> "" Then
		document.all.OUTPUTDEF_ID3.value = DefinitionObj3.ODID
		if document.all.SEQUENCE3.value = "" then
			document.all.SEQUENCE3.value = "30"
		end if
	document.all.BtnFindOD3.width = 16
	document.all.BtnDetachOD3.width = 16
	end if
	
End Sub

Function DetachOD1()
<% If MODE="RO" Then Response.Write(" Exit Function ") %>
	document.all.SEQUENCE1.value = "" 
	document.all.outputdef_id1.value = ""
	DetachOD2()
	document.all.BtnFindOD2.width = 0
	document.all.BtnDetachOD2.width = 0
	
	
End Function

Function DetachOD2()
<% If MODE="RO" Then Response.Write(" Exit Function ") %>
	document.all.SEQUENCE2.value = "" 
	document.all.outputdef_id2.value = ""
	DetachOD3()
	document.all.BtnFindOD3.width = 0
	document.all.BtnDetachOD3.width = 0
End Function

Function DetachOD3()
<% If MODE="RO" Then Response.Write(" Exit Function ") %>
	document.all.SEQUENCE3.value = "" 
	document.all.outputdef_id3.value = ""
End Function


Sub BtnSave_onclick
	if document.all.OUTPUTDEF_ID1.value = "" then
		msgbox("Output Definition is a required field.")
	elseif (document.all.description.value = "" or document.all.ENABLED_FLG.value = "" or document.all.sequence.value = "") then
		msgbox("One or more Required fields is null. Please enter it.")
		return false
	else
		FrmSave.Submit()
	end if
End Sub

Sub onClick_statesflg
   if document.all.allstates_flg.checked = "True" then
      'msgbox(document.all.allstates_flg.checked)
      'msgbox(document.all.state.length)
      strLength = document.all.state.length
      
      for i = 0 to document.all.Frmsave.state.length -1
	     document.all.Frmsave.state.options(i).selected = "True"
	  next
   
   
   else
      for i = 0 to document.all.Frmsave.state.length -1
        document.all.Frmsave.state.options(i).selected = "False"
     next
     document.all.Frmsave.state.options(0).selected = "True"
   end if 
End Sub

-->
</script>
<script LANGUAGE="JavaScript">
function CanDocUnloadNow()
{
	if (false == confirm("Data has changed. Leave page without saving?"))
		return false;
	else
		return true;
}
function CRuleSearchObj()
{
	this.RID = "";
	this.RIDText = "";
	this.RIDType = "";
	this.Selected = false;
}
function CNodeSearchObj()
{
	this.AHSID = "";
	this.Selected = false;
}
function COutputDefinitionSearchObj()
{
	this.ODID = "";
	this.ODIDName = "";
	this.Saved = false;	
	this.Selected = false;	
}
var DefinitionObj1 = new COutputDefinitionSearchObj();
var DefinitionObj2 = new COutputDefinitionSearchObj();
var DefinitionObj3 = new COutputDefinitionSearchObj();
var NodeSearchObj  = new CNodeSearchObj();
var RuleSearchObj  = new CRuleSearchObj();
</script>
</head>


<body BGCOLOR="<%=BODYBGCOLOR%>" topmargin="0" rightmargin="2" leftmargin="0" CanDocUnloadNowInf="NO" bottommargin="0" ScreenMode="<%= MODE %>">
<!--#include file="..\lib\NavBack.inc"-->
<form NAME="FrmSave" ACTION="WCRPWizardExecute.asp?ACTION=SAVE" METHOD="POST">
<input TYPE="HIDDEN" NAME="AHSid" value="<%=Request.QueryString("AHSID")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
   
   <tr><td colspan="2" HEIGHT="4"></td></tr>
   <tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» 
   Workers Comp Routing Plan Wizard</td></tr>
   
   <tr><td colspan="2" HEIGHT="4"></td></tr>
   <tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» 
   Routing Plan Summary
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align="absmiddle" title="Help" WIDTH="7" HEIGHT="8">
   </td>
   <td HEIGHT="5" ALIGN="LEFT">
</table>


<table CELLSPACING="2" CELLPADDING="0" BORDER="0">
<tr><td>
   <table BORDER="0" CELLSPACING="2" CELLPADDING="0" BORDER="1">
      <tr>
         <td CLASS="LABEL" COLSPAN="9">Description:<br><input TYPE="TEXT" NAME="DESCRIPTION" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER;' ") %> CLASS="LABEL" SIZE="80"></td>
      </tr>
      <tr> 
         <td CLASS="LABEL">Dest. Type:<br>
         <select STYLE="WIDTH:150" NAME="DESTINATION_TYPE" CLASS="LABEL" OnChange="setTransmissionValues()" <% If MODE="RO" Then Response.Write(" DISABLED ") %>>
         <option VALUE="Branch" SELECTED>BRANCH
         <option VALUE="Caller">CALLER
         <option VALUE="Risk Location">RISK_LOCATION
         <option VALUE="Insured">INSURED
         <option VALUE="Special">SPECIAL
         <option VALUE="State">STATE
         </select>
         </td>
         <td CLASS="LABEL"><nobr>Input System Name:<br>
         <select CLASS="LABEL" NAME="INPUT_SYSTEM_NAME" STYLE="WIDTH:100%" <% If MODE="RO" Then Response.Write(" DISABLED ") %>>
         <option VALUE="FNS NET" SELECTED>FNS NET
         <option VALUE="OPEN BASIC">OPEN BASIC
         <option VALUE="FNSINETP1">FNSINETP1
         </select>
         </td>
         <td CLASS="LABEL" ALIGN="LEFT" VALIGN="MIDDLE">A.H. Step ID:<br>
            <input READONLY TYPE="TEXT" CLASS="LABEL" NAME="ACCNT_HRCY_STEP_ID" STYLE="BACKGROUND-COLOR:SILVER" VALUE="<%= Request.QueryString("AHSID") %>" SIZE="10">
         </td>
         <td VALIGN="bottom" ALIGN="LEFT"><img src="..\Images\attach.GIF" TITLE="Attach Account Hierarchy Step" STYLE="CURSOR:HAND" align="absbottom" OnClick="AttachNode ACCNT_HRCY_STEP_ID">
         </td>
      </tr>
      <tr>
         <td CLASS="LABEL">LOB:<br>
            <input READONLY TYPE="TEXT" CLASS="LABEL" NAME="LOB_CD" STYLE="BACKGROUND-COLOR:SILVER" VALUE="WOR" SIZE="10">
         </td>
         
         <td CLASS="LABEL">Select All States:<br>
            <input TYPE="CHECKBOX" NAME="ALLSTATES_FLG" onClick="onClick_statesflg" <% If MODE="RO" Then Response.Write(" DISABLED ") %>>
         </td>
         
         <td CLASS="LABEL">State:<br>
            <select NAME="STATE" CLASS="LABEL" MULTIPLE <% If MODE="RO" Then Response.Write(" DISABLED ") %>>
            <!--#include file="..\lib\states.asp"-->
            </select>
         </td>
         
         
         <td CLASS="LABEL">Enabled:<br>
            <input TYPE="CHECKBOX" checked NAME="ENABLED_FLG" <% If MODE="RO" Then Response.Write(" DISABLED ") %>>
         </td>
      </tr>
           
   </table>
      </td>
      <td VALIGN="TOP">
      <table>
         <tr>
             <td CLASS="LABEL"><button CLASS="STDBUTTON" <% If MODE="RO" Then Response.Write(" DISABLED ") %>ACCESSKEY="S" NAME="BtnSave">Save</button></td>
         </tr>
      </table>
</td>
</tr>
</table>

<table WIDTH="100%">
<tr WIDTH="100%">
   <td CLASS="LABEL" NOWRAP>
      <img SRC="../IMAGES/Attach.gif" STYLE="CURSOR:HAND" NAME="BtnAttachEnabledRule" TITLE="Attach Rule" OnClick="AttachRule ENABLEDRULE_ID, ENABLEDRULE_ID_TEXT" WIDTH="16" HEIGHT="16">
      <img SRC="../IMAGES/Detach.gif" STYLE="CURSOR:HAND" NAME="BtnDetachEnabledRule" TITLE="Detach Rule" OnClick="DetachRule ENABLEDRULE_ID, ENABLEDRULE_ID_TEXT" WIDTH="16" HEIGHT="16">
      Enabled Rule: 
      <span ID="ENABLEDRULE_ID_TEXT" CLASS="LABEL" TITLE></span><input TYPE="HIDDEN" NAME="ENABLEDRULE_ID" VALUE="<%= ENABLERULE_ID %>">
   </td>
</tr>
</table>

<!--***************************************************-->

<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
   <table>
   <tr><td colspan="2" HEIGHT="4"></td></tr>
   <tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» <%= StatusLabel %> Transmission Sequence Step</td>
   <td HEIGHT="5" ALIGN="LEFT">
   </table>
 
   <table>
   <tr>
      <td CLASS="LABEL">Transmission Type:<br>
      <select NAME="TRANSMISSION_TYPE_ID" CLASS="LABEL" OnChange="setTransmissionValues()" <% If MODE="RO" Then Response.Write(" DISABLED ") %>>
         <option VALUE="1" SELECTED>FAX
         <option VALUE="2">PRINT
         <option VALUE="6">EMAIL
         <option VALUE="7">LEGACY EMAIL
      </select>
      </td>
   </tr>
   <tr>
      <td CLASS="LABEL">Destination String:<br><input TYPE="TEXT" CLASS="LABEL" NAME="DESTINATION_STRING" SIZE="60" MAXLENGTH="255" VALUE="~CLAIM:BRANCH:PHONE_FAX~" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER;'  ") %>  VALUE="<%= DESTINATION_STRING %>"></td>
   </tr>
   <tr>
      <td CLASS="LABEL">Alternate Destination String:<br><input TYPE="TEXT" CLASS="LABEL" NAME="ALT_DESTINATION_STRING" SIZE="60" MAXLENGTH="255" VALUE="8009659825" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER;'  ") %> VALUE="<%= ALT_DESTINATION_STRING %>"></td>
   </tr>
   <table>
      <tr>
        <td CLASS="LABEL">Sequence:<br><input TYPE="TEXT" CLASS="LABEL" NAME="SEQUENCE" SIZE="5" MAXLENGTH="10" VALUE="1" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER;' ") %> ></td>
        <td CLASS="LABEL">Retry Count:<br><input TYPE="TEXT" NAME="RETRY_COUNT" value="3" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER;' ") %> CLASS="LABEL" SIZE="10" MAXLENGTH="10" ></td>
        <td CLASS="LABEL">Retry Wait Time:<br><input TYPE="TEXT" CLASS="LABEL" NAME="RETRY_WAIT_TIME" SIZE="10" VALUE="180" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER;' ") %> MAXLENGTH="10" VALUE="<%= RETRY_WAIT_TIME %>" ></td>
      </tr>
   </table>
</table>


<!--***************************************************-->

<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
   <table>
      <tr><td colspan="2" HEIGHT="4"></td></tr>
      <tr><td CLASS="GrpLabel" WIDTH="134" HEIGHT="10"><nobr>&nbsp;» <%= StatusLabel %> Output Definitions :</td>
      <td HEIGHT="5" ALIGN="LEFT">
   </table>

   <table>
      <tr>
        <td CLASS="LABEL">Sequence:<br><input TYPE="TEXT" NAME="SEQUENCE1" SIZE="5" MAXLENGTH="10" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER;' ") %> CLASS="LABEL"></td>
        <td CLASS="LABEL" NOWRAP VALIGN="BOTTOM">Output Definition:<br><input READONLY TYPE="TEXT" SIZE="10" CLASS="LABEL" NAME="OUTPUTDEF_ID1" STYLE="BACKGROUND-COLOR:SILVER" VALUE>
        <img SRC="../IMAGES/Attach.gif" ID="BtnFindOD1" NAME="BtnFindOD1" TITLE="Attach Output Definition" STYLE="CURSOR:HAND" align="absbottom" WIDTH="16" HEIGHT="16">
        <img SRC="../IMAGES/Detach.gif" STYLE="CURSOR:HAND" ID="BtnDetachOD1" NAME="BtnDetachOD1" TITLE="Detach OD1" align="absbottom" OnClick="DetachOD1" WIDTH="16" HEIGHT="16">
        </TD>
        
      </tr>
      <tr></tr>
      <tr>
         <td CLASS="LABEL">Sequence:<br><input TYPE="TEXT" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> CLASS="LABEL" NAME="SEQUENCE2" SIZE="5" MAXLENGTH="10"></td>
         <td CLASS="LABEL" NOWRAP VALIGN="BOTTOM">Output Definition:<br><input READONLY TYPE="TEXT" SIZE="10" CLASS="LABEL" NAME="OUTPUTDEF_ID2" STYLE="BACKGROUND-COLOR:SILVER" VALUE>
         <img SRC="../IMAGES/Attach.gif" ID="BtnFindOD2" NAME="BtnFindOD2" TITLE="Attach Output Definition" STYLE="CURSOR:HAND" align="absbottom" WIDTH="16" HEIGHT="16" style="z-index:2">
         <img SRC="../IMAGES/Detach.gif" STYLE="CURSOR:HAND" ID="BtnDetachOD2" NAME="BtnDetachOD2" TITLE="Detach OD2" align="absbottom" OnClick="DetachOD2" WIDTH="16" HEIGHT="16"></td>
      </tr>
      <tr></tr>
      <tr>
         <td CLASS="LABEL">Sequence:<br><input TYPE="TEXT" <% If MODE="RO" Then Response.Write(" READONLY STYLE='BACKGROUND-COLOR:SILVER' ") %> CLASS="LABEL" NAME="SEQUENCE3" SIZE="5" MAXLENGTH="10"></td>
         <td CLASS="LABEL" NOWRAP  VALIGN="BOTTOM">Output Definition:<br><input READONLY TYPE="TEXT" SIZE="10" CLASS="LABEL" NAME="OUTPUTDEF_ID3" STYLE="BACKGROUND-COLOR:SILVER" VALUE>
         <img SRC="../IMAGES/Attach.gif" ID="BtnFindOD3" TITLE="Attach Output Definition" STYLE="CURSOR:HAND" align="absbottom" WIDTH="16" HEIGHT="16">
         <img SRC="../IMAGES/Detach.gif" STYLE="CURSOR:HAND" ID="BtnDetachOD3" NAME="BtnDetachOD3" TITLE="Detach OD3" align="absbottom" OnClick="DetachOD3" WIDTH="16" HEIGHT="16"></td>
      </tr>
   </table>
</table>

</form>
</body>
</html>
