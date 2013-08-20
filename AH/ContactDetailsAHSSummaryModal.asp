<!--#include file="..\lib\genericSQL.asp"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->
<!-- #include file="..\lib\RenderTextinc.asp"-->



<%
	
	Response.Expires = 0
	Response.AddHeader  "Pragma", "no-cache"
	
	Dim ContainerType, Mode, AHSCID
	Dim s_SQL2, Conn, ConnectionString, rs, RS2, s_Select
	Dim s_SeqStepID,s_DestinationID,s_SeqNumber ,s_RetryCount ,s_RetryWaitTime ,s_DestinationStr,s_Transmission 
    Dim AccountTextLen , s_msg

	AccountTextLen = 28
   
   '** Check for appropriate privs...	
   'If HasModifyPrivilege("FNSD_SPECIFICDESTINATION", SECURITYPRIV) <> True Then MODE = "RO"
   ' Mode in this case is either EDIT or NEW
   '******

    Mode = Request.Querystring("MODE")
	AHSCID=Request.Querystring("AHSCID")
	
	
    ' *****************************************************************
	' The following gets executed when posted to self 
	' *****************************************************************
	' The following gets executed when posted to self .
	' Diff scenarios:
	' 1) posted to self when ahscid is known from the first calling scr.( update)
	' 2) posted to self when the AHSCID = "new" from the first calling scr.( insert)
	' 3) posted to self when the AHSCID = newly generated AHSCID ( update after insert without leaving the screen)
	'    
	
	if Request.Querystring("CHECK") = "Save" then
	    
	    '*************************************************
  	    ' DMS: 3/30/00 
	    ' The DB design has been modified and this file takes care
	    ' of edits and inserts into AHS_CONTACT 
	    '*************************************************
    
	   	On Error Resume Next
	    dim NewAHSCOID, InsertSQL, ACTION, SQL_STRING, UpdateSQL,strError
	
		ACTION     = CStr(Request.Form("Txt_Operation"))
		SQL_STRING = Request.Form("Txt_SQLString")
	    
	    If ACTION  = "UPDATE" Then
		   UpdateSQL    = BuildSQL(SQL_STRING,Chr(128), Chr(129), "UPDATE", "AHS_CONTACT", "AHS_CONTACT_ID", "")		 
		   Set RSinsert = Conn.Execute(UpdateSQL)
		   strError     = CheckADOErrors(Conn,"AHS_CONTACT " & ACTION)
	    Elseif ACTION = "INSERT" Then
		   InsertSQL    = ""
		   NewAHSCOID   = NextPkey("AHS_CONTACT","AHS_CONTACT_ID")
		   InsertSQL    = BuildSQL(SQL_STRING, Chr(128), Chr(129), "INSERT", "AHS_CONTACT", "AHS_CONTACT_ID", NewAHSCOID)		 
		   
		   Set RSUpdate = Conn.Execute(InsertSQL)
		   if NewAHSCOID > 0 Then 
	          strError = CheckADOErrors(Conn,"AHS_CONTACT " & ACTION)
	       Else
		      strError = "Unable to obtain next primary key for AHS_CONTACT table."
           End If			
		End If
			    
		If strError <> ""  Then
		   s_msg = strError
		Else
		   s_msg = " Update Successful"
  	   	End If
		
        if AHSCID = "" then 
		   AHSCID = NewAHSCOID
		end if
	    
      	Conn.Close
	End if
'*************************************************
	   ' If EDIT was clicked then display data for that COID etc..
	   ' Only 1 row shd. be retrieved since we r specifying AHSCID
       
	   
	   IF Request.QueryString("MODE") <> "NEW" OR  ( Request.QueryString("MODE") = "NEW" AND Request.Querystring("CHECK") = "Save" ) THEN
		   
	       SET Conn = Server.CreateObject("ADODB.Connection")
		   ConnectionString = CONNECT_STRING
		   
		   s_Select = "Select ahsc.ahs_contact_id, " &_
		           "       ahsc.contact_id, " &_
				   "	   ahsc.accnt_hrcy_step_id,"  &_ 
				   "	   to_char(ahsc.active_start_dt,'MM/DD/YYYY') as active_start_dt, " &_
				   "	   to_char(ahsc.active_end_dt,'MM/DD/YYYY') as active_end_dt, " &_
	               "       ahs.name  "  &_
				   "  From ahs_contact ahsc, account_hierarchy_step ahs " &_
		           " Where ahs.accnt_hrcy_step_id = ahsc.accnt_hrcy_step_id " &_
				   "   and ahs_contact_id         = " & AHSCID
				
				   
		   Conn.Open ConnectionString
		   'response.write(s_Select)
		

		   SET rs = Conn.Execute(s_Select)	
		   If not rs.EOF Then
		     s_ahs_contact_id     = rs("ahs_contact_id")
		     s_contact_id         = rs("contact_id")
		     s_accnt_hrcy_step_id = rs("accnt_hrcy_step_id")
		     s_active_start_dt    = rs("active_start_dt")
		     s_active_end_dt      = rs("active_end_dt")
		     s_account_name       = rs("name")
		     
		   
		  End If
		  rs.Close
		  Conn.Close
		  SET rs = NOTHING
		  SET Conn = NOTHING
	
	 END IF
     'response.write(request.querystring)
	
%>

<HTML>
<HEAD>
<!--#include file="..\lib\tablecommon.inc"-->
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<Title>AHS Contacts</Title>
<Link Rel="StyleSheet" Type="text/css" Href="..\FNSDESIGN.CSS">
<script>


function CAHSSearchObj()
{
	this.AHSID = "";
	this.AHSIDName = "";
	this.Selected = false;	
}
var AHSSearchObj = new CAHSSearchObj();
</script>
<Script Language=VBScript>
<!--#include file="..\lib\Help.asp"-->

Dim g_StatusInfoAvailable
	g_StatusInfoAvailable = false
	
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
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null, "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If
End Sub

Function ExeSave
	
	sResult = ""
	ExeSave = false	
		
	IF document.all.spanSeqStepID.innerTEXT = "NEW" then
	   document.all.Txt_Operation.value = "INSERT"
	else
	   document.all.Txt_Operation.value = "UPDATE"
	end if
	
	If Not f_ValidateScreenData Then EXIT FUNCTION
	
	'Update
	sResult = sResult & "AHS_CONTACT_ID"& Chr(129) & document.all.spanSeqStepID.innerTEXT & Chr(129) & "0" & Chr(128)
	sResult = sResult & "ACTIVE_START_DT" & Chr(129) & "to_date('" & document.all.inpActivestdt.value & "','MM/DD/YYYY')" & Chr(129) & "0" & Chr(128)
	sResult = sResult & "ACTIVE_END_DT"& Chr(129) & "to_date('" & document.all.inpActiveenddt.value & "','MM/DD/YYYY') " & Chr(129) & "0" & Chr(128)
	sResult = sResult & "ACCNT_HRCY_STEP_ID"& Chr(129) & document.all.AHSID_ID.innerTEXT & Chr(129) & "1" & Chr(128)
	
	If document.all.Txt_Operation.value = "INSERT" then
	   sResult = sResult & "CONTACT_ID"& Chr(129) & document.all.spanSpDestID.innerTEXT & Chr(129) & "1" & Chr(128)
	end if 
	document.all.Txt_SQLString.value = sResult
	document.all.frmAHSSummaryModal.Submit()
    ExeSave = True
End Function
	
Sub BtnSaveSeqStep_OnClick
	b_Save = ExeSave
End Sub

Sub BtnCloseSeqStep_OnClick
    'window.close
	parent.window.close
End Sub

Sub updateStatus(s_Msg)
	spanSeqStepID.innerHTML = s_Msg
End Sub

Function f_ValidateScreenData()
   f_ValidateScreenData = True
   
   if  document.all.AHSID_ID.innerTEXT = "" Then
      msgbox("Please attach the AHS Contact to an Account")
	  f_ValidateScreenData = False
   end if 

   if document.all.inpActivestdt.value <> "" then 

       if NOT IsDate(document.all.inpActivestdt.value) then

	      msgbox("Active Start Date has an incorrect format. Format as MM/DD/YYYY" & VBCRLF)
	      f_ValidateScreenData = False
	   end if
	end if
    if document.all.inpActiveenddt.value <> "" then 
       if NOT IsDate(document.all.inpActiveenddt.value) then
	      msgbox("Active End Date has an incorrect format. Format as MM/DD/YYYY" & VBCRLF)
	      f_ValidateScreenData = False
	   end if
	end if


End Function


Function AttachAccount (ID, SPANID)
    
	AHSID = ID.innerText
	
	MODE = document.body.getAttribute("ScreenMode")

	AHSSearchObj.AHSID = AHSID
	AHSSearchObj.AHSIDName = SPANID.title
	AHSSearchObj.Selected = false

	If AHSID = "" Then AHSID = "NEW"
	
	If AHSID = "NEW" And MODE = "RO" Then
		MsgBox "No account currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_CONTACT&SELECTONLY=TRUE&AHSID=" &AHSID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,AHSSearchObj ,"center"

	'if Selected=true update everything, otherwise if AHSID is the same, update text in case of save
	If AHSSearchObj.Selected = true Then
		If AHSSearchObj.AHSID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = AHSSearchObj.AHSID
		end if
		UpdateSpanText SPANID,AHSSearchObj.AHSIDName
	ElseIf ID.innerText = AHSSearchObj.AHSID And AHSSearchObj.AHSID<> "" Then
		UpdateSpanText SPANID,AHSSearchObj.AHSIDName
	End If

End Function

Function Detach(ID, SPANID)
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		ID.innerText = ""
		SPANID.innerText = ""
	end if
End Function

Sub UpdateSpanText (SPANID, inText)
	If Len(inText) < <%=AccountTextLen%> Then
		SPANID.innertext = inText
	Else
		SPANID.innertext = Mid ( inText, 1, <%=AccountTextLen%>) & " ..."
	End If
	SPANID.title = inText
End Sub



</Script>

</HEAD>
<BODY  topmargin=0 leftmargin=0  rightmargin=0  BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<TABLE WIDTH=100% CELLPADDING="0" CELLSPACING="0">
  <TR>
     <TD colspan=2 HEIGHT=4></TD>
  </TR>
  <TR>
     <TD CLASS=GrpLabel WIDTH="134" HEIGHT=10><NOBR>&nbsp&#187 AHS Contacts Detail&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG SRC="../Images/HelpIcon2.gif" STYLE="CURSOR:HAND" align=absmiddle title=Help OnClick='LaunchHelp("Welcome.htm")'></td>
     <TD HEIGHT=5 ALIGN=LEFT>
        <TABLE CELLPADDING="0" CELLSPACING="0" HEIGHT=100%>
           <TR>
	          <TD WIDTH=3 HEIGHT=4></TD><TD WIDTH=300 HEIGHT=4></TD>
	       </TR>
	       <TR>
	          <TD CLASS=GrpLabelDrk WIDTH=3 HEIGHT=8 VALIGN=BOTTOM ALIGN=LEFT></TD>
	          <TD WIDTH=300 HEIGHT=8></TD>
	       </TR>
        </TABLE>
    </TD>
  </TR>
  <TR><TD CLASS=GrpLabelLine colspan=2 HEIGHT=1></TD></TR>
  <TR><TD colspan=2 HEIGHT=1></TD></TR>
</TABLE>

<TABLE style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
  <tr><td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18"><img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report"></td>
  <%if Request.Querystring("CHECK") = "Save" then %>
      <td WIDTH="385">:<SPAN ID="SpanStatusSeqStep" STYLE="COLOR:#006699" CLASS=LABEL><%=s_msg%></SPAN></td></tr>
  <%else%>
      <td WIDTH="385">:<SPAN ID="SpanStatusSeqStep" STYLE="COLOR:#006699" CLASS=LABEL>Ready</SPAN></td></tr>
  <%end if%>
</TABLE>
<BR>

<Table class="LABEL">
   <tr><td width=305 nowrap>Contact ID:&nbsp;<SPAN ID="spanSpDestID" CLASS=LABEL><%= Request.QueryString("COID") %></SPAN></td>
</Table>
<%if Request.Querystring("CHECK") = "Save" then %>
<SPAN CLASS=LABEL>AHS Contact ID:&nbsp<span id="spanSeqStepID"><%=AHSCID%></span>
<%else%>
<SPAN CLASS=LABEL>AHS Contact ID:&nbsp<span id="spanSeqStepID"><%=Request.QueryString("AHSCID")%></span>
<%end if%>
<TABLE CLASS="LABEL" >
  <tr>
	<td>	
	   <IMG NAME=BtnAttachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Attach.gif" TITLE="Attach Account" ONCLICK="VBScript::AttachAccount AHSID_ID, AHSID_TEXT">
	   <IMG NAME=BtnDetachAHSID STYLE="cursor:hand" SRC="..\IMAGES\Detach.gif" TITLE="Detach Account" OnClick="VBScript::Detach AHSID_ID, AHSID_TEXT">
	   Account:&nbsp;<SPAN ID=AHSID_TEXT CLASS=LABEL TITLE="<%=ReplaceQuotesInText(s_ACCOUNT_NAME)%>"><%=TruncateText(s_ACCOUNT_NAME,28)%></SPAN>
	   A.H.Step ID:&nbsp;<SPAN ID=AHSID_ID CLASS=LABEL ><%=s_accnt_hrcy_step_id%></SPAN>
	   
	</td>
  </tr>
	
<tr></tr>
	<tr><td>	
		<Table CLASS="LABEL" cellpadding="2" cellspacing="2">
			<TR><td CLASS=LABEL width=100>Active Start Date:<br><input ScrnInput="TRUE" size=15 CLASS="LABEL" MAXLENGTH=15 TYPE="TEXT" NAME="inpActivestdt" VALUE="<%=s_active_start_dt %>" ></td>
				<td CLASS=LABEL width=100>Active End date:<br><input ScrnInput="TRUE" size=15 CLASS="LABEL" MAXLENGTH=15 TYPE="TEXT" NAME="inpActiveenddt" VALUE="<%=s_active_end_dt %>" ></td>
            </tr>  		
			<table >	 
				<tr><td CLASS="LABEL"><button CLASS="STDBUTTON"  NAME="BtnSaveSeqStep" ACCESSKEY="S"><u>S</u>ave</button></td>
                    <td CLASS="LABEL"><button CLASS="STDBUTTON" NAME="BtnCloseSeqStep" ACCESSKEY="L">C<u>l</u>ose</button>
				</tr>
			</table>	
				
		</Table>
	</tr>
</TABLE>
<Form Name="frmAHSSummaryModal"  <% if Request.Querystring("AHSCID") <> "NEW" then %>
ACTION="ContactDetailsAHSSummaryModal.asp?CHECK=Save&MODE=EDIT&AHSCID=<%=AHSCID%>&COID=<%=Request.Querystring("COID")%>" 
<% else %>
ACTION="ContactDetailsAHSSummaryModal.asp?CHECK=Save&MODE=NEW&COID=<%=Request.Querystring("COID")%>" 
<% end if%>
METHOD="POST" >
	<Input Type="Hidden" Name="Txt_SQLString">
	<Input Type="Hidden" Name="Txt_Operation">
	
</Form>
</BODY>
</HTML>
