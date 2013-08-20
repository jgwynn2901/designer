<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<%	
	Response.Expires=0 
	
	AccountTextLen = 4000	
	
	Dim UID, GID
	
	UID =  CStr(Request.QueryString("UID"))
	
	ACCID =  CStr(Request.QueryString("ACCID"))
	
    dim sqlp ,Conn ,RsLOCATION_USER_ID,RsACCOUNT_USER_ID,RsACCNT_HRCY_STEP_ID 
    dim  RsFNS_CLIENT_CD,RsPHONE_NUMBER,RsGREETING,RsLOB ,RsLOCATION_AHSID ,RsNam 
    
         
    %>
  
 

<html>
<head>
<meta name="VI60_defaultClientScript" content="VBScript">
<title>Access Permissions</title>
<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">




<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
dim g_StatusInfoAvailable
g_StatusInfoAvailable = false

Sub window_onload
<%	if CStr(Request.QueryString("MODE")) = "RO" then %>
	SetScreenFieldsReadOnly true,"DISABLED"
<%	end if %>
End Sub

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

Function CheckDirty
	if CStr(document.body.getAttribute("ScreenDirty")) = "YES" then 
		CheckDirty = true
	else
		CheckDirty = false
	end if
End Function

Sub SetDirty
	document.body.setAttribute "ScreenDirty", "YES"
End Sub

Sub ClearDirty
	document.body.setAttribute "ScreenDirty", "NO"
End Sub

Function ValidateScreenData
	If  document.all.AHSID_ID.innerText = "" then
		MsgBox "Locations AHSID is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	End If
	If  document.all.TxtFnsClientCD.value = "" then
		MsgBox "Client CD  is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	End If
	
	If  document.all.TxtPhoneNumber.value = "" then
		MsgBox "Phone Number is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	End If
	
	if <%=instr(CONNECT_STRING,"SED")%> > 0 THEN 
		If  document.all.GreetingText.innerText = "" then
		MsgBox "Greeting is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
		End If
	else
	
		If  document.all.TxtGreeting.value = "" then
		MsgBox "Greeting is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
		End If
	
	END IF
	
	If  document.all.TxtNAME.value = "" then
		MsgBox "NAME is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	End If
	If  document.all.SearchLOB_CD.value = "" then
		MsgBox "LOB is a required field.",0,"FNSNetDesigner"
		ValidateScreenData = false
		exit Function
	End If
	ValidateScreenData = true
End Function

Function ExeSave

	sResult = ""
	bRet = false
	
	If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
		MsgBox "Edit mode not selected.",0,"FNSNetDesigner"
		ExeSave = bRet
		exit Function
	End If

'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		'document.all.TxtAction.value = "INSERT"
		
	'If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
		If document.all.ACCID.value = "NEW" then
			document.all.TxtAction.value = "INSERT"
		else
			document.all.TxtAction.value = "UPDATE"
		end if
		
		
		if ValidateScreenData = false then 
			ExeSave = false
			exit function
		end if
	
		
		
		sResult = sResult & "LOCATION_USER_ID" & Chr(129) & document.all.SpanACCID.innerText & Chr(129) & "0" & Chr(128)
		sResult = sResult & "ACCNT_HRCY_STEP_ID" & Chr(129) & document.all.TxtAHSID.value & Chr(129) & "1" & Chr(128)
	    sResult = sResult & "ACCOUNT_USER_ID" & Chr(129) & document.all.spanUID.innerText  & Chr(129) & "1" & Chr(128)
	    sResult = sResult & "LOCATION_AHSID" & Chr(129) & document.all.AHSID_ID.innerText & Chr(129) & "1" & Chr(128)
		sResult = sResult & "LOB_CD" & Chr(129) & document.all.SearchLOB_CD.value & Chr(129) & "1" & Chr(128)
        sResult = sResult & "FNS_CLIENT_CD" & Chr(129) & document.all.TxtFnsClientCD.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "NAME" & Chr(129) & document.all.TxtNAME.value & Chr(129) & "1" & Chr(128)
		sResult = sResult & "PHONENUMBER" & Chr(129) & document.all.TxtPhoneNumber.value & Chr(129) &  "1" & Chr(128)
		if <%=instr(CONNECT_STRING,"SED")%> > 0 THEN 
		 sResult = sResult & "GREETINGS_ID" & Chr(129) & document.all.GreetingId.innerText & Chr(129) & "1" & Chr(128)
		 sResult = sResult & "GREETING" & Chr(129) & document.all.GreetingText.innerText & Chr(129) & "1" & Chr(128)
		 else
		 sResult = sResult & "GREETING" & Chr(129) & document.all.TxtGreeting.value & Chr(129) & "1" & Chr(128)
		 END IF
		
		
		
		document.all.TxtSaveData.Value = sResult
		document.body.setAttribute "ScreenDirty", "NO"
		document.all.FrmAccessPermissions.Submit()
		bRet = true
	
	
	ExeSave = bRet
	
End Function

sub Control_OnChange
	if document.body.getAttribute("ScreenMode") <> "RO" then
		document.body.setAttribute "ScreenDirty", "YES"	
		UpdateStatus("Ready")
	end if
end sub

sub SetScreenFieldsReadOnly(bReadOnly)
	for iCount = 0 to document.all.length-1
		if document.all(iCount).getAttribute("ScrnBtn") = "TRUE" then
			document.all(iCount).disabled = bReadOnly
		end if
	next
end sub

Sub StatusRpt_OnClick
	If g_StatusInfoAvailable = true Then
		lret = window.showModalDialog ("..\StatusRpt\StatusRpt.asp", Null,  "dialogWidth=580px; dialogHeight=380px; center=yes")
	Else
		MsgBox "No other detail status reported.",0,"FNSNetDesigner"	
	End If
End Sub


Function AttachGreetingsID(ID, SPANID)
'msgbox "WARNING: Changing the Location Step ID will attach the routing plan to a different account.", 0 , "FNSDesigner"

	MyGreetingID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	GreetingsSearchObj.GreetingID = MyGreetingID
	GreetingsSearchObj.GreetingText = SPANID.TITLE
	GreetingsSearchObj.Selected = false

	If MyGreetingID = "" Then MyGreetingID = "NEW"
	
	If MyGreetingID = "NEW" And MODE = "RO" Then
		MsgBox "No Greeting currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If
	
	strURL = "..\MyGreetings\MyGreetingMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_GREETING&SELECTONLY=TRUE&GreetingID=" & MyGreetingID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,GreetingsSearchObj ,"dialogWidth=650px; dialogHeight=700px; center=yes"
	
	If GreetingsSearchObj.Selected = true Then
	
		If GreetingsSearchObj.GreetingID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = GreetingsSearchObj.GreetingID
		end if
		UpdateSpanText SPANID,GreetingsSearchObj.GreetingText
	ElseIf ID.innerText = GreetingsSearchObj.GreetingId And GreetingsSearchObj.GreetingId<> "" Then
		UpdateSpanText SPANID,GreetingsSearchObj.GreetingText
	End If
		
End Function

Sub UpdateSpanText (SPANID, inText)
	If Len(inText) < <%=AccountTextLen%> Then
		SPANID.innertext = inText
	Else
		SPANID.innertext = Mid ( inText, 1, <%=AccountTextLen%>) & " ..."
	End If
	SPANID.title = inText
End Sub

Function AttachNODE(ID)
msgbox "WARNING: Changing the Location Step ID will attach the routing plan to a different account.", 0 , "FNSDesigner"

	AHSID = ID.innerText
	MODE = document.body.getAttribute("ScreenMode")

	AHSSearchObj.AHSID = AHSID
	AHSSearchObj.Selected = false

	If AHSID = "" Then AHSID = "NEW"
	
	If AHSID = "NEW" And MODE = "RO" Then
		MsgBox "No account currently attached.",0,"FNSNetDesigner"
		Exit Function
	End If


	
	strURL = "..\AH\AHSMaintenance.asp?CONTAINERTYPE=MODAL&SECURITYPRIV=FNSD_USERS&SELECTONLY=TRUE&AHSID=" &AHSID
	if MODE = "RO" then strURL = strURL & "&DETAILONLY=TRUE"
	
	showModalDialog  strURL  ,AHSSearchObj ,"dialogWidth=650px; dialogHeight=700px; center=yes"
		If AHSSearchObj.Selected = true Then
			If AHSSearchObj.AHSID <> ID.innerText then
			document.body.setAttribute "ScreenDirty", "YES"	
			ID.innerText = AHSSearchObj.AHSID
			end if
	'	UpdateSpanText SPANID,AHSSearchObj.AHSIDName
'	ElseIf ID.innerText = AHSSearchObj.AHSID And AHSSearchObj.AHSID<> "" Then
		'UpdateSpanText SPANID,AHSSearchObj.AHSIDName
	'	ID.value = AHSSearchObj.AHSID
	End If
	
End Function


</script>
<script language =jscript>
function CGreetingsSearchObj()
{
	this.GreetingID = "";
	this.GreetingText = "";
	this.Selected = false;
}



function CAHSSearchObj()
{
	this.AHSID = "";
	//this.AHSIDName = "";
	this.Selected = false;	
}

var AHSSearchObj = new CAHSSearchObj();
var GreetingsSearchObj = new CGreetingsSearchObj();
</script>
</head>


<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0">
<tr><td colspan="2" HEIGHT="4"></td></tr>

<td HEIGHT="5" ALIGN="LEFT">
<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%">
<tr><td WIDTH="3" HEIGHT="4"></td><td WIDTH="300" HEIGHT="4"></td></tr>
<tr><td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
<td WIDTH="300" HEIGHT="8"></td></tr>
</table></td></tr>
<tr><td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td></tr>
<tr><td colspan="2" HEIGHT="1"></td></tr>
</table>

<form Name="FrmAccessPermissions" METHOD="POST" ACTION="AccessLocationsSave.asp" TARGET="hiddenPage">
<input TYPE="HIDDEN" NAME="TxtSaveData">
<input TYPE="HIDDEN" NAME="TxtAction">

<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>">
<input type="hidden" NAME="UID" value="<%=Request.QueryString("UID")%>">
<input type="hidden" NAME="ACCID" value="<%=Request.QueryString("ACCID")%>">


<%
if  ACCID <> ""   then
      if  ACCID <> "NEW" then   
      
       Set Conn = Server.CreateObject("ADODB.Connection")
	         Conn.Open CONNECT_STRING
        
	if instr(CONNECT_STRING,"SED")> 0 THEN
	      SQLST = "SELECT LU.LOCATION_USER_ID, LU.ACCOUNT_USER_ID,LU.ACCNT_HRCY_STEP_ID,"
          SQLST = SQLST & " LU.FNS_CLIENT_CD, LU.PHONENUMBER, "
          SQLST = SQLST & " G.TEXT AS GREETING, LU.LOB_CD , LU.LOCATION_AHSID ,LU.NAME, G.GREETINGS_ID "
          SQLST = SQLST & " FROM LOCATIONS_USER LU, GREETINGS G "
          SQLST = SQLST & "  WHERE LU.LOCATION_USER_ID = " & ACCID
          SQLST = SQLST & " AND LU.GREETINGS_ID = G.GREETINGS_ID "
     
   else
         SQLST = "SELECT LOCATION_USER_ID, ACCOUNT_USER_ID,ACCNT_HRCY_STEP_ID,"
          SQLST = SQLST & " FNS_CLIENT_CD, PHONENUMBER, "
          SQLST = SQLST & " GREETING, LOB_CD , LOCATION_AHSID ,NAME "
          SQLST = SQLST & " FROM LOCATIONS_USER "
          SQLST = SQLST & "  WHERE LOCATION_USER_ID = " & ACCID
   end if
          
		 Set RS = Conn.Execute(SQLST)
		
	          If Not RS.EOF Then
		            RsLOCATION_USER_ID       =  "" & ReplaceQuotesInText(RS("LOCATION_USER_ID"))
			        RsACCOUNT_USER_ID        =  "" & ReplaceQuotesInText(RS("ACCOUNT_USER_ID"))
			        RsACCNT_HRCY_STEP_ID     =  "" & RS("ACCNT_HRCY_STEP_ID")
			        RsFNS_CLIENT_CD          =  "" & RS("FNS_CLIENT_CD")
			        RsPHONE_NUMBER           =  "" & RS("PHONENUMBER")
			        RsGREETING               =  "" & ReplaceQuotesInText(RS("GREETING"))                                       
                    RsLOB                    =  "" & ReplaceQuotesInText(RS("LOB_CD"))                                        
                    RsLOCATION_AHSID         =  "" & ReplaceQuotesInText(RS("LOCATION_AHSID")) 
                    RsName                   =  "" & ReplaceQuotesInText(RS("NAME"))   
                    if instr(CONNECT_STRING,"SED")> 0 THEN 
                    RsGreetings_ID			= "" & RS("GREETINGS_ID")
                    end if                      
            End If
             RS.Close
		     Set RS = Nothing
		     Set Conn = Nothing
        
  End If
 
%>

<table LANGUAGE="JScript" ONDRAGSTART="return false;" style="{position:absolute;top:20;}" class="Label" WIDTH="100%" CELLSPACING="0" CELLPADDING="0">
<tr>
<td ALIGN="LEFT" VALIGN="CENTER" WIDTH="18">
<img ID="StatusRpt" SRC="../users/..\images\StatusRpt.gif" width="16" height="16" ALIGN="LEFT" VALIGN="BOTTOM" ALT="View Status Report">
</td>
<td WIDTH="385">
:<span ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL">Ready</span>
</td>
</tr>
</table>

<table CLASS="LABEL" border=0  >
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr></tr>
<tr>
 <tr><td>User ID:&nbsp;<span id="spanUID"><%=Request.QueryString("UID")%></span></td></tr>
          
    
   <% If UID <> ""  Then 
   'call SelectAllRecordSet()
       Set Conn = Server.CreateObject("ADODB.Connection")
		  Conn.Open CONNECT_STRING
		  SQLST1 = "SELECT ACCNT_HRCY_STEP_ID FROM ACCOUNT_USER WHERE USER_ID = " & UID
	       Set RS1 = Conn.Execute(SQLST1)
		          If Not RS1.EOF Then
			         RSASHID=  "" & RS1("ACCNT_HRCY_STEP_ID")
	              End If

        
         
		  RS1.Close
		   Set RS1 = Nothing
		   Set Conn = Nothing

      end if	%>
       <!--<tr><td>Location User ID:&nbsp;<span id="SpanLID"><%=RsLOCATION_USER_ID%></span></td></tr>-->
       <tr><td>Location User ID:&nbsp;<span id="SpanACCID"><%=Request.QueryString("ACCID")%></span></td></tr>
      
       
   <td CLASS="LABEL"><B>AHSID:</B><br>
	     <input ScrnInput="TRUE" 
	            size="20" 
	            MAXLENGTH="10" 
	            TYPE="TEXT" 
	            NAME="TxtAHSID" 
	            VALUE="<%=RSASHID%>" 
	            TABINDEX = 1
	            STYLE="BACKGROUND-COLOR:SILVER" 
	            READONLY CLASS="READONLY"
	            ONKEYPRESS="VBScript::Control_OnChange" 
	            ONCHANGE="VBScript::Control_OnChange" ID="Text1">
	  </td>
    
    <TD CLASS=LABEL ALIGN=LEFT VALIGN=MIDDLE >

		<IMG src="..\Images\attach.GIF" TITLE="Attach Account Hierarchy Step" STYLE="CURSOR:HAND" align=absbottom OnClick='AttachNode AHSID_ID'>
		<br>
		<B>Location AHS ID:</B><BR><SPAN ID=AHSID_ID CLASS=LABEL><%=RsLOCATION_AHSID%></SPAN><INPUT type=hidden CLASS=LABEL NAME=TxtLocationAhsID STYLE="BACKGROUND-COLOR:SILVER" VALUE="<%=RsLOCATION_AHSID%>" SIZE=10 ID="Text6">
     </TD>
     <td CLASS="LABEL"><b>Client Code:</b><br>
	       <input ScrnInput="TRUE" 
	           size="3" 
	           CLASS="LABEL" 
	           MAXLENGTH="10" 
	           TYPE="TEXT" 
	           NAME="TxtFnsClientCD" 
	           VALUE="<%=RsFNS_CLIENT_CD %>" 
	           TABINDEX =3
	           ONKEYPRESS="VBScript::Control_OnChange" 
	           ONCHANGE="VBScript::Control_OnChange" ID="Text3">
	    </td>
	 <tr></tr>
	<tr>         
	    <td CLASS="LABEL"><b>Phone Number:</b><br>
	      <input ScrnInput="TRUE" 
	            size="20" 
	            CLASS="LABEL" 
	            MAXLENGTH="40" 
	            TYPE="TEXT" 
	            NAME="TxtPhoneNumber" 
	            VALUE="<%= RsPHONE_NUMBER %>" 
	            TABINDEX = 4
	            ONKEYPRESS="VBScript::Control_OnChange" 
	            ONCHANGE="VBScript::Control_OnChange" ID="Text4">
	      </td>
	      <td CLASS="LABEL"><b>Name:</b><br>
	      <input ScrnInput="TRUE" 
	            size="40" 
	            CLASS="LABEL" 
	            MAXLENGTH="30" 
	            TYPE="TEXT" 
	            NAME="TxtNAME" 
	            VALUE="<%=RsNAME %>" 
	            TABINDEX = 5
	            ONKEYPRESS="VBScript::Control_OnChange" 
	            ONCHANGE="VBScript::Control_OnChange" ID="Text4">
	      </td>
	      <td CLASS="LABEL"><B>LOB:<B><br>
	      <select NAME="SearchLOB_CD" 
	              CLASS="LABEL" 
	              tabindex=6 
	              ID="Select1"><%=GetControlDataHTML("LOB","LOB_CD","LOB_CD",RsLOB ,true)%> 
	              </select></td>
	   </tr></table>
	 
 
 <table ID="Table5" border ="0"><tr>
<%	if instr(CONNECT_STRING,"SED")> 0 THEN %>
<tr> <IMG src="..\Images\attach.GIF" TITLE="Attach Greetings Id" STYLE="CURSOR:HAND" align=absbottom OnClick='AttachGreetingsID GreetingId,GreetingText'></TD>
 <td class = "LABEL"><B>Greeting Id:</B> <SPAN ID="GreetingId" CLASS=LABEL><%=RsGreetings_ID%></SPAN><INPUT READONLY TYPE=hidden CLASS=LABEL NAME=TxtGreetingsID STYLE="BACKGROUND-COLOR:SILVER" VALUE="<%=RsGreetings_ID%>" SIZE=10 ID="Text2">
</TD>
<%	end if	%>
	</tr>
		<tr>
  <td CLASS="LABEL" colspan ="1"><B>Greeting:</B><br>
  <%	if instr(CONNECT_STRING,"SED")> 0 THEN %>
  <SPAN ID="GreetingText" CLASS=LABEL TITLE = "<%=RsGREETING%>" > <%=RsGREETING%> </SPAN>
  <%ELSE %>

	       <input ScrnInput="TRUE" 
	           size="80" 
	           CLASS="LABEL" 
	           MAXLENGTH="225" 
	           TYPE="TEXT" 
	           NAME="TxtGreeting" 
	           VALUE="<%=RsGREETING%>" 
	           TABINDEX = 7
	           ONKEYPRESS="VBScript::Control_OnChange" 
	           ONCHANGE="VBScript::Control_OnChange" ID="Text5">
  <%END IF%>
	           
	   </td>

    </tr>
</table>
    </tr>
</table>

</form>


<% Else %>

<div style="margin-top:170px;margin-left:170px" CLASS="LABEL">
   UserID can have only one location
</div>


<% End If %>

</body>
</html>



