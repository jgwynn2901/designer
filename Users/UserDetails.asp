<!--#include file="..\lib\common.inc"-->
<!--#include file="..\lib\StatusRptinc.asp"-->
<!--#include file="..\lib\ControlData.inc"-->
<!--#include file="..\lib\RenderTextinc.asp"-->
<%	
'--------------------------------------------------------------------------------------------------------------------*/
' WORK REQUEST – MALB-0036 [TDS/SOW Document # if Exists : UserDetails.ASP]

'FNS DESIGNER/FNS CLAIMCAPTURE   
'Client			:	GBS
'Object			:	UserDetails.ASP.asp   
'Script Date: 08/02/2005		Script By: Narayan Ramachandran,Sutapa Majumdar
'Modified :08/03/2005

'Work Request/ILog #	:	MALB-0036
'Requirement		: Designer Code Change for Sentry Security	 

'*/
'---------------------------------------------------------------------------------------------------------------------->

    Response.Expires=0 
    Response.Buffer = true

    Dim UID ,ACTIVE
	dim getSession,checkSentry, checkSedgwick 'Added This Line as per Work Request MALB-0036
      getSession=Session("SecurityObj").m_ConnectionString
      checkSentry=instr(getSession,"SEN")'Added This Line as per Work Request MALB-0036
      checkSedgwick = instr(getSession,"SED") 'Added This Line as per Work Request TSOL-0105
      
    UID	= CStr(Request.QueryString("UID"))
	
    Dim RsName, RsPassword, RsSiteID, RsSiteName, RsActive, RsLName, RsFName
    Dim RsAddress1, RsAddress2, RsCity, RsState, RsZipCode, RsPhoneWork, RsPhoneExt
    Dim RsFaxNumber, RsEmail, RsCallerType, RsCallerDep, RsSupName, RsSupPhone, RsOldName
    Dim RsSupExt, RsActivStartDate, RsActivEndDate, RsPasswCreatDate
    
    dim RsOperatorID 'Added This Line as per Work Request MALB-0036
    
    Dim RsLastChangeDate, RsNewUser, FullOldName, strOldName
    Dim strNewFName, strNewLname, strNewFullName, RsPasswExpirDate
    Dim dispExpirPassword, dispPassCreatDate, dispLastChangedDate, sCreatDate, PassExpDate
    
    strOldName = ""
    RsOldName = ""

If UID <> "NEW" Then
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open CONNECT_STRING
    SQLST = "SELECT u.NAME, u.PASSWORD, u.SITE_ID, u.ACTIVE, u.LAST_NAME, S.NAME as SITE_NAME,"
    SQLST = SQLST & " u.FIRST_NAME, u.ADDRESS_LINE_1, u.ADDRESS_LINE_2,"
    SQLST = SQLST & " u.CITY, u.STATE, u.ZIP_CODE, u.PHONE_WORK, u.PHONE_WORK_EXTENSION,"
    SQLST = SQLST & " u.FAX_NUMBER, u.EMAIL_ADDRESS, u.CALLER_TYPE, u.CALLER_DEPARTMENT,"
    SQLST = SQLST & " u.SUPERVISOR_NAME, u.SUPERVISOR_PHONE_WORK, u.SUPERVISOR_PHONE_EXTENSION,"
    SQLST = SQLST & " TO_CHAR(u.ACTIVE_START_DATE, 'mm/dd/yyyy') AS ACTIVE_START_DATE,"
    SQLST = SQLST & " TO_CHAR(u.ACTIVE_END_DATE, 'mm/dd/yyyy') AS ACTIVE_END_DATE,"
    SQLST = SQLST & " TO_CHAR(u.PASSWORD_CREATION_DATE, 'mm/dd/yyyy') AS PASSWORD_CREATION_DATE,"
    SQLST = SQLST & " TO_CHAR(u.PASSWORD_EXPIRATION_DATE, 'mm/dd/yyyy') AS PASSWORD_EXPIRATION_DATE,"
    SQLST = SQLST & " TO_CHAR(u.LAST_CHANGE_DATE, 'mm/dd/yyyy') AS LAST_CHANGE_DATE,"
    SQLST = SQLST & " u.NEW_USER, u.REUSE"
     if cint(checkSentry)>0 then 'Added This Line as per Work Request MALB-0036
		SQLST = SQLST & " ,u.OPERATOR_ID"
    end if
    SQLST = SQLST & " FROM USERS_VIEW u, SITE s "
    SQLST = SQLST & " WHERE u.SITE_ID = s.SITE_ID "
    SQLST = SQLST & "  and USER_ID =" & UID
    
      Set RS = Conn.Execute(SQLST)
      
    If Not RS.EOF Then
        RsName = "" & ReplaceQuotesInText(RS("NAME"))
        RsPassword = "" & ReplaceQuotesInText(RS("PASSWORD"))
        RsSiteID = "" & RS("SITE_ID")
        RsSiteName = "" & RS("SITE_NAME")
        RsActive = "" & RS("ACTIVE")
        RsLName = "" & ReplaceQuotesInText(RS("LAST_NAME"))
        RsFName = "" & ReplaceQuotesInText(RS("FIRST_NAME"))
        RsAddress1 = "" & ReplaceQuotesInText(RS("ADDRESS_LINE_1"))
        RsAddress2 = "" & ReplaceQuotesInText(RS("ADDRESS_LINE_2"))
        RsCity = "" & RS("CITY")
        RsState = "" & RS("STATE")
        RsZipCode = "" & RS("ZIP_CODE")
        RsPhoneWork = "" & RS("PHONE_WORK")
        RsPhoneExt = "" & RS("PHONE_WORK_EXTENSION")
        RsFaxNumber = "" & RS("FAX_NUMBER")
        RsEmail = "" & RS("EMAIL_ADDRESS")
        RsCallerType = "" & RS("CALLER_TYPE")
        RsCallerDep = "" & ReplaceQuotesInText(RS("CALLER_DEPARTMENT"))
        RsSupName = "" & ReplaceQuotesInText(RS("SUPERVISOR_NAME"))
        RsSupPhone = "" & RS("SUPERVISOR_PHONE_WORK")
        RsSupExt = "" & RS("SUPERVISOR_PHONE_EXTENSION")
        RsActivStartDate = "" & RS("ACTIVE_START_DATE")
        RsActivEndDate = "" & RS("ACTIVE_END_DATE")
        RsPasswCreatDate = "" & RS("PASSWORD_CREATION_DATE")
        RsPasswExpirDate = "" & RS("PASSWORD_EXPIRATION_DATE")
        RsLastChangeDate = "" & RS("LAST_CHANGE_DATE")
        RsNewUser = "" & RS("NEW_USER")
        RsReuse = "" & RS("REUSE")
        
       
        if cint(checkSentry)>0 then  'Added This Line as per Work Request MALB-0036
				RsOperatorID="" & RS("OPERATOR_ID")
		end if
	 
        'save full name for comp

        RsOldName = Trim(Left(RsFName, 1)) & Trim(Left(RsLName, 6))
    End If
    RS.Close
    Set RS = Nothing
    Set Conn = Nothing
End if
%>
<html>
	<head>
		<meta name="VI60_defaultClientScript" content="VBScript">
		<title>User Details</title>
		<link rel="stylesheet" type="text/css" href="..\FNSDESIGN.css">
			<script>
var g_StatusInfoAvailable = false;
			</script>
			<script LANGUAGE="JavaScript" FOR="window" EVENT="onload">
    <%
    Dim dDate, dDatePlus3, dEnd

    If CStr(Request.QueryString("MODE")) = "RO" Then
    %>
    SetScreenFieldsReadOnly(true,"DISABLED");
    <%  End If %>
    if (document.all.DataFrame != null)
    document.all.DataFrame.style.height = document.body.clientHeight - 150;
    if (document.all.fldSet != null)
    document.all.fldSet.style.height = document.body.clientHeight - 150;
    if (document.all.SPANDATA != null)
    document.all.SPANDATA.innerText = "";
    <%
    If Request.QueryString("UID") <> "NEW" Then
        %>
        if ("<%=RsNewUser%>" == "Y")
			document.all.newUser[0].checked = true;
		else
			document.all.newUser[1].checked = true;
        if ("<%=RsActive%>" == "Y")
			document.all.activeUser[0].checked = true;
		else
			{
			document.all.activeUser[1].checked = true;
			//document.all.rec0.style.visibility = "visible";
			document.all.recycle[1].checked = true;
			}
        <%
    Else
        dDatePlus3 = FormatDateTime(DateAdd("m", 3, Now), 2)
        dDate = FormatDateTime(Now, 2)
        dEnd = FormatDateTime(DateAdd("yyyy", 1, dDate), 2)
        %>
        document.all.TxtExpirDate.value = "<%=dDatePlus3%>";
        document.all.TxtActStartDate.value = "<%=dDate%>";
        document.all.TxtActEndDate.value = "<%=dEnd%>";
        document.all.newUser[1].disabled = true;
        <%
    End If
    %>
			</script>
			
	<script language="javascript">
	function username_onBlur(userName)
    {
		// Username allowed in upper/lower case only for SED
		if (document.getElementById('TxtDBNAME').value.toUpperCase().match('SED'))
		{
			// If username is neither upper or lowercase, alert user
			if (!(userName.value == userName.value.toUpperCase() ||
				userName.value == userName.value.toLowerCase()))
			{
				alert('Usernames must be either \'UPPER CASE\' or \'lower case\'.');
				userName.focus();
			}
		}
    }
    var isJqueryError; // isAjax(JQuery) Script Error
    isJqueryError = false;		
    function generateUserName()
    { 		
	    var nTryCnt,cUsrName,txtFName,txtLName;
	    txtFName = document.getElementById("TxtFName").value;
	    txtLName = document.getElementById("TxtLName").value;
    	
	    if(txtLName == "" || txtFName == "")
	    {
			    alert("Please enter the first and last name");
	    }
	    else
	    {			
		    nTryCnt = 1;
		    do
		    {
			    cUsrName = ConstrUserName(txtLName,txtFName);
			    if (parent.frames("userLookup").doesUserExist(cUsrName)) 
			    {
				    //Try Another One
				    nTryCnt +=1; 
			    }
			    else
			    {
				    document.getElementById("TxtUserName").value = cUsrName;				
				    return;
			    }
		    }while (nTryCnt<=3)		
		    if(!isJqueryError)
			    alert("Failed to generate a unique Username. Please try manually.");
	    }	////ending else			   
    }
	
	function ResetPassword()
	{
	 document.all.TxtPass.value = parent.frames("userLookup").getDefaultPassword();
	 Control_OnChange();
	}
	</script>
	
			<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
sub VBStop
stop
end sub

Sub PostTo(strURL)
    FrmDetails.Action = strURL
    FrmDetails.method = "GET"
    FrmDetails.Target = "_parent"
    FrmDetails.submit
End Sub

Sub UpdateUID(inUID)
 document.All.UID.Value = inUID
 document.All.spanUID.innerText = inUID

 <%  'AHSID will only be available when launched from the tree
 If CStr(Request.QueryString("AHSID")) <> "" Then %>
 sResult = "(" & document.All.UID.Value
 sResult = sResult & " ," & "<%=CStr(Request.QueryString("AHSID"))%>" & ")"
 document.All.TxtSaveData.Value = sResult
 document.All.TxtAction.Value = "INSERT"
 FrmDetails.Action = "UserAccountsSave.asp"
 FrmDetails.method = "POST"
 FrmDetails.Target = "hiddenPage"
 FrmDetails.submit
 <%  End If %>
End Sub

sub enableTabs
parent.parent.enableTab("Details")
parent.parent.enableTab("Groups")
parent.parent.enableTab("Permissions")
parent.parent.enableTab("Accounts")
parent.parent.enableTab("Locations")
end sub

sub lockUserName
document.All.TxtUserName.readOnly = true
document.all.genName.style.visibility = "hidden"
end sub

Sub UpdateStatus(inStatus)
 document.All.SpanStatus.innerHTML = inStatus
End Sub

Sub SetStatusInfoAvailableFlag(bAvailable)
 g_StatusInfoAvailable = bAvailable
 If bAvailable = True Then
    document.All.StatusRpt.Style.cursor = "HAND"
 Else
    document.All.StatusRpt.Style.cursor = "DEFAULT"
 End If
End Sub

Function GetUID()
 GetUID = document.All.UID.Value
End Function

Function GetUIDName()
    GetUIDName = document.All.TxtName.Value
End Function

Function CheckDirty()
    If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
        CheckDirty = True
    Else
        CheckDirty = False
    End If
End Function

Sub SetDirty()
    document.body.setAttribute "ScreenDirty", "YES"
End Sub

Sub ClearDirty()
    document.body.setAttribute "ScreenDirty", "NO"
End Sub

Function ValidateScreenData()

    Dim errStr, errEndtr
    Dim obj
    errStr = ""
    errEndtr = ""

    If document.All.UID.Value = "NEW" Then
        '   is name unique?
        Set obj = document.All.TxtUserName
        If CBool(Parent.frames("userLookup").doesUserExist(obj.Value)) Then
            MsgBox "Name '" & obj.Value & "' already exists.", 16, "FNSDesigner"
            ValidateScreenData = False
            obj.focus()
            Exit Function
        End If
    End If

    '====FIRST NAME ECAH NOT BE NULL
    If document.All.TxtFName.Value = "" Then
        MsgBox "First Name  is a required field.", 0, "FNSNetDesigner"
        Set obj = document.All.TxtFName
        ValidateScreenData = False
        obj.focus()
        Exit Function
    End If
    '==== LAST  NAME ECAH NOT BE NULL
    If document.All.TxtLName.Value = "" Then
        MsgBox "Last Name is a required field.", 0, "FNSNetDesigner"
        Set obj = document.All.TxtLName
        ValidateScreenData = False
        obj.focus()
        Exit Function
    End If

    '==== sITE CAN NOT BE NULL
    If document.All.TxtSite.Value = "" Then
        MsgBox "Site is a required field.", 0, "FNSNetDesigner"
        Set obj = document.All.TxtSite
        ValidateScreenData = False
        obj.focus()
        Exit Function
    End If

    '====VALIDATION STRT DAYE'
    If errStr = "" Then
        If document.All.TxtActStartDate.Value = "" Then
            errStr = errStr & "Account Start Date is a required field." ',0,"FNSNetDesigner"
        elseIf Not isDate(document.All.TxtActStartDate.Value) Then
            errStr = errStr & "Account Start Date has an incorrect format. Format as MM/DD/YYYY" & vbCrLf
        End If

        If errStr = "" Then
            ValidateScreenData = True
        Else
            Set obj = document.All.TxtActStartDate
            MsgBox errStr, 0, "FNSNetDesigner"
            ValidateScreenData = False
            obj.focus()
            Exit Function
        End If
    End If 

    '====VALIDATION AND DATE
    If errEndtr = "" Then
        If document.All.TxtActEndDate.Value = "" Then
            errEndtr = errEndtr & "Account End Date is a required field." ',0,"FNSNetDesigne"
        elseIf Not isDate(document.All.TxtActEndDate.Value) Then
            errEndtr = errEndtr & "Account End Date has an incorrect format. Format as MM/DD/YYYY" & vbCrLf
        elseif DateDiff("d", cDate(document.All.TxtActStartDate.Value), cDate(document.All.TxtActEndDate.Value)) < 0 then
			errEndtr = errEndtr & "Account End Date has to be a later date than Account Start Date." & vbCrLf        
        end If
        If errEndtr = "" Then
            ValidateScreenData = True
        Else
            Set obj = document.All.TxtActEndDate
            MsgBox errEndtr, 0, "FNSNetDesigner"
            ValidateScreenData = False
            obj.focus()
            Exit Function
        End If
    End If 

    '*** Password Expiration Date  TxtExpirDate
    If errEndtr = "" Then  'if errEndtr= "" then
        If document.All.TxtExpirDate.Value = "" Then
            errEndtr = errEndtr & "Password Expiration Date is a required field." ',0,"FNSNetDesigne"
        elseIf Not isDate(document.All.TxtExpirDate.Value) Then
            errEndtr = errEndtr & "Password Expiration Date has an incorrect format. Format as MM/DD/YYYY" & vbCrLf
        End If
        If errEndtr = "" Then
            ValidateScreenData = True
        Else
            Set obj = document.All.TxtExpirDate
            MsgBox errEndtr, 0, "FNSNetDesigner"
            ValidateScreenData = False
            obj.focus()
            Exit Function
        End If
    End If 'if errEndtr= "" then

	If document.All.activeUser(1).Checked Then
		If document.All.recycle(0).Checked Then
			if msgbox( "Recycling the account will delete all the related account Information." & chr(10) & chr(13) & "Do you want to proceed?", 36, "Recycle User Account") = 6 then
				ValidateScreenData = True
			else
				ValidateScreenData = false
			end if
		else
			ValidateScreenData = True
		end if
	else
		ValidateScreenData = True
	end if
	
	'For Sentry, Operator ID is a required field 
	'Added This Line as per Work Request JMCA - 0098
	<% if cint(checkSentry)>0 then %>
	
		if document.All.TxtOperatorID.Value = "" then
			msgbox "Operator ID is a required field", 0, "FNSnet Designer"
			ValidateScreenData = false
		else
		ValidateScreenData = True
		end if	
	<%end if %>
	 	
End Function

Function GetSelectedGID()
    GetSelectedGID = document.frames("DataFrame").GetSelectedGID
End Function

Function InEditMode()
    InEditMode = True

    If CStr(document.body.getAttribute("ScreenMode")) = "RO" Then
        MsgBox "This screen is read only.", 0, "FNSNetDesigner"
        InEditMode = False
    End If

End Function

Function ConstrUserName(cLastName, cFirstName) '--- CREATE USER NAME
    Dim sCreatedUserName
    Dim sStart, sRandomNumber, sTotalNumber, nTotalUserName
    Dim iCount, nDigits

    'CRETE user Nme
    cFirstName = Trim(Left(cFirstName, 1))
    if len(cLastName) > 7 then
		cLastName = Trim(Left(cLastName, 7))
	else
		cLastName = Trim(cLastName)
	end if
    nTotalUserName = 1 + Len(cLastName)
    'amount of digit add to user name
    ' user name can not les then 6 digits
    If nTotalUserName < 6 Then
        iCount = 6 - nTotalUserName
        iCount = iCount + 2
    Else
        iCount = 2
    End If
    nDigits = iCount
    'USE THE TIME didit
    sStart = Timer
    sTotalNumber = CStr(Replace(sStart, ".", ""))
    sRandomName = Right(sTotalNumber, nDigits)

    'full user name
    sCreatedUserName = Trim(cFirstName & cLastName & sRandomName)
    ConstrUserName = sCreatedUserName
End Function

Function ExeSave()
    If Not InEditMode Then
        ExeSave = False
        Exit Function
    End If
    If document.All.UID.Value = "" Then
        ExeSave = False
        Exit Function
    End If
    bRet = False
    If CStr(document.body.getAttribute("ScreenDirty")) = "YES" Then
        If ValidateScreenData = False Then
            ExeSave = False
            Exit Function
        End If
        If document.All.UID.Value = "NEW" Then
            document.All.TxtAction.Value = "INSERT"
        Else
            document.All.TxtAction.Value = "UPDATE"
        End If
        <%
        If UID = "NEW" Then
        %>
			document.all.TxtPass.value = parent.frames("userLookup").getDefaultPassword()
		<%
		end if
		%>
       sResult = sResult & "inUserId" & Chr(129) & document.All.UID.Value & Chr(128)
        sResult = sResult & "inSiteId" & Chr(129) & document.All.TxtSite.Value & Chr(128)
        If document.All.activeUser(0).Checked Then
            sResult = sResult & "inActive" & Chr(129) & "Y" & Chr(128)
            sResult = sResult & "inReuse" & Chr(129) & "N" & Chr(128)
        Else
            sResult = sResult & "inActive" & Chr(129) & "N" & Chr(128)
			If document.All.recycle(0).Checked Then
				sResult = sResult & "inReuse" & Chr(129) & "Y" & Chr(128)
				document.All.TxtReuse.value = "Y"
			Else
				sResult = sResult & "inReuse" & Chr(129) & "N" & Chr(128)
			End If
        End If
        sResult = sResult & "inLastName" & Chr(129) & document.All.TxtLName.Value &  Chr(128)
        sResult = sResult & "inFirstName" & Chr(129) & document.All.TxtFName.Value & Chr(128)

        'If document.All.UID.Value = "NEW" Then
            if instr(document.All.TxtDBNAME.value,"SED")>0 then
			    sResult = sResult & "inName" & Chr(129) & document.All.TxtUserName.Value & Chr(128)
		    else			
				sResult = sResult & "inName" & Chr(129) & UCase(document.All.TxtUserName.Value) & Chr(128)
			End If
		'end if 	
        sResult = sResult & "inPassword" & Chr(129) & document.All.TxtPass.Value & Chr(128)
        sResult = sResult & "inAddressLine1" & Chr(129) & document.All.TxtAddress1.Value & Chr(128)
        sResult = sResult & "inAddressLine2" & Chr(129) & document.All.TxtAddress2.Value & Chr(128)
        sResult = sResult & "inCity" & Chr(129) & document.All.TxtCity.Value & Chr(128)
        sResult = sResult & "inState" & Chr(129) & document.All.TxtState.Value & Chr(128)
        sResult = sResult & "inZipCode" & Chr(129) & document.All.TxtZip.Value & Chr(128)
        sResult = sResult & "inPhoneWork" & Chr(129) & document.All.TxtPhonWork.Value & Chr(128)
        sResult = sResult & "inPhoneWorkExtension" & Chr(129) & document.All.TxtPhonWorkExt.Value & Chr(128)
        sResult = sResult & "inFaxNumber" & Chr(129) & document.All.TxtFaxNumber.Value & Chr(128)
        sResult = sResult & "inEmailAddress" & Chr(129) & document.All.TxtEmailAddress.Value & Chr(128)
        sResult = sResult & "inCallerType" & Chr(129) & document.All.TxtCallerType.Value & Chr(128)
        sResult = sResult & "inCallerDepartment" & Chr(129) & document.All.TxtCallerDepartment.Value & Chr(128)
        sResult = sResult & "inSupervisorName" & Chr(129) & document.All.TxtSupName.Value & Chr(128)
        sResult = sResult & "inSupervisorPhoneWork" & Chr(129) & document.All.TxtSupPhonExt.Value & Chr(128)
        sResult = sResult & "inSupervisorPhoneExtension" & Chr(129) & document.All.TxtSupPhonExt.Value & Chr(128)
        sResult = sResult & "inActiveStartDate" & Chr(129) & document.All.TxtActStartDate.Value & Chr(128)
        sResult = sResult & "inActiveEndDate" & Chr(129) & document.All.TxtActEndDate.Value & Chr(128)        
		sResult = sResult & "inPasswordExpirationDate" & Chr(129) & document.All.TxtExpirDate.Value & Chr(128) 
		' added by sajjad
		if document.All.newUser(0).Checked Then
            sResult = sResult & "inNewUser" & Chr(129) & "Y" & Chr(128) 
        else
            sResult = sResult & "inNewUser" & Chr(129) & "N" & Chr(128)
        end if
		' end
		if ucase(trim(document.All.TxtSite.options(document.All.TxtSite.selectedIndex).innertext)) = "INTERNET" then
			sResult = sResult & "inInternetUser" & Chr(129) & "Y" & Chr(128)
		else
			sResult = sResult & "inInternetUser" & Chr(129) & "N" & Chr(128)
		end if

		document.All.TxtSaveData.Value = sResult
        FrmDetails.Action = "UserSave.asp"
        FrmDetails.method = "POST"
        FrmDetails.Target = "hiddenPage"
        FrmDetails.submit
        bRet = True
    Else
        SpanStatus.innerHTML = "Nothing to Save"
    End If
    ExeSave = bRet

End Function

Sub Control_OnChange()
    If document.body.getAttribute("ScreenMode") <> "RO" Then
        document.body.setAttribute "ScreenDirty", "YES"
    End If
End Sub

sub activeAcc_OnChange
	Control_OnChange
	If not document.All.activeUser(0).Checked Then
		'document.All.rec0.style.visibility = "visible"
	else
		document.All.rec0.style.visibility = "hidden"
	end if
end sub

Sub SetScreenFieldsReadOnly(bReadOnly, strNewClass)

    For iCount = 0 To document.All.length - 1
        If document.All(iCount).getAttribute("ScrnInput") = "TRUE" Then
            document.All(iCount).ReadOnly = bReadOnly
            document.All(iCount).className = strNewClass
        ElseIf document.All(iCount).getAttribute("ScrnBtn") = "TRUE" Then
            document.All(iCount).disabled = bReadOnly
        End If
    Next
End Sub

Sub StatusRpt_OnClick()
    If g_StatusInfoAvailable = True Then
        lret = window.showModalDialog("..\StatusRpt\StatusRpt.asp", Null, "dialogWidth=580px; dialogHeight=380px; center=yes")
    Else
        MsgBox "No other detail status reported.", 0, "FNSNetDesigner"
    End If
End Sub

</script>
		<!--#include file="..\lib\UserBtnControl.inc"-->
	</head>
	<body topmargin="0" leftmargin="0" rightmargin="0" BGCOLOR="<%=BODYBGCOLOR%>" ScreenDirty="NO" ScreenMode="<%=Request.QueryString("MODE")%>">
		<table WIDTH="100%" CELLPADDING="0" CELLSPACING="0" ID="Table1">
			<tr>
				<td CLASS="GrpLabel" WIDTH="134" HEIGHT="10">
					<nobr>&nbsp;» Details</td>
				<td HEIGHT="5" ALIGN="LEFT">
					<table CELLPADDING="0" CELLSPACING="0" HEIGHT="100%" ID="Table2">
						<tr>
							<td CLASS="GrpLabelDrk" WIDTH="3" HEIGHT="8" VALIGN="BOTTOM" ALIGN="LEFT"></td>
							<td WIDTH="300" HEIGHT="8"></td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td CLASS="GrpLabelLine" colspan="2" HEIGHT="1"></td>
			</tr>
			<tr>
				<td colspan="2" HEIGHT="1"></td>
			</tr>
		</table>
		<form Name="FrmDetails" METHOD="POST" ACTION="UserSave.asp" TARGET="hiddenPage" ID="Form1">
			<input TYPE="HIDDEN" NAME="TxtSaveData" ID="Hidden1"> <input TYPE="HIDDEN" NAME="TxtAction" ID="Hidden2">
			<input TYPE="hidden" NAME="TxtReuse" value="N" ID="Hidden5"> <input TYPE="hidden" NAME="TxtDBNAME" value="<%=Session("SecurityObj").m_ConnectionString%>" ID="Hidden3">
			<% 'need to maintain these values in order to post back to the search tab %>
			<input type="hidden" name="SearchUID" value="<%=Request.QueryString("SearchUID")%>" ID="Hidden4">
			<input type="hidden" name="SearchName" value="<%=Request.QueryString("SearchName")%>" ID="Hidden6">
			<input type="hidden" name="SearchSite" value="<%=Request.QueryString("SearchSite")%>" ID="Hidden7">
			<input type="hidden" name="SearchType" value="<%=Request.QueryString("SearchType")%>" ID="Hidden8">
			<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" ID="Hidden9">
			<input type="hidden" NAME="UID" value="<%=Request.QueryString("UID")%>" ID="Hidden10">
			<input type="hidden" NAME="MODE" value="<%=Request.QueryString("MODE")%>" ID="Hidden11">
			<%	
Function GetActive()
    If (RsActive = "Y" Or RsActive = "") Then
        GetActive = "<OPTION VALUE='Y' SELECTED>Y</OPTION><OPTION VALUE='N'>N</OPTION>"
    Else
        GetActive = "<OPTION VALUE='Y'>Y</OPTION><OPTION VALUE='N' SELECTED>N</OPTION>"
    End If
End Function
    
Function GetNew()
    If (RsNewUser = "Y" Or RsNewUser = "") Then
        GetNew = "<OPTION VALUE='Y' SELECTED>Y</OPTION><OPTION VALUE='N'>N</OPTION>"
    Else
        GetNew = "<OPTION VALUE='Y'>Y</OPTION><OPTION VALUE='N' SELECTED>N</OPTION>"
    End If

End Function
%>
			<table class="Label" style="{position:absolute;top:20;}" ID="Table3" cellpadding="1" cellspacing="1">
				<tr>
					<td VALIGN="middle" WIDTH="5">
						<img ID="StatusRpt" SRC="..\images\StatusRpt.gif" width="16" height="16" ALT="View Status Report">
					</td>
					<td width="429">
						:<SPAN ID="SpanStatus" STYLE="COLOR:#006699" CLASS="LABEL"><font size="1">Ready</font></SPAN>
					</td>
					<td align="right" width="329">
						<%=FormatDateTime(now, 1)%>
					</td>
				</tr>
			</table>
			<br>
			<table class="LABEL" BORDER="0" height="116" width="789" ID="Table4" cellspacing="0" cellpadding="1">
				<tr>
					<td height="18" width="152" colspan="2"><font size="1">User ID:&nbsp;</font><span id="spanUID"><font size="1"><%=Request.QueryString("UID")%></font></span></td>
					<td align="right" height="18" width="202">
					</td>
					<td colspan="4" align="right" bordercolorlight="#008000" bordercolor="#008000" style="border-style: inset; border-width: 1"
						height="18" width="343">
						<div align="center"><font size="1">« Fields in <b>bold</b> are mandatory »</font></div>
					</td>
					<td width="72"></td>
				</tr>
				<tr>
					<td colspan="8" height="12" width="783">
						<p></p>
					</td>
				</tr>
				<tr>
					<td CLASS="LABEL" width="135"><b><font size="1">User Name:</font></b><br>
						<input ScrnInput="TRUE" 
<%
if Request.QueryString("UID") <> "NEW" then
%>
			   'READONLY
<%
end if
%>			   
	           size="22" 
	           CLASS="LABEL" 
	           MAXLENGTH="40" 
	           TYPE="TEXT" 
	           NAME="TxtUserName" 
	           id="TxtUserName" 
	           VALUE="<%=RsName%>"
	           TABINDEX =1
	           ONCHANGE="VBScript::Control_OnChange" 
	           onblur="username_onBlur(this);"
	           ALT="Type the Username or click the 'Generate' button " ID="Text1">
					</td>
					<td CLASS="LABEL" width="13"><br>
						<%
if Request.QueryString("UID") = "NEW" then
%>
						<img border="0" src="..\images\BD21298_.gif" align="bottom" width="13" height="13">
						<%
end if
%>
					</td>
					<td width="202">
						<%
if Request.QueryString("UID") = "NEW" then
%>
						<font size="1">
							<br>
							<input name="genName" type="button" id="genName" value="Generate" style="font-family: Verdana; font-size: 8pt" onclick="generateUserName();" /></font>
						<%
end if
%>
					</td>
					<td CLASS="LABEL" width="191" colspan="2">
						<font size="1">Password</font> (System generated):
						<br>
						<%
                if Request.QueryString("UID") <> "NEW" then
					'cType="password"
				else
					cType=""
				end if
                %>
						<input 
	           size="22" 
	           class="LABEL" 
	           maxlength="40" 
	           name="TxtPass" 
	           value="<%=RsPassword%>"
	           tabindex =2
	           onChange="VBScript::Control_OnChange" ID="Text2" type="<%=cType%>"> &nbsp;
			   <% if Request.QueryString("UID") <> "NEW" then %>
					<input name="ResetPass" type="button" id="ResetPass" value="Reset" style="font-family: Verdana; font-size: 8pt" onclick="ResetPassword();" /></td>
			   <% end if %>			   
					<td CLASS="LABEL" colspan="3" width="226"><font size="1"><b>Pwd</b></font><b><font size="1">
								Expiration Date:</font><br>
							<input ScrnInput="TRUE" 
	           size="20" 
	           CLASS="LABEL" 
	           MAXLENGTH="40"
	           NAME="TxtExpirDate" 
	           VALUE="<%=RsPasswExpirDate%>" 
	           TABINDEX =27
	           ONCHANGE="VBScript::Control_OnChange" ID="Text3" style="text-align: right"></b></td>
				</tr>
				<tr>
					<td colspan="6" width="593">
						<br>
					</td>
				</tr>
				<tr>
					<td CLASS="LABEL" width="150" rowspan="2" height="40" colspan="2"><b><font size="1">First 
								Name:</font></b><br>
						<input ScrnInput="TRUE" 
	           size="22" 
	           CLASS="LABEL" 
	           MAXLENGTH="40" 
	           NAME="TxtFName"
	           id="TxtFName" 
	           VALUE="<%=RsFName%>"
	           TABINDEX =3
	           ONCHANGE="VBScript::Control_OnChange" ID="Text4">
					</td>
					<td CLASS="LABEL" width="182" rowspan="2">
						<b><font size="1">Last Name:</font></b><br>
						<input scrninput="TRUE" 
	            size="22" 
	            class="LABEL" 
	            maxlength="40" 
	            name="TxtLName" 
	            id="TxtLName" 
	            value="<%=RsLName%>"
	            tabindex = 4
	            onChange="VBScript::Control_OnChange" ID="Text5">
					</td>
					<td CLASS="LABEL" width="66" rowspan="2">
						<strong><font size="1">Site:</font></strong><br>
						<select scrnbtn="TRUE" name="TxtSite" class="LABEL" tabindex="5" onChange="VBScript::Control_OnChange"
							ID="Select1">
							<%=GetControlDataHTML("SITE","SITE_ID","NAME",RsSiteName,true)%>
						</select>
					</td>
					<td CLASS="LABEL" align="center">
						<b><font size="1">New User</font></b><br>
						<input name="newUser" type="radio" CHECKED value="Y" ID="Radio1" onChange="VBScript::Control_OnChange"><font size="1">Yes</font></td>
					<!-- Operator ID-->
					<% if cint(checkSentry)>0 then'Added This Line as per Work Request MALB-0036 %>
					<td CLASS="LABEL" width="150" rowspan="2" height="40" colspan="2"><b><font size="1">Operator 
								ID:</font></b><br>
						<input ScrnInput="TRUE" 
	           size="22" 
	           CLASS="LABEL" 
	           MAXLENGTH="3" 
	           NAME="TxtOperatorID" 
	           VALUE="<%=RsOperatorID%>"
	           TABINDEX =5
	           ONCHANGE="VBScript::Control_OnChange" ID="Text21">
					</td>
					<% end if %>
				</tr>
				<tr>
					<td CLASS="LABEL" width="183" align="center">
						<input name="newUser" type="radio" value="N" ID="Radio3" onChange="VBScript::Control_OnChange"><font size="1">
							No</font>
					</td>
				</tr>
			</table>
			<input type="hidden" NAME="TxtOldName" value="" ID="Hidden12"> <input type="hidden" NAME="TxtName" value="" ID="Hidden13">
			<input type="hidden" NAME="TxtPassword" value="" ID="Hidden14"> <input type="hidden" NAME="saveActiveFlage" value="" ID="Hidden15">
			<br>
			<table width="646" ID="Table5" cellspacing="0" cellpadding="1">
				<tr>
					<td CLASS="LABEL" width="158"><font size="1">Address 1:</font><br>
						<input ScrnInput="TRUE" 
	           size="25" 
	           CLASS="LABEL" 
	           MAXLENGTH="80" 
	           TYPE="TEXT" 
	           NAME="TxtAddress1" 
	           VALUE="<%=RsAddress1%>"
	           TABINDEX = 9
	           ONCHANGE="VBScript::Control_OnChange" ID="Text6">
					</td>
					<td CLASS="LABEL" width="158"><font size="1">Address 2:</font><br>
						<input ScrnInput="TRUE" 
	            size="25" 
	            CLASS="LABEL" 
	            MAXLENGTH="80" 
	            NAME="TxtAddress2" 
	            VALUE="<%=RsAddress2%>"
	            TABINDEX = 10
	            ONCHANGE="VBScript::Control_OnChange" ID="Text7">
					</td>
					<td CLASS="LABEL" width="130"><font size="1">City:</font><br>
						<input ScrnInput="TRUE" 
	            size="20" 
	            CLASS="LABEL" 
	            MAXLENGTH="40" 
	            TYPE="TEXT" 
	            NAME="TxtCity" 
	            VALUE="<%=RsCity%>"
	            TABINDEX = 11
	            ONCHANGE="VBScript::Control_OnChange" ID="Text8">
					</td>
					<td CLASS="LABEL" width="80"><font size="1">Zip:</font><br>
						<input ScrnInput="TRUE" 
	              size="10" 
	              CLASS="LABEL" 
	              MAXLENGTH="20" 
	              TYPE="TEXT"
	              NAME="TxtZip" 
	              VALUE="<%=RsZipCode%>" 
	              TABINDEX = 12
	              ONCHANGE="VBScript::Control_OnChange" ID="Text9">
					</td>
					<td CLASS="LABEL" width="49"><font size="1">State:</font><br>
						<SELECT tabindex="13" NAME="TxtState" CLASS="LABEL" ONCHANGE="VBScript::Control_OnChange"
							ID="Select2">
							<OPTION VALUE="<%=RsState%>" SELECTED><%=RsState%></OPTION>
							<!--#include file="..\lib\states.asp"-->
						</SELECT></td>
				</tr>
			</table>
			<table width="722" ID="Table6" cellspacing="0" cellpadding="1">
				<tr>
					<td CLASS="LABEL" width="96"><font size="1">Phone Work:</font><br>
						<input ScrnInput="TRUE" 
	           size="15" 
	           CLASS="LABEL" 
	           MAXLENGTH="14" 
	           TYPE="TEXT" 
	           NAME="TxtPhonWork" 
	           VALUE="<%=RsPhoneWork%>"
	           TABINDEX = 14
	           ONCHANGE="VBScript::Control_OnChange" ID="Text10">
					</td>
					<td CLASS="LABEL" width="86"><font size="1">Phone Ext:</font><br>
						<input ScrnInput="TRUE" 
	            size="10" 
	            CLASS="LABEL" 
	            MAXLENGTH="10" 
	            TYPE="TEXT" 
	            NAME="TxtPhonWorkExt" 
	            VALUE="<%=RsPhoneExt%>" 
	            TABINDEX = 15
	            ONCHANGE="VBScript::Control_OnChange" ID="Text11">
					</td>
					<td CLASS="LABEL" width="122"><font size="1">Fax Number:</font><br>
						<input ScrnInput="TRUE" 
	            size="15" 
	            CLASS="LABEL" 
	            MAXLENGTH="16" 
	            TYPE="TEXT" 
	            NAME="TxtFaxNumber" 
	            VALUE="<%=RsFaxNumber%>" 
	            TABINDEX = 16
	            ONCHANGE="VBScript::Control_OnChange" ID="Text12">
					</td>
					<td CLASS="LABEL" width="158"><font size="1">Email Address:</font><br>
						<input ScrnInput="TRUE" 
	              size="40" 
	              CLASS="LABEL" 
	              MAXLENGTH="50"
	              NAME="TxtEmailAddress" 
	              VALUE="<%=RsEmail%>" 
	              TABINDEX = 17
	              ONCHANGE="VBScript::Control_OnChange" ID="Text13">
					</td>
				</tr>
			</table>
			<table width="779" cellspacing="0" cellpadding="1" style="border-collapse: collapse" bordercolor="#111111"
				ID="Table7">
				<tr>
					<td CLASS="LABEL" rowspan="3" width="22"><font size="1">&nbsp;<i><u>Supervisor</u>:</i></font><i><br>
							&nbsp;</i></td>
					<td CLASS="LABEL" width="70"><font size="1">Name:</font><br>
						<input ScrnInput="TRUE" 
	           size="26" 
	           CLASS="LABEL" 
	           MAXLENGTH="40" 
	           TYPE="TEXT" 
	           NAME="TxtSupName" 
	           VALUE="<%=RsSupName%>" 
	           TABINDEX = 20
	           ONCHANGE="VBScript::Control_OnChange" ID="Text14"></td>
					<td CLASS="LABEL" width="130"><font size="1">Caller Type:</font><br>
						<input ScrnInput="TRUE" 
	              size="10" 
	              CLASS="LABEL" 
	              MAXLENGTH="40"      
	              TYPE="TEXT"          
	              NAME="TxtCallerType" 
	              VALUE="<%=RsCallerType %>"
	              TABINDEX = 18
	              ONCHANGE="VBScript::Control_OnChange" ID="Text15"></td>
					<td CLASS="LABEL" width="50" rowspan="2"><i><font size="1"><u>Account</u>:</font></i></td>
					<td CLASS="LABEL" width="94"><b><font size="1">Start Date:</font></b><br>
						<input ScrnInput="TRUE" 
	           size="20" 
	           CLASS="LABEL" 
	           MAXLENGTH="20" 
	           NAME="TxtActStartDate"                       
	           VALUE="<%=RsActivStartDate%>"
	           TABINDEX =23                                 
	           ONKEYPRESS="VBScript::Control_OnChange"       
	           ONCHANGE="VBScript::Control_OnChange" ID="Text16" style="text-align: right"></td>
					<td CLASS="LABEL" width="134" align="center">
						<b><font size="1">Active</font><br>
						</b><input name="activeUser" type="radio" CHECKED value="Y" ID="Radio5" onClick="VBScript::activeAcc_OnChange"><font size="1">Yes</font></td>
				</tr>
				<tr>
					<td CLASS="LABEL" width="70"><font size="1">Phone:</font><br>
						<input ScrnInput="TRUE" 
	           size="15" 
	           CLASS="LABEL" 
	           MAXLENGTH="21" 
	           TYPE="TEXT" 
	           NAME="TxtSupPhonWork" 
	           VALUE="<%=RsSupPhone%>" 
	           TABINDEX = 21
	           ONCHANGE="VBScript::Control_OnChange" ID="Text17"></td>
					<td CLASS="LABEL" width="130"><font size="1">Caller Department:</font><br>
						<input ScrnInput="TRUE" 
	              size="15" 
	              CLASS="LABEL" 
	              MAXLENGTH="40" 
	              TYPE="TEXT"
	              NAME="TxtCallerDepartment" 
	              VALUE="<%=RsCallerDep%>"
	              TABINDEX = 19
	              ONCHANGE="VBScript::Control_OnChange" ID="Text18"></td>
					<td CLASS="LABEL" width="94"><b><font size="1">End Date:</font></b><br>
						<input ScrnInput="TRUE" 
	           size="20" 
	           CLASS="LABEL" 
	           MAXLENGTH="20" 
	           NAME="TxtActEndDate" 
	           VALUE="<%=RsActivEndDate%>" 
	           TABINDEX =24
	           ONCHANGE="VBScript::Control_OnChange" ID="Text19" style="text-align: right"></td>
					<td CLASS="LABEL" width="134" align="center">
						<input name="activeUser" type="radio" value="N" ID="Radio6" onClick="VBScript::activeAcc_OnChange"><font size="1">
							No</font></td>
				</tr>
				<tr>
					<td CLASS="LABEL" width="469" colspan="5"><font size="1">Extension:</font><br>
						<input ScrnInput="TRUE" 
	           size="10" 
	           CLASS="LABEL" 
	           MAXLENGTH="10" 
	           TYPE="TEXT" 
	           NAME="TxtSupPhonExt" 
	           VALUE="<%=RsSupExt%>"
	           TABINDEX =22
	           ONCHANGE="VBScript::Control_OnChange" ID="Text20"></td>
				</tr>
				<tr id="rec0" style="visibility:hidden">
					<td CLASS="LABEL" colspan="3"><br>
					</td>
					<td CLASS="LABEL" colspan="3" align="center">
						<b><font size="1">Recycle Account?</font></b><input name="recycle" type="radio" CHECKED value="Y" ID="Radio7" onClick="VBScript::Control_OnChange"><font size="1">Yes</font><input name="recycle" type="radio" value="N" ID="Radio8" onClick="VBScript::Control_OnChange"><font size="1">
							No</font></td>
				</tr>
			</table>
		</form>
	</body>
</html>
