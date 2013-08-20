<% Response.Expires = 0 %>
<!--#include file="lib\StatusRptInc.asp"-->
<!--#include file="lib\common.inc"-->
<!--#include file="lib\commonError.inc"-->
<!--#include file="lib\AHSTree.inc"-->
<HTML>
<HEAD>
<SCRIPT>
<%
	On Error Resume Next
	ACTION = CStr(Request.Form("ACTION"))
	SAVEFILTER = CStr(Request.Form("SAVEFILTER"))
	SAVEFAVORITES= CStr(Request.Form("SAVEFAVORITES"))
	SAVEMAXRECORDS = Request.Form("SAVEMAXRECORDS")
	TREELEVELS = Request.Form("TREELEVELS")
	TREECOUNT = Request.Form("TREECOUNT")
	LCWidth = Request.Form("LCWidth")
	LCHeight = Request.Form("LCHeight")
	
	If SAVEFILTER = "" And SAVEFAVORITES = "" And SAVEMAXRECORDS = "" and LCWidth = Session("LayoutCtlWidth") and LCHeight = Session("LayoutCtlHeight") Then
		LogStatusGroupBegin
		LogStatusGroupEnd %>
		parent.frames('TOP').UpdateStatus('Nothing to save.');
		parent.frames('TOP').SetStatusInfoAvailableFlag(false);	
<%
	Else
		Set Conn = Server.CreateObject("ADODB.Connection")
		ConnectionString = CONNECT_STRING
		Conn.Open ConnectionString
	
		If ACTION = "SAVE" Then
			If SAVEFILTER = "on" Then
				strError = SaveFilters
			End If
			If SAVEFAVORITES = "on" Then
				If strError <> "" Then strError = strError & VBCRLF
				strError = strError & SaveFavoritesSettings
			End If
			if SAVEMAXRECORDS <> "" Then
				If strError <> "" Then strError = strError & VBCRLF
				strError = strError &  SaveMaxRecordsetting
				Session("USERMAXRECORDS") = SAVEMAXRECORDS
			End If
			if TREELEVELS <> "" Then
				If strError <> "" Then strError = strError & VBCRLF
				strError = strError &  SaveTreeLevel
				Session("USERTREELEVELS") = TREELEVELS
			end if
			if TREECOUNT <> "" Then
				If strError <> "" Then strError = strError & VBCRLF
				strError = strError &  SaveTreeCount
				Session("USERTREECOUNT") = TREECOUNT			
			end if
			if Session("LayoutCtlHeight") <> LCHeight then
				If strError <> "" Then 
					strError = strError & VBCRLF
				end if
				strError = strError &  SaveLCHeight
				Session("LayoutCtlHeight") = LCHeight 
			end if
			if Session("LayoutCtlWidth") <> LCWidth then
				If strError <> "" Then 
					strError = strError & VBCRLF
				end if
				strError = strError &  SaveLCWidth
				Session("LayoutCtlWidth") = LCWidth 
			end if
		Elseif ACTION = "CLEAR" Then
			If SAVEFILTER = "on" Then
				RemoveAllFilters "DESIGNER_AHSFILTER"
				strError = SaveFilters
			End If
			If SAVEFAVORITES = "on" Then
				RemoveAllFilters "DESIGNER_FAVORITES"
				Session("EXPANDLIST") = ""
				Session("AHLIST") = ""
				If strError <> "" Then strError = strError & VBCRLF
				strError =  strError & ClearFavoritesSettings
			End If
			if SAVEMAXRECORDS = "" Then
				If strError <> "" Then strError = strError & VBCRLF
				strError =  strError & ClearMaxRecords
				Session("USERMAXRECORDS") = 30
			End If
		End If
	
		If strError <> "" Then
			LogStatusGroupBegin
			LogStatus S_ERROR, strError, "SETTING", "", 0, ""
			LogStatusGroupEnd
		%>
			parent.frames("TOP").UpdateStatus("<SPAN STYLE='COLOR:#FF0000'>Error!</SPAN> Update unsuccessful, check status report.");
			parent.frames('TOP').SetStatusInfoAvailableFlag(true);

<%		Else
			LogStatusGroupBegin
			LogStatusGroupEnd %>
			parent.frames('TOP').UpdateStatus('Update successful.');
			parent.frames('TOP').SetStatusInfoAvailableFlag(false);		
<%		End If
		
		Conn.close
	End If


	
%>


<%	Function SaveFilters
		On Error Resume Next
		Set Obj = Session("AHSTreeFilter")
		
		strKeysArray = Obj.Keys
		strItemsArray = Obj.Items

		Conn.BeginTrans
		SQL = "DELETE FROM SETTING WHERE TYPE IN " &_
		"('DESIGNER_AHSFILTER','DESIGNER_RPFILTER','DESIGNER_CFFILTER','DESIGNER_POLICYFILTER','DESIGNER_BILLINGFILTER','DESIGNER_USERFILTER','DESIGNER_BAFILTER') " &_
		"AND USER_ID = " &  Session("SecurityObj").m_UserId
		Conn.Execute(SQL)
		

		strError = CheckADOErrors(Conn,"Setting DELETE")
		
		If strError <> "" Then
			Conn.RollbackTrans
			SaveFilters = strError
			Exit Function		
		End If
		
		For intLoop = 0 To Obj.Count -1
			strThisKey = strKeysArray(intLoop)
			parts = Split(strThisKey,":")
			strKeyID = parts(0)
			strType = parts(1)
			strName = parts(2)
			strValue = strItemsArray(intLoop)
			strValue = Replace(strValue,"'","''")

			if (strType = "DESIGNER_AHSFILTER") Or (strType = "DESIGNER_RPFILTER") Or _
			 (strType = "DESIGNER_CFFILTER") Or (strType = "DESIGNER_POLICYFILTER") Or _
			 (strType = "DESIGNER_BILLINGFILTER") Or (strType = "DESIGNER_USERFILTER") Or _
			 (strType = "DESIGNER_BAFILTER") Then 			
				SQL = "{call Designer.GetValidSeq('SETTING', 'SETTING_ID', {resultset 1, outResult})}"
				Set RSNextID = Conn.Execute(SQL)
			
				strError = CheckADOErrors(Conn,"Setting GetValidSeq")
			
				If strError <> "" Then
					Conn.RollbackTrans
					SaveFilters = strError
					Exit Function		
				End If

				newSettingID = RSNextID("outResult")
				RSNextID.close
		
				SQL = "INSERT INTO SETTING (SETTING_ID,USER_ID,NAME,VALUE,TYPE,VERSION,KEY_ID) VALUES("		
				SQL = SQL & newSettingID & ","
				SQL = SQL & Session("SecurityObj").m_UserId & "," 
				SQL = SQL & "'" & strName & "'," 
				SQL = SQL & "'" & strValue & "',"
				SQL = SQL & "'" & strType & "',"
				SQL = SQL & "1" & ","
				SQL = SQL & strKeyID & ")"
			
				Conn.Execute(SQL)

				strError = CheckADOErrors(Conn,"Setting INSERT")
			
				If strError <> "" Then
					Conn.RollbackTrans
					SaveFilters = strError
					Exit Function		
				End If

			End If
		Next
		
		Conn.CommitTrans
		SaveFilters = ""
	End Function
	
	
Function SaveFavoritesSettings

		On Error Resume Next
		Conn.BeginTrans
		SQL = ""
		SQL = SQL & "DELETE FROM SETTING WHERE TYPE LIKE 'DESIGNER_FAVORITES' AND USER_ID = " &  Session("SecurityObj").m_UserId
		Conn.Execute(SQL)
		strError = CheckADOErrors(Conn,"Setting DELETE")
		
		If strError <> "" Then
			Conn.RollbackTrans
			SaveFavoritesSettings = strError
			Exit Function		
		End If
		Conn.CommitTrans
	
			strKeyID = "null"
			strKeyID_Exp = "null"
			strType = "DESIGNER_FAVORITES"
		
			strName = "FAVORITES_AHSID"
			strName_Exp = "EXP_FAVORITES_AHSID"
			strValue = Session("AHLIST")
			strValue_exp = Session("EXPANDLIST")
			strValue = Replace(strValue,"'","''")
			strValueexp = Replace(strValue,"'","''")
			
						
			SQL = "{call Designer.GetValidSeq('SETTING', 'SETTING_ID', {resultset 1, outResult})}"
			Set RSNextID = Conn.Execute(SQL)
			Set RSNextID_Exp = Conn.Execute(SQL)
			
			strError = CheckADOErrors(Conn,"Setting GetValidSeq")
			
			If strError <> "" Then
				Conn.RollbackTrans
				SaveFavoritesSettings = strError
				Exit Function		
			End If

			newSettingID = RSNextID("outResult")
			newSettingID2 = RSNextID_Exp("outResult")
			
			RSNextID.close
			RSNextID_Exp.Close
			
			SQL = "INSERT INTO SETTING (SETTING_ID,USER_ID,NAME,VALUE,TYPE,VERSION,KEY_ID) VALUES("		
			SQL = SQL & newSettingID & ","
			SQL = SQL & Session("SecurityObj").m_UserId & "," 
			SQL = SQL & "'" & strName & "'," 
			SQL = SQL & "'" & strValue & "',"
			SQL = SQL & "'" & strType & "',"
			SQL = SQL & "1" & ","
			SQL = SQL & strKeyID & ")"
			Conn.Execute(SQL)
			strError = CheckADOErrors(Conn,"Setting INSERT")
			If strError <> "" Then
				Conn.RollbackTrans
				SaveFavoritesSettings = strError
				Exit Function		
			End If
			Conn.CommitTrans
			
			SQL = "INSERT INTO SETTING (SETTING_ID,USER_ID,NAME,VALUE,TYPE,VERSION,KEY_ID) VALUES("		
			SQL = SQL & newSettingID2 & ","
			SQL = SQL & Session("SecurityObj").m_UserId & "," 
			SQL = SQL & "'" & strName_Exp & "'," 
			SQL = SQL & "'" & strValue_exp & "',"
			SQL = SQL & "'" & strType & "',"
			SQL = SQL & "1" & ","
			SQL = SQL & strKeyID_exp & ")"

			Conn.Execute(SQL)
			
			strError = CheckADOErrors(Conn,"Setting INSERT")
			If strError <> "" Then
				Conn.RollbackTrans
				SaveFavoritesSettings = strError
				Exit Function		
			End If
			
		
		Conn.CommitTrans
		SaveFavoritesSettings = ""
	End Function
	
	
Function ClearFavoritesSettings
	On Error Resume Next
		
	Conn.BeginTrans
	SQL = ""
	SQL = SQL & "DELETE FROM SETTING WHERE TYPE LIKE 'DESIGNER_FAVORITES' AND USER_ID = " &  Session("SecurityObj").m_UserId
	Conn.Execute(SQL)
		

	strError = CheckADOErrors(Conn,"Setting DELETE")
		
	If strError <> "" Then
		Conn.RollbackTrans
		SaveFavoritesSettings = strError
		Exit Function		
	End If
	Conn.CommitTrans
	SaveFavoritesSettings = ""
End Function


Function SaveMaxRecordsetting()
On Error Resume Next
		Conn.BeginTrans
		SQL = ""
		SQL = SQL & "DELETE FROM SETTING WHERE TYPE LIKE 'DESIGNER_MAXRECORDS' AND USER_ID = " &  Session("SecurityObj").m_UserId
		Conn.Execute(SQL)
		strError = CheckADOErrors(Conn,"Setting DELETE")
		
		If strError <> "" Then
			Conn.RollbackTrans
			SaveMaxRecordsetting = strError
			Exit Function		
		End If
		Conn.CommitTrans
	
			strKeyID = "null"
			strKeyID_Exp = "null"
			strType = "DESIGNER_MAXRECORDS"
		
			strName = "DESIGNER_MAXRECORDS"
			strValue = SAVEMAXRECORDS
			strValue = Replace(strValue,"'","''")
								
			SQL = "{call Designer.GetValidSeq('SETTING', 'SETTING_ID', {resultset 1, outResult})}"
			Set RSNextID = Conn.Execute(SQL)
			
			strError = CheckADOErrors(Conn,"Setting GetValidSeq")
			
			If strError <> "" Then
				Conn.RollbackTrans
				SaveMaxRecords = strError
				Exit Function		
			End If

			newSettingID = RSNextID("outResult")
			
			RSNextID.close
	
			SQL = "INSERT INTO SETTING (SETTING_ID,USER_ID,NAME,VALUE,TYPE,VERSION,KEY_ID) VALUES("		
			SQL = SQL & newSettingID & ","
			SQL = SQL & Session("SecurityObj").m_UserId & "," 
			SQL = SQL & "'" & strName & "'," 
			SQL = SQL & "'" & strValue & "',"
			SQL = SQL & "'" & strType & "',"
			SQL = SQL & "1" & ","
			SQL = SQL & strKeyID_exp & ")"
			Conn.Execute(SQL)
			
			strError = CheckADOErrors(Conn,"Setting INSERT")
			If strError <> "" Then
				Conn.RollbackTrans
				SaveMaxRecordsetting = strError
				Exit Function		
			End If
			
		
		Conn.CommitTrans
		
		SaveMaxRecordsetting = ""

End Function

Function SaveLCWidth
dim cSQL, cError, oRS, cSettingID, cKeyID_Exp

On Error Resume Next
SaveLCWidth = ""
Conn.BeginTrans

cSQL = "DELETE FROM SETTING WHERE TYPE = 'LAYOUTCTLWIDTH' AND USER_ID = " &  Session("SecurityObj").m_UserId
Conn.Execute(cSQL)
cError = CheckADOErrors(Conn,"Setting DELETE")
		
If cError <> "" Then
	Conn.RollbackTrans
	SaveLCWidth = cError
	cError = ""
else
	Conn.CommitTrans
End If
cKeyID_Exp = "null"
								
cSQL = "{call Designer.GetValidSeq('SETTING', 'SETTING_ID', {resultset 1, outResult})}"
Set oRS = Conn.Execute(cSQL)
cError = CheckADOErrors(Conn,"Setting GetValidSeq")
			
If cError <> "" Then
	Conn.RollbackTrans
	SaveLCWidth = cError
	cError = ""
end If

cSettingID = oRS("outResult")
oRS.close
	
cSQL = "INSERT INTO SETTING (SETTING_ID,USER_ID,NAME,VALUE,TYPE,VERSION,KEY_ID) VALUES("		
			cSQL = cSQL & cSettingID & ","
			cSQL = cSQL & Session("SecurityObj").m_UserId & "," 
			cSQL = cSQL & "'LAYOUTCTLWIDTH'," 
			cSQL = cSQL & "'" & LCWidth & "',"
			cSQL = cSQL & "'LAYOUTCTLWIDTH',"
			cSQL = cSQL & "1" & ","
			cSQL = cSQL & cKeyID_exp & ")"
Conn.Execute(cSQL)
			
cError = CheckADOErrors(Conn,"Setting INSERT")
If cError <> "" Then
	Conn.RollbackTrans
	SaveLCWidth = cError
	cError = ""
else			
	Conn.CommitTrans
End If
End Function


Function SaveLCHeight
dim cSQL, cError, oRS, cSettingID, cKeyID_Exp

On Error Resume Next
SaveLCHeight = ""
Conn.BeginTrans

cSQL = "DELETE FROM SETTING WHERE TYPE = 'LAYOUTCTLHEIGHT' AND USER_ID = " &  Session("SecurityObj").m_UserId
Conn.Execute(cSQL)
cError = CheckADOErrors(Conn,"Setting DELETE")
		
If cError <> "" Then
	Conn.RollbackTrans
	SaveLCHeight = cError
	cError = ""
else
	Conn.CommitTrans
End If
cKeyID_Exp = "null"
								
cSQL = "{call Designer.GetValidSeq('SETTING', 'SETTING_ID', {resultset 1, outResult})}"
Set oRS = Conn.Execute(cSQL)
cError = CheckADOErrors(Conn,"Setting GetValidSeq")
			
If cError <> "" Then
	Conn.RollbackTrans
	SaveLCHeight = cError
	cError = ""
end If

cSettingID = oRS("outResult")
oRS.close
	
cSQL = "INSERT INTO SETTING (SETTING_ID,USER_ID,NAME,VALUE,TYPE,VERSION,KEY_ID) VALUES("		
			cSQL = cSQL & cSettingID & ","
			cSQL = cSQL & Session("SecurityObj").m_UserId & "," 
			cSQL = cSQL & "'LAYOUTCTLHEIGHT'," 
			cSQL = cSQL & "'" & LCHeight & "',"
			cSQL = cSQL & "'LAYOUTCTLHEIGHT',"
			cSQL = cSQL & "1" & ","
			cSQL = cSQL & cKeyID_exp & ")"
Conn.Execute(cSQL)
			
cError = CheckADOErrors(Conn,"Setting INSERT")
If cError <> "" Then
	Conn.RollbackTrans
	SaveLCHeight = cError
	cError = ""
else			
	Conn.CommitTrans
End If
End Function


Function ClearMaxRecords()

On Error Resume Next
		Conn.BeginTrans
		SQL = ""
		SQL = SQL & "DELETE FROM SETTING WHERE TYPE LIKE 'DESIGNER_MAXRECORDS' AND USER_ID = " &  Session("SecurityObj").m_UserId
		Conn.Execute(SQL)
		strError = CheckADOErrors(Conn,"Setting DELETE")
		
		If strError <> "" Then
			Conn.RollbackTrans
			SaveMaxRecords = strError
			Exit Function		
		End If
		Conn.CommitTrans
		
End Function


Function SaveTreeLevel()
On Error Resume Next

		Conn.BeginTrans
		SQL = ""
		SQL = SQL & "DELETE FROM SETTING WHERE TYPE LIKE 'DESIGNER_TREELEVEL' AND USER_ID = " &  Session("SecurityObj").m_UserId
		Conn.Execute(SQL)
		strError = CheckADOErrors(Conn,"Setting DELETE")
		
		If strError <> "" Then
			Conn.RollbackTrans
			SaveTreeLevel = strError
			Exit Function		
		End If
		Conn.CommitTrans

			strKeyID = "null"
			strKeyID_Exp = "null"
			strType = "DESIGNER_TREELEVEL"
		
			strName = "DESIGNER_TREELEVEL"
			strValue = TREELEVELS
			strValue = Replace(strValue,"'","''")
								
			SQL = "{call Designer.GetValidSeq('SETTING', 'SETTING_ID', {resultset 1, outResult})}"
			Set RSNextID = Conn.Execute(SQL)
			
			strError = CheckADOErrors(Conn,"Setting GetValidSeq")
			
			If strError <> "" Then
				Conn.RollbackTrans
				SaveTreeLevel = strError
				Exit Function		
			End If

			newSettingID = RSNextID("outResult")
			
			RSNextID.close
	
			SQL = "INSERT INTO SETTING (SETTING_ID,USER_ID,NAME,VALUE,TYPE,VERSION,KEY_ID) VALUES("		
			SQL = SQL & newSettingID & ","
			SQL = SQL & Session("SecurityObj").m_UserId & "," 
			SQL = SQL & "'" & strName & "'," 
			SQL = SQL & "'" & strValue & "',"
			SQL = SQL & "'" & strType & "',"
			SQL = SQL & "1" & ","
			SQL = SQL & strKeyID_exp & ")"
			Conn.Execute(SQL)
			
			strError = CheckADOErrors(Conn,"Setting INSERT")
			If strError <> "" Then
				Conn.RollbackTrans
				SaveTreeLevel = strError
				Exit Function		
			End If
			
		
		Conn.CommitTrans
		
		SaveTreeLevel = ""

End Function

Function SaveTreeCount()
On Error Resume Next

		Conn.BeginTrans
		SQL = ""
		SQL = SQL & "DELETE FROM SETTING WHERE TYPE LIKE 'DESIGNER_TREECOUNT' AND USER_ID = " &  Session("SecurityObj").m_UserId
		Conn.Execute(SQL)
		strError = CheckADOErrors(Conn,"Setting DELETE")
		
		If strError <> "" Then
			Conn.RollbackTrans
			SaveTreeCount = strError
			Exit Function		
		End If
		Conn.CommitTrans


			strKeyID = "null"
			strKeyID_Exp = "null"
			strType = "DESIGNER_TREECOUNT"
		
			strName = "DESIGNER_TREECOUNT"
			strValue = TREECOUNT
			strValue = Replace(strValue,"'","''")
								
			SQL = "{call Designer.GetValidSeq('SETTING', 'SETTING_ID', {resultset 1, outResult})}"
			Set RSNextID = Conn.Execute(SQL)
			
			strError = CheckADOErrors(Conn,"Setting GetValidSeq")
			
			If strError <> "" Then
				Conn.RollbackTrans
				SaveTreeCount = strError
				Exit Function		
			End If

			newSettingID = RSNextID("outResult")
			
			RSNextID.close
	
			SQL = "INSERT INTO SETTING (SETTING_ID,USER_ID,NAME,VALUE,TYPE,VERSION,KEY_ID) VALUES("		
			SQL = SQL & newSettingID & ","
			SQL = SQL & Session("SecurityObj").m_UserId & "," 
			SQL = SQL & "'" & strName & "'," 
			SQL = SQL & "'" & strValue & "',"
			SQL = SQL & "'" & strType & "',"
			SQL = SQL & "1" & ","
			SQL = SQL & strKeyID_exp & ")"
			Conn.Execute(SQL)
			
			strError = CheckADOErrors(Conn,"Setting INSERT")
			If strError <> "" Then
				Conn.RollbackTrans
				SaveTreeCount = strError
				Exit Function		
			End If
			
		
		Conn.CommitTrans
		
		SaveTreeCount = ""

End Function

Function ClearRecords()

On Error Resume Next
		Conn.BeginTrans
		SQL = ""
		SQL = SQL & "DELETE FROM SETTING WHERE TYPE LIKE 'DESIGNER_MAXRECORDS' AND USER_ID = " &  Session("SecurityObj").m_UserId
		Conn.Execute(SQL)
		strError = CheckADOErrors(Conn,"Setting DELETE")
		
		If strError <> "" Then
			Conn.RollbackTrans
			SaveMaxRecords = strError
			Exit Function		
		End If
		Conn.CommitTrans
		
End Function

 %>

</SCRIPT>
</HEAD>
<BODY>
</BODY>
