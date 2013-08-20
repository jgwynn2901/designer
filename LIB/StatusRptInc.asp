<%
Const S_INFO = "Info"
Const S_WARNING = "Warning"
Const S_ERROR = "Error"
Const S_INTERNAL = "Internal"

Dim LogStatusPendingUpd
LogStatusPendingUpd = False

Function LogStatusGroupBegin()
	Session("StatusRptLogGrpEnabled") = True
	Session("StatusRptMsgNextGrpID") = Session("StatusRptMsgNextGrpID") + 1
	LogStatusGroupBegin = Session("StatusRptMsgNextGrpID")
End Function


Function LogStatusGroupEnd()
	Session("StatusRptLogGrpEnabled") = False
	LogStatusGroupEnd = Session("StatusRptMsgNextGrpID")
End Function


Function LogStatus(strSeverity, strMessage, strSourceTable, strSourceField, nSourceRowID, strSourceRowDesc)

	Dim IdPair, objStatus

	IdPair = "0,0"

	Set objStatus = Session("StatusRptRS")

	If IsObject(objStatus) Then

		Session("StatusRptMsgNextID") = Session("StatusRptMsgNextID") + 1

		If Session("StatusRptLogGrpEnabled") = False Then
			Session("StatusRptMsgNextGrpID") = Session("StatusRptMsgNextGrpID") + 1
		End If

		IdPair = Session("StatusRptMsgNextID") & "," & Session("StatusRptMsgNextGrpID")

		objStatus.AddNew
		
		objStatus("MsgID") = Session("StatusRptMsgNextID")
		objStatus("GrpID") = Session("StatusRptMsgNextGrpID")
		objStatus("Severity") = strSeverity
		objStatus("Message") = strMessage
		objStatus("SourceTable") = strSourceTable
		objStatus("SourceField") = strSourceField
		objStatus("SourceRowID") = nSourceRowID
		objStatus("SourceRowDesc") = strSourceRowDesc

		If LogStatusPendingUpd = False Then
			objStatus.Update
		End If
	End If

	LogStatus = IdPair
End Function


Function LogStatusEx(strSeverity, strMessage, strSourceTable, strSourceField, strSourceRowID, strSourceRowDesc, strAssocTable, strAssocField, nAssocRowID, strAssocRowDesc)

	Dim IdPair, objStatus

	IdPair = "0,0"

	Set objStatus = Session("StatusRptRS")

	If IsObject(objStatus) Then

		LogStatusPendingUpd = True
		IdPair = LogStatus(strSeverity, strMessage, strSourceTable, strSourceField, strSourceRowID, strSourceRowDesc)
		LogStatusPendingUpd = False

		objStatus("AssocTable") = strAssocTable
		objStatus("AssocField") = strAssocField
		objStatus("AssocRowID") = nAssocRowID
		objStatus("AssocRowDesc") = strAssocRowDesc

		objStatus.Update
	End If

	LogStatusEx = IdPair
End Function
%>
