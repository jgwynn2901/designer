<%

Dim LogRefCountPendingUpd
LogRefCountPendingUpd = False

Function LogRefCountGroupBegin()
	Session("RefCountRptLogGrpEnabled") = True
	Session("RefCountRptMsgNextGrpID") = Session("RefCountRptMsgNextGrpID") + 1
	LogRefCountGroupBegin = Session("RefCountRptMsgNextGrpID")
End Function


Function LogRefCountGroupEnd()
	Session("RefCountRptLogGrpEnabled") = False
	LogRefCountGroupEnd = Session("RefCountRptMsgNextGrpID")
End Function


Function LogRefCount(strMessage, strSourceTable, strSourceField, nSourceRowID, strSourceRowDesc)

	Dim IdPair, objRefCount

	IdPair = "0,0"

	Set objRefCount = Session("RefCountRptRS")

	If IsObject(objRefCount) Then

		Session("RefCountRptMsgNextID") = Session("RefCountRptMsgNextID") + 1

		If Session("RefCountRptLogGrpEnabled") = False Then
			Session("RefCountRptMsgNextGrpID") = Session("RefCountRptMsgNextGrpID") + 1
		End If

		IdPair = Session("RefCountRptMsgNextID") & "," & Session("RefCountRptMsgNextGrpID")

		objRefCount.AddNew
		
		objRefCount("MsgID") = Session("RefCountRptMsgNextID")
		objRefCount("GrpID") = Session("RefCountRptMsgNextGrpID")
		objRefCount("Message") = strMessage
		objRefCount("SourceTable") = strSourceTable
		objRefCount("SourceField") = strSourceField
		objRefCount("SourceRowID") = nSourceRowID
		objRefCount("SourceRowDesc") = strSourceRowDesc

		If LogRefCountPendingUpd = False Then
			objRefCount.Update
		End If
	End If

	LogRefCount = IdPair
End Function


Function LogRefCountEx(strMessage, strSourceTable, strSourceField, strSourceRowID, strSourceRowDesc, strAssocTable, strAssocField, nAssocRowID, strAssocRowDesc)

	Dim IdPair, objRefCount

	IdPair = "0,0"

	Set objRefCount = Session("RefCountRptRS")

	If IsObject(objRefCount) Then

		LogRefCountPendingUpd = True
		IdPair = LogRefCount(strMessage, strSourceTable, strSourceField, strSourceRowID, strSourceRowDesc)
		LogRefCountPendingUpd = False

		objRefCount("AssocTable") = strAssocTable
		objRefCount("AssocField") = strAssocField
		objRefCount("AssocRowID") = nAssocRowID
		objRefCount("AssocRowDesc") = strAssocRowDesc

		objRefCount.Update
	End If

	LogRefCountEx = IdPair
End Function


Sub ClearRefCountRpt()
	Dim objRefCountRpt
	Set objRefCountRpt = Session("RefCountRptRS")
	
	If Not IsObject(objRefCountRpt) Then
		Exit Sub
	End If

	If objRefCountRpt.RecordCount = 0 Then
		Exit Sub
	End If
	
	objRefCountRpt.MoveFirst
	Do While Not objRefCountRpt.EOF
		objRefCountRpt.Delete
		objRefCountRpt.MoveNext
	Loop
End Sub
%>
