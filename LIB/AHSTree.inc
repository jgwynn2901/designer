<%
Function AddExpandedNode(AHSNodeId)

	Set Obj = Session("AHSTreeCurExpandedNodes")

	part = Split(AHSNodeId,"=")
	key = "AHSID=" & part(1)

	Obj.Item(key) = CStr(key)

End Function

Function addShowNodes(AHSNodeId)

	Set Obj = Session("AHSTreeShowNodes")

	part = Split(AHSNodeId,"=")
	key = "AHSID=" & part(1)

	Obj.Item(key) = CStr(key)

End Function

Function addShowAllNode(AHSNodeId)

	Set Obj = Session("AHSTreeShowAllNodes")

	part = Split(AHSNodeId,"=")
	key = "AHSID=" & part(1)

	Obj.Item(key) = CStr(key)

End Function


Sub SetShowPolicyIndicator(bFlag)
	Session("ShowPolicyIndicator") = bFlag
End Sub


Sub SetShowCallFlowIndicator(bFlag)
	Session("SetShowCallFlowIndicator") = bFlag
End Sub


Sub SetShowRoutingPlanIndicator(bFlag)
	Session("SetShowRoutingPlanIndicator") = bFlag
End Sub


Function GetShowPolicyIndicator
	GetShowPolicyIndicator = Session("ShowPolicyIndicator")
End Function


Function GetShowCallFlowIndicator
	GetShowCallFlowIndicator = Session("GetShowCallFlowIndicator")
End Function


Function GetShowRoutingPlanIndicator
	GetShowRoutingPlanIndicator = Session("GetShowRoutingPlanIndicator")
End Function



Sub SetShowPolicyList(bFlag)
	Session("ShowPolicyList") = bFlag
End Sub


Sub SetShowCallFlowList(bFlag)
	Session("SetShowCallFlowList") = bFlag
End Sub


Sub SetShowRoutingPlanList(bFlag)
	Session("SetShowRoutingPlanList") = bFlag
End Sub


Function GetShowPolicyList
	GetShowPolicyList = Session("ShowPolicyList")
End Function


Function GetShowCallFlowList
	GetShowCallFlowList = Session("GetShowCallFlowList")
End Function


Function GetShowRoutingPlanList
	GetShowRoutingPlanList = Session("GetShowRoutingPlanList")
End Function


Sub SetTopMostNodeAllowed(AHSNodeId)
	Session("TopMostNodeAllowed") = AHSNodeId
End Sub


Function GetTopMostNodeAllowed
	GetTopMostNodeAllowed = Session("TopMostNodeAllowed")
End Function


Sub SetCurrentTopNode(AHSNodeId, AHSNodeDesc)
	Session("CurAHSTreeTopID") = AHSNodeId
	Session("CurAHSTreeTopDesc") = AHSNodeDesc
End Sub


Function GetCurrentTopNodeID
	GetCurrentTopNode = Session("CurAHSTreeTopID")
End Function

Function GetCurrentTopNodeDesc
	GetCurrentTopNode = Session("CurAHSTreeTopDesc")
End Function


Function SetFilter(AHSNodeId, FilterType, UseWhereClause, WhereClause, NodesMustIncluded, NodesMustExcluded, NodeDelimiter)

	Set Obj = Session("AHSTreeFilter")

	part = Split(AHSNodeId,"=")
	key = CStr(part(1))

	curkey = key & ":" & FilterType & ":USEWHERECLAUSE"
	Obj.Item(curkey) = CStr(UseWhereClause)

	curkey = key & ":" & FilterType & ":WHERECLAUSE"
	Obj.Item(curkey) = CStr(WhereClause)
	
	curkey = key & ":" & FilterType & ":MUSTINCLUDE"
	Obj.Item(curkey) = CStr(NodesMustIncluded)
	
	curkey = key & ":" & FilterType & ":MUSTEXCLUDE"
	Obj.Item(curkey) = CStr(NodesMustExcluded)

	curkey = key & ":" & FilterType & ":NODEDELIM"
	Obj.Item(curkey) = CStr(NodeDelimiter)

	SetFilter = true
End Function


Function SetFilterByName(AHSNodeId, FilterType, Name, Value)

	Set Obj = Session("AHSTreeFilter")

	part = Split(AHSNodeId,"=")
	key = CStr(part(1))

	curkey = key & ":" & FilterType & ":" & Name
	Obj.Item(curkey) = CStr(Value)

	SetFilterByName = true
End Function



Function RemoveFilter(AHSNodeId, FilterType)

	Set Obj = Session("AHSTreeFilter")

	part = Split(AHSNodeId,"=")
	key = CStr(part(1)) & ":" & FilterType
	

	strKeysArray = Obj.Keys
	For intLoop = 0 To Obj.Count - 1
		strCurKey = strKeysArray(intLoop)
		if Left(strCurKey,Len(key)) = key then
			Obj.Remove(strCurKey)
		end if 
	Next

End Function


Function RemoveAllFilters(FilterType)

	Set Obj = Session("AHSTreeFilter")

	key = ":" & FilterType
	
	strKeysArray = Obj.Keys
	For intLoop = 0 To Obj.Count - 1
		strCurKey = strKeysArray(intLoop)
		if InStr(strCurKey,key) > 0 then
			Obj.Remove(strCurKey)
		end if 
	Next
End Function


Function HasSpecificFilter(AHSNodeId, FilterType)
	Set Obj = Session("AHSTreeFilter")

	part = Split(AHSNodeId,"=")
	key = CStr(part(1)) & ":" & FilterType

	strKeysArray = Obj.Keys
	For intLoop = 0 To Obj.Count - 1
		strCurKey = strKeysArray(intLoop)
		if Left(strCurKey,Len(key)) = key then
			HasSpecificFilter = true
			exit function
		end if 
	Next

	HasSpecificFilter = false

End Function

Function GetSpecificFilter(AHSNodeId, FilterType, WhichFilter)
	Set Obj = Session("AHSTreeFilter")

	part = Split(AHSNodeId,"=")
	key = CStr(part(1)) & ":" & FilterType & ":" & WhichFilter

	strKeysArray = Obj.Keys
	strItemsArray = Obj.Items
	For intLoop = 0 To Obj.Count - 1
		strCurKey = strKeysArray(intLoop)
		if Left(strCurKey,Len(key)) = key then
			GetSpecificFilter = strItemsArray(intLoop)
			exit function
		end if 
	Next

	GetSpecificFilter = ""

End Function

Function DisplayFilter()

	Set Obj = Session("AHSTreeFilter")
	
	strKeysArray = Obj.Keys
	strItemsArray = Obj.Items
	For intLoop = 0 To Obj.Count -1
		strThisKey = strKeysArray(intLoop)
		strThisItem = strItemsArray(intLoop)
		Response.Write strThisKey & " = " & strThisItem & "<BR>"
	Next
End Function
%>