<%

function HasPrivilege(inFunction, inParentFunction, inPrivilege)
	
	If Session("SecurityObj").CanUserDo(CStr(inFunction),CStr(inPrivilege)) <> True Then  
		HasPrivilege = False
		Exit Function
	ElseIf  inParentFunction <> "" Then
		If Session("SecurityObj").CanUserDo(CStr(inParentFunction),CStr(inPrivilege)) <> True Then 	
			HasPrivilege = False
			Exit Function
		End If
	End If

	HasPrivilege = True

end function

function HasModifyPrivilege(inFunction, inParentFunction)
	HasModifyPrivilege = HasPrivilege(inFunction, inParentFunction, "MODIFY")
end function

function HasDeletePrivilege(inFunction, inParentFunction)
	HasDeletePrivilege = HasPrivilege(inFunction, inParentFunction, "DELETE")
end function

function HasAddPrivilege(inFunction, inParentFunction)
	HasAddPrivilege = HasPrivilege(inFunction, inParentFunction, "ADD")
end function

function HasViewPrivilege(inFunction, inParentFunction)
	HasViewPrivilege = HasPrivilege(inFunction, inParentFunction, "View")
end function

function HasAutomaticSecurityPrivilege()
	If CStr(Session("SecurityObj").m_UserName) = "ADMINISTRATOR" Then
		HasAutomaticSecurityPrivilege = True
	Else
		HasAutomaticSecurityPrivilege = False
	End If 
end function
	
%>
