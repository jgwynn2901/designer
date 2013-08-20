<%
sub IncludeNavigateBackJS(inContainerContext)
	if CStr(inContainerContext) = "DRILLIN" then %>
<SCRIPT LANGUAGE="JScript" SRC='..\LIB\NavigateBack.js'></SCRIPT>		
<%	
	end if
end sub

function IncludeNavigateBackInvoke(inContainerContext)
	if CStr(inContainerContext) = "DRILLIN" then 
		IncludeNavigateBackInvoke = true
	else
		IncludeNavigateBackInvoke = false
	end if
end function
%>