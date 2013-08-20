<%
if Application("lExecutingBillingReport") then
	Response.Write "The variable is already set. Somebody is using it."
else	
	Application("lExecutingBillingReport") = true
	Response.Write "The variable is set."
end if	
%>
