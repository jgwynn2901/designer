<%
const cServerName = ""

if Application("lExecutingBillingReport") = true then
	response.redirect "inUse.htm"
end if
if not isEmpty( request.QueryString("AGT") ) then
	response.redirect cServerName & "/agentBilling/default.aspx?" & request.QueryString 
elseif not isEmpty( request.QueryString("HABill") ) then
	response.redirect cServerName & "/HABill/default.aspx?" & request.QueryString 
else 'TPAL-0139 Passing Connection string to Billing Summary for all clients
	' Create a form, add the session variables to pass through as Hidden Inputs
	' Post to the billing summary default page
	' Send the query string for te default page to process the report
	' as per normal once it's gathered the session variables
	Response.Write("<form name=""SessionPass"" id=""SessionPass"" action=""" & cServerName & "/billingSummary/default.aspx?" & request.QueryString & """ method=""post"" >")
	Response.Write("<input type=""hidden"" name=""ConnectionString"" value=""" & Session("ConnectionString") & """ />")
	Response.Write("</form>")
	response.Write("<script>SessionPass.submit();</script>")
end if	
%>