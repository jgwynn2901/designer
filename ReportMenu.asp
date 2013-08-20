<!--#include file="lib\common.inc"-->
<!--#include file="lib\security.inc"-->
<% Response.Expires = 0 
dim lHasBillingAccess, lHasAgentAccess, lHasHABillingAccess, lHasFinanceBillingAccess
	
lHasBillingAccess = HasAutomaticSecurityPrivilege() Or HasViewPrivilege("FNSD_BILLINGREP","")
lHasAgentAccess = HasAutomaticSecurityPrivilege() Or HasViewPrivilege("FNSD_AGENTBILLING","")
lHasHABillingAccess = HasAutomaticSecurityPrivilege() Or HasViewPrivilege("FNSD_HABILLING","")
lHasFinanceBillingAccess = HasViewPrivilege("FNSD_FINANCE_BILLING","")
%>
<HTML>
<HEAD>
<META name=VI60_defaultClientScript content=JavaScript>
<link rel="stylesheet" type="text/css" href="FNSDESIGN.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
parent.frames("WORKAREA").location.href = "blank.htm";
}

//-->
</SCRIPT>
</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 BGCOLOR=BLACK LANGUAGE=javascript onload="return window_onload()">
<script language="JavaScript" src="reports/toolbar.js"></script>
<script language="JavaScript">

var ToolBar_Supported = ToolBar_Supported ;
var lIsFirst = true;
if (ToolBar_Supported != null && ToolBar_Supported == true)
{
	Frame_Supported = false;
    var environment = '<%=Session("ENVIRONMENT_ABBREVIATION")%>';    
	if  (environment== "P")  {
		setDefaultICPMenuColor("Red", "White", "#6495ed"); 
	}else if (environment == "PP")  {
		setDefaultICPMenuColor("Yellow", "Black", "#6495ed");	
	}else if(environment == "QA")  {
		setDefaultICPMenuColor("Blue", "White", "#6495ed"); 		
	}else if (environment == "BA"){ 
	setDefaultICPMenuColor("Black", "White", "#6495ed"); 
	}else { // Black as Default
	    setDefaultICPMenuColor("Black", "White", "#6495ed");
	}
	
	setToolbarBGColor("white");
	<%If lHasBillingAccess Then%>
		addICPMenu("BillingMenu", " Billing Report", "Service Billing Report", "Reports/Billing/default.htm", "WORKAREA");
	<%end if%>
	<%If lHasAgentAccess Then%>
		addICPMenu("AgentMenu", " Agent Billing", "Agent Billing Report", "Reports/Agent/default.htm", "WORKAREA");
	<%end if%>
	<%If lHasHABillingAccess Then%>
		addICPMenu("HAMenu", " Health Alliance Billing", "Health Alliance Billing Report", "Reports/HABilling/default.htm", "WORKAREA");
	<%end if%>
	<%If lHasFinanceBillingAccess Then%>
		addICPMenu("FinanceMenu", " Finance Report", "Finance Billing Report", "Reports/financeBilling/finRep.asp", "WORKAREA");
	<%end if%>
	lNoBar = true;
	addICPMenu("BackMenu", " <= Back", "", "TopPane.asp", "_self");	
	drawToolbar();
}

</script>
</BODY>
</HTML>
