<%
'***************************************************************
'event handler for MABtnControl.
'
'$History: $ 
'***************************************************************
%>
<SCRIPT LANGUAGE="JavaScript" FOR="MABtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
		case "EDITBUTTONCLICK":
				ExeEditMailboxRule();
			break;

		case "NEWBUTTONCLICK":
				ExeNewMailboxRule();
			break;

		case "REMOVEBUTTONCLICK":
				ExeRemoveMailboxRule();
			break;

		default:
			break;
	}
   
</SCRIPT>