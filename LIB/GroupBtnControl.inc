<SCRIPT LANGUAGE="JavaScript" FOR="GroupBtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
		case "ATTACHBUTTONCLICK":
				ExeButtonsAttach();
			break;
		case "NEWBUTTONCLICK":
				ExeButtonsNew();
			break;
		case "REMOVEBUTTONCLICK":
				ExeButtonsRemove();
			break;
		default:
				alert("NOT A FEATURE");
			break;
	}
</SCRIPT>