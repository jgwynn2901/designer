<SCRIPT LANGUAGE="JavaScript" FOR="UserBtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
		case "ATTACHBUTTONCLICK":
				ExeButtonsAttach();
			break;
		case "REMOVEBUTTONCLICK":
				ExeButtonsRemove();
			break;
		case "NEWBUTTONCLICK":
				ExeButtonsNew();
			break;
		default:
				alert("NOT A FEATURE");
			break;
	}
</SCRIPT>