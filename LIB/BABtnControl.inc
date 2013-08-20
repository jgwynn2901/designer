<SCRIPT LANGUAGE="JavaScript" FOR="BABtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
		case "EDITBUTTONCLICK":
				ExeEditBranchRule();
			break;

		case "NEWBUTTONCLICK":
				ExeNewBranchRule();
			break;

		case "COPYBUTTONCLICK":
				alert("NOT A FEATURE");
			break;

		case "PASTEBUTTONCLICK":
				alert("NOT A FEATURE");
			break;

		case "SEARCHBUTTONCLICK":
				alert("NOT A FEATURE");
			break;

		case "REMOVEBUTTONCLICK":
				ExeRemoveBranchRule();
			break;

		default:
			break;
	}
   
</SCRIPT>