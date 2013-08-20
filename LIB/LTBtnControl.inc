<SCRIPT LANGUAGE="JavaScript" FOR="LTBtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
		case "EDITBUTTONCLICK":
				ExeEditCode();
			break;

		case "NEWBUTTONCLICK":
				ExeNewCode();
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
				ExeRemoveCode();
			break;

		default:
			break;
	}
   
</SCRIPT>