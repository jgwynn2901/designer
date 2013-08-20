<SCRIPT LANGUAGE="JavaScript" FOR="BRBtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
	case "EDITBUTTONCLICK":
       	EditClick()
		break;

	case "NEWBUTTONCLICK":
		NewClick()
		break;

 	case "COPYBUTTONCLICK":
		CopyClick()
		break;

	case "PASTEBUTTONCLICK":
		PasteClick()
		break;

	case "SEARCHBUTTONCLICK":
		SearchClick()
		break;

	case "ATTACHBUTTONCLICK":
		AttachClick()
		break;

	case "REFRESHBUTTONCLICK":
		RefreshClick()
		break;

	case "REMOVEBUTTONCLICK":
		RemoveClick()
		break;

		default:
		break;
	}
   
</SCRIPT>