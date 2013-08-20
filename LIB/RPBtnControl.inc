<SCRIPT LANGUAGE="JavaScript" FOR="RPBtnControl" EVENT="onscriptletevent (event, obj)">
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

	case "REMOVEBUTTONCLICK":
		RemoveClick()
		break;
	
	case "REFRESHBUTTONCLICK"
		RefreshClick();
		break;
	case "ATTACHBUTTONCLICK"
		AttachClick()
		break;
}
   
</SCRIPT>