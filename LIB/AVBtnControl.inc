<SCRIPT LANGUAGE="JavaScript" FOR="AVBtnControl" EVENT="onscriptletevent (event, obj)">
   switch (event)
	{
		case "EDITBUTTONCLICK":
				EditAccVendor();
			break;

		case "NEWBUTTONCLICK":
				AddAccVendor();
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
				DelAccVendor();
			break;

		default:
			break;
	}
   
</SCRIPT>
