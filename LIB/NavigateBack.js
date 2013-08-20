function SetInfoForNavigateBack(inParam)
{
	var i = 0;
	
	for (i=0; i< parent.frames.length; i++)
	{
		if (parent.frames(i).name == "NAVIGATE_BACK")
		{
			parent.frames("NAVIGATE_BACK").AddBackObj(inParam);
			return true;
		}
	}
	
	return false;
}