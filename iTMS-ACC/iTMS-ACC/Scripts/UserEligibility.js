
function CheckUserType(sUserType,sCurrentStatus)
{
		
	if (sCurrentStatus=="")
	{
		if (sUserType == "SU")
		{
			return true;
		}
		else
		{
			return false;
		}
	}	
	else
	{
		if (sCurrentStatus == "010104" || sCurrentStatus == "04") //accounted
		{
			if (sUserType == "SU")
			{
				return true;
			}	
			else
			{
				alert("You do not have access rights to perform this action");	
				return false;
			}	
		}
	}
}
