function GoToModuleHome()
{
	var passedValue;
	passedValue = window.document.forms[0].cmbModules.value;

	var strArray = passedValue.split("~");

	var rdrPath = strArray[1]+"?AppCode="+strArray[0];

//alert(window.document.forms[0].cmbModules.item(window.document.forms[0].cmbModules.selectedindex).index);
//alert(window.document.forms[0].cmbModules.item(window.document.forms[0].cmbModules.selectedindex).selected);
//window.location.href(rdrPath);
	window.location.href=rdrPath;
	return false;
}

function redirectToPage() {
	switch(window.document.forms[0].moduleOption.value)
		{
			case "0":
			{
				window.location.href("../Maintenance/Index_Maintenance.asp")
				break;
			}
			case "1":
			{
				window.location.href("../production/index_production.html")
				break;
			}
			case "2":
			{
				window.location.href("../inventory/index_inventory.asp")
				break;
			}
			case "3":
			{
				window.location.href("../purchase/Index_Purchase.html")
				break;
			}
			case "4":
			{
				window.location.href("../sales/index_sales.html")
				break;
			}
			case "5":
			{
				window.location.href("../accounts/Index_accounts.html")
				break;
			}
			case "6":
			{
				window.location.href("../qualitycontrol/Index_QC.html")
				break;
			}
			case "7":
						{
							window.location.href("../EnggServices/Index_EnggServices.html")
							break;
			}

	}// end of switch statement
}// end of redirectToPage function