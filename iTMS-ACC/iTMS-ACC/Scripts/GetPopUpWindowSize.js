function GetWindowSizeForPopup(sPopupType) {
	switch (Number(sPopupType)) {
	case 1:
		return "ItemSelectRelPartyCommon.asp:500:850";
	case 2:
		return "PartySelection.asp:500:500";
	case 3:
		return "DynamicNoSelection.asp:500:850";
	case 4:
		return "MisPartySelection.asp:500:350";
	case 5:
		return "GLHeadSelection.asp:500:350";
	case 6:
		return "ItemSelectCommonForQuote.asp:500:850";
	case 7:
		return "PartySelectionWithParTypeSel.asp:500:500";
	case 8:
		return "EmpSelPop.asp:500:500";
	case 9:
		return "PartySelPop.asp:500:500";
	case 10:
		return "SupplierItemSelectCommon.asp:500:850";
	case 11:
		return "PackingLotSerialDetails.asp:460:650";
	case 12:
		return "PartySelectionAcc.asp:500:500";
	default:
		return "";
	}
}
