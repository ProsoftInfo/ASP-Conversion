Function GetWindowSizeForPopup(sPopupType)
'Popup Type
'1 - ItemSelection Popup
'2 - Party Selection
'3 - Reference No Selection Popup - DynamicNoSelection
'4 - MisPartySelection
'5 - GL head Selection 
'6 - Item Seletion Common Popup for the Quotation

'ProgramName:Height:Width

	Select Case sPopupType
		Case 1 
			GetWindowSizeForPopup = "ItemSelectRelPartyCommon.asp:500:850"
		Case 2
			GetWindowSizeForPopup = "PartySelection.asp:500:500"
		Case 3
			GetWindowSizeForPopup = "DynamicNoSelection.asp:500:850"
		Case 4
			GetWindowSizeForPopup = "MisPartySelection.asp:500:350"
		Case 5
			GetWindowSizeForPopup = "GLHeadSelection.asp:500:350" 	
		Case 6
			GetWindowSizeForPopup = "ItemSelectCommonForQuote.asp:500:850"
		Case 7
			GetWindowSizeForPopup = "PartySelectionWithParTypeSel.asp:500:500"
		Case 8
			GetWindowSizeForPopup = "EmpSelPop.asp:500:500"
		Case 9
			GetWindowSizeForPopup = "PartySelPop.asp:500:500"
		Case 10 
			GetWindowSizeForPopup = "SupplierItemSelectCommon.asp:500:850"
		Case 11
			GetWindowSizeForPopup = "PackingLotSerialDetails.asp:460:650"
		Case 12
		    GetWindowSizeForPopup = "PartySelectionAcc.asp:500:500"
	End Select

End Function