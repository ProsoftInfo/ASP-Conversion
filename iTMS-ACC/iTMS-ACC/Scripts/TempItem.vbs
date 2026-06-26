Function CreateTempItem(appCode,modCode,creStage,sItemtype,sOrgCode)
	OutDataValue = showModalDialog("../../inventory/master/tempitmCreationEntry.asp?appCode="&appCode&"&modCode="&modCode&"&creStage="&creStage&"&itmType="&sItemtype&"&sOrgCode="&sOrgCode,"A","dialogHeight:230px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No")
	CreateTempItem = OutDataValue
End Function
