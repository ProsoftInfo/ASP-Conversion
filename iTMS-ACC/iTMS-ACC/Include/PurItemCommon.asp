<%
'' To retrieve Item Full description for given Item code Or Drawing version no.
function getItemFullDesc(spOrg,ipClass,ipItem,spDrgNo,spFlag)

Dim spSql,rsTemp,spClassDesc,spItemDesc

Set rsTemp = server.CreateObject("Adodb.Recordset")

spSql = "SELECT ItemDescription FROM INV_M_ITEMMASTER  " &_
	" WHERE OrganisationCode= '" & spOrg & "'" &_
	" AND ClassificationCode = " & ipClass & " AND ITEMCODE = " & ipItem & ""

'response.write spSql
With rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = con
	.Source = spSql
	.Open
End With
Set rsTemp.ActiveConnection = Nothing

'Set spClassDesc = rsTemp(0)
'Set spItemDesc  = rsTemp(1)

Set spItemDesc  = rsTemp(0)

'Response.Write "<p> Eof = " + trim(rsTemp.eof)

if not rsTemp.eof then
	'getClassItemDesc = spItemDesc + "|" & + spClassDesc
	getItemFullDesc = spItemDesc
else
	'getClassItemDesc = "NA|NA"
	getItemFullDesc = "NA"
end if
rsTemp.close

end function
%>

<%
'' To retrieve Item Full description for given Item code Or Drawing version no.
'' Itemname - Cat. No - Drg No. - MGR No. - Pg. No. - Pos. No.
function getItemDescCatNo(spOrg,ipClass,ipItem,spDrgNo,spFlag)

Dim spSql,rsTemp,spClassDesc,spItemDesc,sCatlogueNo, sItemDrgNo, sMGRNo,sItemPgNo,sItemPosNo

Set rsTemp = server.CreateObject("Adodb.Recordset")

If spFlag = "I"	Then	'Item Code & Classification are passed
	spSql = "SELECT IC.GroupName,IM.ItemDescription,isnull(IM.DrawingNumber,''),isnull(IM.CatalogueNo,''),isnull(IM.MGRNo,''),isnull(IM.PageNo,''),isnull(IM.PositionNo,'') FROM INV_M_ITEMMASTER IM," &_
		" INV_M_CLASSIFICATION IC WHERE IM.OrganisationCode= '" & spOrg & "'" &_
		" AND IM.ClassificationCode = " & ipClass & " AND IM.ITEMCODE = " & ipItem & " AND " &_
		" IM.ClassificationCode = IC.GroupCode "
ElseIf spFlag = "D"	Then	'Drawing Version No. is passed
	spSql = "SELECT IC.GroupName,IM.ItemDescription,isnull(IM.DrawingNumber,''),isnull(IM.CatalogueNo,''),isnull(IM.MGRNo,''),isnull(IM.PageNo,''),isnull(IM.PositionNo,'') FROM INV_M_ITEMMASTER IM, " &_
		" INV_M_CLASSIFICATION IC, tpomsitmUnitrel IU WHERE IM.OrganisationCode=IU.OrganisationCode " &_
		" AND IM.ClassificationCode = IU.ClassificationCode AND IM.ITEMCODE =IU.ITEMCODE " &_
		" AND IM.ClassificationCode = IC.GroupCode and IU.DrawingVersionNo = '" & spDrgNo & "'"
End if
'response.write spSql
With rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = con
	.Source = spSql
	.Open
End With
Set rsTemp.ActiveConnection = Nothing

Set spClassDesc = rsTemp(0)
Set spItemDesc  = rsTemp(1)
Set sItemDrgNo = rsTemp(2)
Set sCatlogueNo  = rsTemp(3)
Set sMGRNo = rsTemp(4)
Set sItemPgNo  = rsTemp(5)
Set sItemPosNo = rsTemp(6)

if not rsTemp.eof then
	if trim(sItemDrgNo) <> "" then spItemDesc = spItemDesc + " - " + sItemDrgNo
	if trim(sCatlogueNo) <> "" then spItemDesc = spItemDesc + " - " + sCatlogueNo
	if trim(sMGRNo) <> "" then spItemDesc = spItemDesc + " - " + sMGRNo
	if trim(sItemPgNo) <> "" then spItemDesc = spItemDesc + " - " + sItemPgNo
	if trim(sItemPosNo) <> "" then spItemDesc = spItemDesc + " - " + sItemPosNo

	getItemDescCatNo = spItemDesc
else
	getItemDescCatNo = "NA"
end if
rsTemp.close

end function
%>

<%
'function to Populate Item List based on Unit Code / Organisation code, Item type
Function getItemList(pUnitCode,pOrgCode,sItemTypeID)
Dim sOrgCode,sSql,sDrawNo, sGroupName, sShortDesc
Dim dRSet

	' Declaration of variables

	if Trim(pUnitCode)<>"0" then
		sOrgCode = getUnitNoOUDefID(pUnitCode,"O")
	else
		sOrgCode	= pOrgCode
	end if


	sSql = " select A.DrawingVersionNo,v2.groupName,v2.ItemDescription " &_
		"from tpoMsItmUnitRel A, vwItem v1, vwItemDetails v2 where " &_
		"A.ItemCode = v2.ItemCode and A.ClassificationCode=v2.GroupCode " &_
		"and  A.OrganisationCode = v1.OrganisationCode  " &_
		"and A.ItemCode = v1.ItemCode and A.ClassificationCode=v1.ClassificationCode " &_
		"and v2.ItemTypeId='"&sItemTypeID&"'  " &_
		"and v2.itemcode = v1.ItemCode and v2.GroupCode = v1.classificationCode " &_
		"and v1.OrganisationCode = '"&sOrgCode &"' " &_
		"and v1.PurchaseEligible=1 and v1.ItemOnHold=0 and v1.ItemActive='Y' "


		'Declaration of Objects
		Set dRSet = Server.CreateObject("ADODB.RecordSet")
		with dRSet
			.CursorLocation = 3
			.CursorType = 3
			.Source = sSql
			.ActiveConnection = con
			.Open
		end with
		set dRSet.ActiveConnection = nothing


		If dRSet.EOF Then
			Response.Write("<OPTION VALUE=""0"">Item Not Defined for the Item Type</OPTION>" &vbcrlf)
		Else
			set sDrawNo = dRSet(0)
			set sGroupName = dRSet(1)
			set sShortDesc = dRSet(2)

			Do While Not dRSet.EOF
				Response.Write("<OPTION VALUE="""&trim(sDrawNo)&""">"&trim(sGroupName)&"--"&trim(sShortDesc) &"</OPTION>" &vbcrlf)
				dRSet.MoveNext
			Loop
		End If
		dRSet.Close
End Function
%>


<%
'function to Populate Item List based on Unit Code / Organisation code, Item type, classCode
Function getClassItemList(pUnitCode,pOrgCode,sItemTypeID,sClassCode)

' Declaration of variables
Dim sOrgCode,sSql,sDrawNo, sGroupName, sShortDesc
Dim dRSet,saCode,iCtr,blnDataExist




	blnDataExist = false
	if Trim(pUnitCode)<>"0" then
		sOrgCode = getUnitNoOUDefID(pUnitCode,"O")
	else
		sOrgCode	= pOrgCode
	end if

	saCode = split(sClassCode,",")
	for iCtr = 0 to ubound(saCode)

		sSql = " select A.DrawingVersionNo,v2.groupName,v2.ItemDescription " &_
			"from tpoMsItmUnitRel A, vwItem v1, vwItemDetails v2 where " &_
			"A.ItemCode = v2.ItemCode and A.ClassificationCode=v2.GroupCode " &_
			"and  A.OrganisationCode = v1.OrganisationCode  " &_
			"and A.ItemCode = v1.ItemCode and A.ClassificationCode=v1.ClassificationCode " &_
			"and v2.ItemTypeId='"&sItemTypeID&"'  " &_
			"and v2.itemcode = v1.ItemCode and v2.GroupCode = v1.classificationCode " &_
			"and v1.OrganisationCode = '"&sOrgCode &"' " &_
			"and A.ClassificationCode='"&saCode(iCtr)&"' "  & _
			"and v1.PurchaseEligible=1 and v1.ItemOnHold=0 and v1.ItemActive='Y' "

			'Response.Write "<p> test" & sSql

			Set dRSet = Server.CreateObject("ADODB.RecordSet")
			with dRSet
				.CursorLocation = 3
				.CursorType = 3
				.Source = sSql
				.ActiveConnection = con
				.Open
			end with
			set dRSet.ActiveConnection = nothing


			If not dRSet.EOF Then

				blnDataExist = true
				set sDrawNo = dRSet(0)
				set sGroupName = dRSet(1)
				set sShortDesc = dRSet(2)

				Do While Not dRSet.EOF
					Response.Write("<OPTION VALUE="""&trim(sDrawNo)&""">"&trim(sGroupName)&"--"&trim(sShortDesc) &"</OPTION>" &vbcrlf)
					dRSet.MoveNext
				Loop
			End If
			dRSet.Close
		Next

		'if not blnDataExist then
		'	Response.Write("<OPTION VALUE=""0"">Item Not Defined for the Item Type</OPTION>" &vbcrlf)
		'end if
End Function
%>

<%
' note : this function is used to return company ITem Code for Regular / temporary item
'about Parameter :  1.sDrgno - Drawing Version Number of Regular Item
'					2.nItemCode it has contain regular itemcode OR Temporary Item code
'					3.sTempItem - it contain "Y" for Temporary Item Else "" or "N"
Function GetCompanyItemCode(sDrgno,nItemCode,sTempItem)
Dim sRetVal,sSqlTemp

Dim saTemp

Dim rsTemp

set rsTemp = server.CreateObject("ADODB.RecordSet")

sSqlTemp = ""
sRetVal = ""

'if Drawing Version No is Passed then find Item Code of Regular Item
'if Trim(sTempItem) = "" or trim(sTempItem) = "N" then
'	saTemp = split(getItemClassfromDRG(sDrgno),"|")
'	if UBound(saTemp) >= 0 then
'		nItemCode = saTemp(0)
'	end if
'end if 'if Trim(sTempItem) = "" or trim(sTempItem) = "N" then

if trim(nItemCode) <> "" then
	if Trim(sTempItem) = "" or trim(sTempItem) = "N" then
		sSqlTemp = "select isNull(CompanyItemCode,'') from Inv_M_ItemMaster where ItemCode = " & nItemCode
	elseif Trim(sTempItem) = "Y" then
		sSqlTemp = "select isNull(GenItemCode,'') from Ms_TemporaryItemMaster where TempItemCode = " & nItemCode
	end if
end if 'if trim(nItemCode) <> "" then

'Response.Write "<p> " & sSqlTemp

if trim(sSqlTemp) <> ""  then
	with rsTemp
		.ActiveConnection = con
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSqlTemp
		.Open
	end with

	set rsTemp.ActiveConnection = nothing

	if not rsTemp.EOF then
		sRetVal = trim(rsTemp(0))
	end if
	rsTemp.Close
end if 'if trim(sSqlTemp) <> ""  then
GetCompanyItemCode = sRetVal
End Function
%>