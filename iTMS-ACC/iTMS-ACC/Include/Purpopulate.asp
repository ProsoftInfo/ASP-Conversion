<%
	Dim iSAApplicationPop1,iSAProcessPop1,iSAActivityPop1,iEmpNoPopulate1

	iSAApplicationPop1 = Session("iApplication")
	iSAProcessPop1 = Session("iProcess")
	iSAActivityPop1 = Session("iActivity")
	iEmpNoPopulate1 = Session("employeenumber")
%>
<%
    Function popDisRecIssType(sel)
        ' Declaration of variables
		Dim dRSet,sName,sCode

		'Declaration of Objects
		Set dRSet = Server.CreateObject("ADODB.RecordSet")
		with dRSet
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ReceiptIssueTypeCode,ReceiptIssueTypeDesc FROM APP_M_ReceiptIssueTypes where ReceiptIssueTypeCode  ='"&  sel &"'  ORDER BY ReceiptIssueType"
			.ActiveConnection = con
			.Open
		end with
		if not dRset.eof then	
			popDisRecIssType =trim(dRset(1))
		end if 
		drset.close
    End Function
%>

<%
	Function popRecIssType(sel,sAppFor)
		' Declaration of variables
		Dim dRSet,sName,sCode

		'Declaration of Objects
		Set dRSet = Server.CreateObject("ADODB.RecordSet")
		with dRSet
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ReceiptIssueTypeCode,ReceiptIssueTypeDesc FROM APP_M_ReceiptIssueTypes where ApplicableFor  ='"&  sAppFor &"'  ORDER BY ReceiptIssueType"
			.ActiveConnection = con
			.Open
		end with
		set dRSet.ActiveConnection = nothing

		set sCode = dRSet(0)
		set sName = dRSet(1)
		Do While Not dRSet.EOF
			if cstr(trim(sel)) = cstr(trim(sCode)) then
				Response.Write("<OPTION VALUE="""&trim(cstr(sCode))&""" Selected>"&trim(sName)&"</OPTION>" &vbcrlf)
			else
				Response.Write("<OPTION VALUE="""&trim(cstr(sCode))&""">"&trim(sName)&"</OPTION>" &vbcrlf)
			end if
			dRSet.MoveNext
		Loop
		dRSet.Close
	End Function
	
 %>
<%
	' Function to populate Receipt Type list
	Function popCEXReceiptType(sel)
		' Declaration of variables
		Dim dRSet,sName,sCode

		'Declaration of Objects
		Set dRSet = Server.CreateObject("ADODB.RecordSet")
		with dRSet
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ReceiptCode,ReceiptType FROM Cex_M_ReceiptTypes ORDER BY ReceiptCode"
			.ActiveConnection = con
			.Open
		end with
		set dRSet.ActiveConnection = nothing

		set sCode = dRSet(0)
		set sName = dRSet(1)
		Do While Not dRSet.EOF
			if cstr(trim(sel)) = cstr(trim(sCode)) then
				Response.Write("<OPTION VALUE="""&trim(cstr(sCode))&""" Selected>"&trim(sName)&"</OPTION>" &vbcrlf)
			else
				Response.Write("<OPTION VALUE="""&trim(cstr(sCode))&""">"&trim(sName)&"</OPTION>" &vbcrlf)
			end if
			dRSet.MoveNext
		Loop
		dRSet.Close
	End Function
%>

<%
function getUnitName(sUnitNo)
Dim dRSet, sQuery

'Declaration of Objects
Set dRSet = Server.CreateObject("ADODB.RecordSet")

'' To fetch unit name from DCS_OrganizationUnitDefinitions

If len(trim(sUnitNo)) = 6 Then
	sQuery = "Select OrgUnitDescription From DCS_OrganizationUnitDefinitions Where " &_
			" OUDefinitionID = '" & trim(sUnitNo) & "'"
Else
	sQuery = "Select A.OrgUnitDescription From DCS_OrganizationUnitDefinitions A, tpomsGroupUnits B " &_
			" Where A.OUDefinitionID = B.OUDefinitionID and B.UnitNumber = '" & trim(sUnitNo) & "'"
End if
'Response.Write "<p> sQuery = " & sQuery
with dRSet
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set dRSet.ActiveConnection = nothing

if not dRSet.EOF then
	getUnitName = trim(dRSet(0))
else
	getUnitName = ""
end if
dRSet.Close

End Function
%>

<%
'--------------------------to be changed--------------------------------------
' Function to populate the Units list
Function populateUnit(sel)
	' Declaration of variables
	Dim dcrs,sUnitID,sUnitName,sQuery
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	If iSAApplicationPop1 <> "" then
		sQuery = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID IN (SELECT DISTINCT ORGANISATIONCODE FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & iEmpNoPopulate1 & " AND APPLICATIONCODE = " & iSAApplicationPop1 & " AND PROCESSCODE = " & iSAProcessPop1 & " AND ACTIVITYCODE = " & iSAActivityPop1 & ") ORDER BY OUDEFINITIONID"
	Else
		sQuery = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID"
	End If
	'Response.Write sQuery
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	set sUnitID = dcrs(0)
	set sUnitName = dcrs(1)
	If not dcrs.EOF then
		Do While Not dcrs.EOF
			if cstr(sel) = cstr(sUnitID) then
				Response.Write("<OPTION VALUE="""&trim(sUnitID)&""" selected>"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			else
				Response.Write("<OPTION VALUE="""&trim(sUnitID)&""">"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			end if
			dcrs.MoveNext
		Loop
	end if
	dcrs.Close

End Function
%>

<%
	' Function to populate the PartyType list
	Function populatePartyType()
		' Declaration of variables
		Dim objRs,iParTypeID,sParType,sParTypeName,sQuery

		'Declaration of Objects
		Set objRs = Server.CreateObject("ADODB.RecordSet")

		sQuery = " select PartySubType,SubTypeName,PartyType  from APP_M_PartyTypes "
		with objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
		end with
		set objRs.ActiveConnection=nothing

		set iParTypeID = objRs(0)
		set sParTypeName = objRs(1)
		set sParType = objRs(2)

		If not objRs.EOF then
			Do While Not objRs.EOF
				Response.Write("<OPTION VALUE="&trim(iParTypeID)&"|"&trim(sParType)&">"&trim(sParTypeName)&"</OPTION>")
				objRs.MoveNext
			Loop
		end if
		objRs.Close

	End Function
%>
<%
Function getUnitNoOUDefID(sOrgUnit,Flag)
Dim dRSet,strGetID

'Declaration of Objects
Set dRSet = Server.CreateObject("ADODB.RecordSet")

If Flag	= "U" Then	' To fetch Unit Number for given OU definition ID
	strGetID = "Select OrganizationUnitId from DCS_OrganizationUnitDefinitions Where OUDefinitionID='" & trim(sOrgUnit) & "'"
ElseIf Flag	= "O" Then	' To fetch OU definition ID for given Unit Number
	strGetID = "Select OUDefinitionID from DCS_OrganizationUnitDefinitions Where OrganizationUnitId='" & trim(sOrgUnit) & "'"
End if

with dRSet
	.CursorLocation = 3
	.CursorType = 3
	.Source = strGetID
	.ActiveConnection = con
	.Open
end with
set dRSet.ActiveConnection = nothing

if not dRSet.EOF then
	getUnitNoOUDefID = dRSet(0)
else
	getUnitNoOUDefID = ""
end if
dRSet.Close

End Function
%>
<%
	' Function to populate the Item Type list
	Function popSelItemType(sel)
		' Declaration of variables
		Dim dcrs,stypID,stypName,sTypeNo
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
		'	if iSAApplicationPop1 <> "" then
		'		.Source = "SELECT DISTINCT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE WHERE ITEMTYPEID IN (SELECT DISTINCT ITEMTYPEID FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & iEmpNoPopulate1 & " AND APPLICATIONCODE = " & iSAApplicationPop1 & " AND PROCESSCODE = " & iSAProcessPop1 & " AND ACTIVITYCODE = " & iSAActivityPop1 & ") ORDER BY ITEMTYPENO"
		'	else
				.Source = "SELECT DISTINCT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE ORDER BY ITEMTYPENO"
		'	end if
			.ActiveConnection = con
			.Open
		end with


		'Response.Write "<p> <p> " & dcrs.source
		set stypID = dcrs(0)
		set stypName = dcrs(1)
		set sTypeNo = dcrs(2)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				If cstr(sel) = cstr(stypID) OR cstr(sel) = cstr(sTypeNo) Then
					Response.Write("<OPTION VALUE="""&trim(stypID)&""" Selected>"&trim(stypName)&"</OPTION>" &vbcrlf)
				Else
					Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				End if
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing
	End Function
%>
<%
function getClassItemDesc(spOrg,ipClass,ipItem)

Dim spSql,rsTemp,spClassDesc,spItemDesc

Set rsTemp = server.CreateObject("Adodb.Recordset")

spSql = "SELECT IC.GroupName,IM.ItemDescription FROM INV_M_ITEMMASTER IM," &_
	" INV_M_CLASSIFICATION IC WHERE IM.OrganisationCode= '" & spOrg & "'" &_
	" AND IM.ClassificationCode = " & ipClass & " AND IM.ITEMCODE = " & ipItem & " AND " &_
	" IM.ClassificationCode = IC.GroupCode "

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

if not rsTemp.eof then
	'Blocked By manas on 23/03/2005 to display only Item Desc
	getClassItemDesc = spItemDesc + "|" & + spClassDesc
	'getClassItemDesc = spItemDesc
else
	getClassItemDesc = "NA|NA"
	'getClassItemDesc = "NA"
end if
rsTemp.close

end function
%>
<% 
	' Function to replace single Quote
	Function packQuote(Value)
	  dim sValue
	  sValue = Value
	  Const SQ = "'" ' single quote


	  sValue = trim(Replace(sValue, SQ, SQ & SQ))
	  packQuote = ucase(sValue)

	End Function

%>

<%
''''''''Function to populate activities'''''
Function popActivity
Dim rsTemp, sSQL, iActivityNo, sActivityName

Set rsTemp = Server.CreateObject("ADODB.RecordSet")
	sSql = "SELECT ActivityNumber,ActivityName FROM Pur_M_Activities order by ActivityNumber"
	with rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	end with
	set rsTemp.ActiveConnection = nothing

	set iActivityNo	= rsTemp(0)
	set sActivityName = rsTemp(1)

	Do While Not rsTemp.EOF
		Response.Write("<OPTION VALUE="""&iActivityNo&""">"&sActivityName&"</OPTION>" &vbcrlf)
		rsTemp.MoveNext
	Loop
	rsTemp.Close
End Function
%>

<%
	' Function to populate the PartyType list for particular party
	Function populatePartyTypeForParty(sPassPartyType,sPassPartySubType)
		' Declaration of variables
		Dim objRs,iParTypeID,sParType,sParTypeName,sQuery

		'Declaration of Objects
		Set objRs = Server.CreateObject("ADODB.RecordSet")

		'sQuery = " select PartySubType,SubTypeName,PartyType  from APP_M_PartyTypes "
		sQuery = " select PartySubType,SubTypeName,PartyType  from APP_M_PartyTypes where PartyType='" & sPassPartyType  & "' and PartySubType = " & sPassPartySubType & ""
		with objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
		end with
		set objRs.ActiveConnection=nothing

		set iParTypeID = objRs(0)
		set sParTypeName = objRs(1)
		set sParType = objRs(2)

		If not objRs.EOF then
			Do While Not objRs.EOF
				Response.Write("<OPTION VALUE="&trim(iParTypeID)&"|"&trim(sParType)&">"&trim(sParTypeName)&"</OPTION>")
				objRs.MoveNext
			Loop
		end if
		objRs.Close

	End Function
%>


<%
	' Function to populate Purchase Type list
	Function popSelPurTypeFull(sel)
		' Declaration of variables
		Dim dRSet,sPurTypeName,iPurTypeNum

		'Declaration of Objects
		Set dRSet = Server.CreateObject("ADODB.RecordSet")
		with dRSet
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT PURCHASETYPE,PURTYPESHORTNAME,PURCHASETYPENAME FROM APP_M_PURCHASETYPES ORDER BY PURCHASETYPE"
			.ActiveConnection = con
			.Open
		end with
		set dRSet.ActiveConnection = nothing

		set iPurTypeNum = dRSet(0)
		set sPurTypeName = dRSet(2)
		Do While Not dRSet.EOF
			if cstr(sel) = cstr(iPurTypeNum) then
				Response.Write("<OPTION VALUE="""&trim(cstr(iPurTypeNum))&""" Selected>"&trim(sPurTypeName)&"</OPTION>" &vbcrlf)
			else
				Response.Write("<OPTION VALUE="""&trim(cstr(iPurTypeNum))&""">"&trim(sPurTypeName)&"</OPTION>" &vbcrlf)
			end if
			dRSet.MoveNext
		Loop
		dRSet.Close
	End Function
%>
<%
	' Function to populate Currency list
	Function popSelCurrency(sel)
		' Declaration of variables
		Dim dRSet,sName,sCode
		Response.write sel
		'Declaration of Objects
		Set dRSet = Server.CreateObject("ADODB.RecordSet")
		with dRSet
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CurrencyCode,CurrencyShortName FROM Ms_CurrencyMaster ORDER BY CurrencyCode"
			.ActiveConnection = con
			.Open
		end with
		set dRSet.ActiveConnection = nothing

		set sCode = dRSet(0)
		set sName = dRSet(1)
		Do While Not dRSet.EOF
			if cstr(sel) = cstr(sCode) then
				Response.Write("<OPTION VALUE="""&trim(cstr(sCode))&""" Selected>"&trim(sName)&"</OPTION>" &vbcrlf)
			else
				Response.Write("<OPTION VALUE="""&trim(cstr(sCode))&""">"&trim(sName)&"</OPTION>" &vbcrlf)
			end if
			dRSet.MoveNext
		Loop
		dRSet.Close
	End Function
%>
<%	'To fetch Temporary Item Desc

	Function getTempItemDesc(iTempItemcode)
		Dim dRSet
		Set dRSet = Server.CreateObject("ADODB.RecordSet")

		with dRSet
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ItemDescription FROM Ms_TemporaryItemMaster Where TempItemCode="&iTempItemcode&""
			.ActiveConnection = con
			.Open
		end with
		set dRSet.ActiveConnection = nothing
		'Response.Write dRSet.source
		if not dRSet.EOF Then
			getTempItemDesc=trim(dRSet(0))
		Else
			getTempItemDesc = ""
		End if
		dRSet.Close
	End function
%>
<%	'To fetch Item Desc
'ADDED NEWLY ON 26 TH DEC 2007 BY MAHESHWARI 
	Function getItemDesc(iItem,iOrgID)
		Dim dRSet
		Set dRSet = Server.CreateObject("ADODB.RecordSet")

		with dRSet
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ItemDescription FROM VWITEM Where ItemCode="&iItem&" and OrganisationCode = '"&iOrgID &"' "
			.ActiveConnection = con
			.Open
		end with
		set dRSet.ActiveConnection = nothing
		'Response.Write dRSet.source
		if not dRSet.EOF Then
			getItemDesc=trim(dRSet(0))
		Else
			getItemDesc = ""
		End if
		dRSet.Close
	End function
%>
<%
	' Function to populate Document Type list
	Function popSelDocumentType(sel)
		' Declaration of variables
		Dim dRSet,sName,sCode

		'Declaration of Objects
		Set dRSet = Server.CreateObject("ADODB.RecordSet")
		with dRSet
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DocumentTypeNo,DocumentType FROM Cex_M_DocumentTypes ORDER BY DocumentTypeNo"
			.ActiveConnection = con
			.Open
		end with
		set dRSet.ActiveConnection = nothing

		set sCode = dRSet(0)
		set sName = dRSet(1)
		Do While Not dRSet.EOF
			if cstr(sel) = cstr(sCode) then
				Response.Write("<OPTION VALUE="""&trim(cstr(sCode))&""" Selected>"&trim(sName)&"</OPTION>" &vbcrlf)
			else
				Response.Write("<OPTION VALUE="""&trim(cstr(sCode))&""">"&trim(sName)&"</OPTION>" &vbcrlf)
			end if
			dRSet.MoveNext
		Loop
		dRSet.Close
	End Function
%>
<%
	Function getCurrencyName(iCurrCode)
		Dim dRSet
		Set dRSet = Server.CreateObject("ADODB.RecordSet")

		with dRSet
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CurrencyShortName FROM Ms_CurrencyMaster Where CurrencyCode="&iCurrCode&""
			.ActiveConnection = con
			.Open
		end with
		set dRSet.ActiveConnection = nothing

		if not dRSet.EOF Then
			getCurrencyName=trim(dRSet(0))
		Else
			getCurrencyName = ""
		End if
		dRSet.Close
	End function
%>
<%	'To fetch PaymentTerms Name

	Function getPayTermName(iPayTerm)
		Dim dRSet
		Set dRSet = Server.CreateObject("ADODB.RecordSet")
		if isnull(iPayTerm) or trim(iPayTerm) = "" then iPayTerm = "0"
		
		with dRSet
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT PymtTermsShortDesc FROM APP_M_PaymentTermsHeader Where PaymentTermsNo="&iPayTerm&""
			.ActiveConnection = con
			.Open
		end with
		set dRSet.ActiveConnection = nothing

		if not dRSet.EOF Then
			getPayTermName=trim(dRSet(0))
		Else
			getPayTermName = ""
		End if
		dRSet.Close
	End function
%>


<%	'To fetch Mode of Payment Name

	Function getMOPName(iMOPcode)
		Dim dRSet
		Set dRSet = Server.createObject("ADODB.RecordSet")

		with dRSet
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ShortPaymentMode FROM APP_M_ModeOfPayment Where PaymentModeNo="&iMOPcode&""
			.ActiveConnection = con
			.Open
		end with
		set dRSet.ActiveConnection = nothing

		if not dRSet.EOF Then
			getMOPName=trim(dRSet(0))
		Else
			getMOPName = ""
		End if
		dRSet.Close
	End function
%>

<%	'To fetch Mode of Despatch Name

	Function getMODName(iMODcode)
		Dim dRSet
		Set dRSet = Server.CreateObject("ADODB.RecordSet")

		with dRSet
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ShortDespatchMode FROM APP_M_ModeOfDespatch Where DespatchModeNo="&iMODcode&""
			.ActiveConnection = con
			.Open
		end with
		set dRSet.ActiveConnection = nothing

		if not dRSet.EOF Then
			getMODName=trim(dRSet(0))
		Else
			getMODName = ""
		End if
		dRSet.Close
	End function
%>


<%	'To fetch Basis of Pricing Name

	Function getBOPName(iBOPcode)
		Dim dRSet
		Set dRSet = Server.CreateObject("ADODB.RecordSet")

		with dRSet
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ShortBasisofPricing FROM PUR_M_BasisOfPricing Where BasisOfPricingNo="&iBOPcode&""
			.ActiveConnection = con
			.Open
		end with
		set dRSet.ActiveConnection = nothing

		if not dRSet.EOF Then
			getBOPName=trim(dRSet(0))
		Else
			getBOPName = ""
		End if
		dRSet.Close
	End function
%>
<%
' Function to populate the Units short name list
Function populateShortUnit(sel)
	' Declaration of variables
	Dim dcrs,sUnitID,sUnitName,sQuery

	If iSAApplicationPop1 <> "" then
		sQuery = " SELECT OUDEFINITIONID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID IN (SELECT DISTINCT ORGANISATIONCODE FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & iEmpNoPopulate1 & " AND APPLICATIONCODE = " & iSAApplicationPop1 & " AND PROCESSCODE = " & iSAProcessPop1 & " AND ACTIVITYCODE = " & iSAActivityPop1 & ") ORDER BY OUDEFINITIONID"
	Else
		sQuery = "SELECT OUDEFINITIONID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID"
	End If
	'response.write sQuery


	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	set sUnitID = dcrs(0)
	set sUnitName = dcrs(1)
	If not dcrs.EOF then
		Do While Not dcrs.EOF
			if trim(sel) = trim(sUnitID) then
				Response.Write("<OPTION VALUE="""&trim(sUnitID)&""" selected>"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			else
				Response.Write("<OPTION VALUE="""&trim(sUnitID)&""">"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			end if
			dcrs.MoveNext
		Loop
	end if
	dcrs.Close

End Function
%>
<%
Function getDRGfromItemClass(ipItemCode,ipClassCode,spOrgID)
	Dim rsTemp,strTemp,sDrgNo
	Set rsTemp = server.CreateObject("Adodb.Recordset")

	strTemp = "Select rtrim(DrawingNumber) from Inv_M_ItemMaster " &_
			" where Itemcode=" &trim(ipItemCode) & " and ClassificationCode=" & trim(ipClassCode) & "" &_
			" and OrganisationCode='" & trim(spOrgID) & "'"
	'response.write strTemp
	With rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = strTemp
		.Open
	End With
	Set rsTemp.ActiveConnection = Nothing
'response.write strTemp

	if not rsTemp.EOF then
		sDrgNo = rsTemp(0)
		getDRGfromItemClass = sDrgNo
	Else
		getDRGfromItemClass = "0"
	End if
	rsTemp.Close
End function
%>
<%
Function PopulateScheduleType(Sel)
Sel = ucase(trim(sel))
If Sel = "0" then
	Response.Write("<option value='0' Selected >Select Schedule Type</option>" &vbcrlf)
Else
	Response.Write("<option value='0' >Select Schedule Type</option>" &vbcrlf)
End If

If Sel = "D" then
	Response.Write("<option value='D' Selected>Date - (DD/MM/YYYY)</option>" &vbcrlf)
Else
	Response.Write("<option value='D' >Date - (DD/MM/YYYY)</option>" &vbcrlf)
End If

If Sel = "M" then
	Response.Write("<option value='M' Selected>Month &amp; Year - (MMYYYY)</option>" &vbcrlf)
Else
	Response.Write("<option value='M' >Month &amp; Year - (MMYYYY)</option>" &vbcrlf)
End If

If Sel = "W" then
	Response.Write("<option value='W' Selected>Week &amp; Year - (WWYYYY)</option>" &vbcrlf)
Else
	Response.Write("<option value='W' >Week &amp; Year - (WWYYYY)</option>" &vbcrlf)
End If

If Sel = "Y" then
	Response.Write("<option value='Y' Selected>Week of Month &amp;Year - (MMWWYYYY)</option>" &vbcrlf)
Else
	Response.Write("<option value='Y' >Week of Month &amp;Year - (MMWWYYYY)</option>" &vbcrlf)
End If

End Function
%>

<%
'----------To Get Alternate UOMs (Rate UoM)-------------------------------
Function popRateUOM(sOrgID,iClassCode,iItemCode,selUOM)
Dim iUOMCode,sUOMDesc,rsTemp,sSql,iOptBaseRate,iOptBaseOperator

Set rsTemp = server.CreateObject("ADODB.Recordset")

	''Purchase UOM
	sSql ="Select PurchaseUoM,UoMShortDescription from INV_M_ITEMORGMASTER,MS_UnitOfMeasurement " &_
	 " Where itemcode="&iItemCode&" and classificationcode="&iClassCode&" and organisationcode='"&sOrgID &"'and PurchaseUOM = UOMCode"
	Response.Write sSql
	With rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	End With
	Set rsTemp.ActiveConnection = nothing

	If Not rsTemp.EOF then
		iUOMCode = rsTemp(0)
		sUOMDesc = rsTemp(1)
		'iOptBaseRate = 1
		'iOptBaseOperator = 0
		if trim(selUOM) = trim(iUOMCode) then
			Response.Write("<OPTION VALUE="""&trim(iUOMCode)& """ Selected>"&sUOMDesc&"</OPTION>" &vbcrlf)
		else
			Response.Write("<OPTION VALUE="""&trim(iUOMCode)& """ >"&sUOMDesc&"</OPTION>" &vbcrlf)
		end if
	End if
	rsTemp.Close

	''Optional UOM
	sSql ="Select IM.UoMCode,UM.UoMShortDescription,IM.OptionToBaseRate,IM.OptionToBaseOperator from INV_M_ITEMORGOPTIONALUOM IM,MS_UnitOfMeasurement UM" &_
	 " Where IM.itemcode="&iItemCode&" and IM.classificationcode="&iClassCode&" and organisationcode='"&sOrgID &"' " &_
	 " and IM.UOMCode = UM.UOMCode And IM.OptionalUoMFor='P'"
	 Response.Write sSql
	With rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	End With
	Set rsTemp.ActiveConnection = nothing

	If Not rsTemp.EOF then

	Do while not rsTemp.Eof
		iUOMCode = rsTemp(0)
		sUOMDesc = rsTemp(1)
		'iOptBaseRate = rsTemp(2)
		'iOptBaseOperator = rsTemp(3)
		if trim(selUOM) = trim(iUOMCode) then
			Response.Write("<OPTION VALUE="""&trim(iUOMCode) & """ selected>"&sUOMDesc&"</OPTION>" &vbcrlf)
		else
			Response.Write("<OPTION VALUE="""&trim(iUOMCode)& """>"&sUOMDesc&"</OPTION>" &vbcrlf)
		end if
	rsTemp.MoveNext
	Loop


	End if
	rsTemp.Close
End Function
'-------------------------------------------------------------------------------------
%>

<%
function getOptionBaseRate(sOrgID,iClassCode,iItemCode,sRateUoM)

Dim sSql,rsItem,iOptToBaseRate,iOptToBaseOperator, sTemp
Set rsItem = Server.createobject("adodb.recordset")

	sSql = "Select IM.OptionToBaseRate,IM.OptionToBaseOperator from INV_M_ITEMORGOPTIONALUOM IM " &_
	 " Where IM.itemcode="& trim(iItemCode) &" and IM.classificationcode="& trim(iClassCode) &" and IM.organisationcode='"&trim(sOrgID)&"' " &_
	 " and IM.UOMCode = '" & trim(sRateUoM) & "' And IM.OptionalUoMFor='P'"
'response.write ssql + "<br>"
	With rsItem
		.CursorLocation = 3
		.CursorType = 3
		.Source =  sSql
		.ActiveConnection = con
		.Open
	End With
	Set rsItem.ActiveConnection = nothing

	If Not rsItem.EOF then
		iOptToBaseRate = rsItem(0)
		iOptToBaseOperator = rsItem(1)
	Else
		iOptToBaseRate = 1
		iOptToBaseOperator = 0
	End if
	rsItem.Close

	sTemp = cstr(iOptToBaseRate) + "|" + cstr(iOptToBaseOperator)

	getOptionBaseRate = sTemp

end function
%>

<%
function getRatePerQtyUoM(iOptToBaseOperator,iOptToBaseRate,RATE)
Dim dRatePerQtyUoM

	'' if conversion operator is 0 : Multiply

	if cint(iOptToBaseOperator) = 0 then	' Multiply

		dRatePerQtyUoM = cdbl(RATE) * cdbl(iOptToBaseRate)

	elseif cint(iOptToBaseOperator) = 1 then ' Divide

		'' if conversion operator is 1 : divide
		dRatePerQtyUoM = cdbl(RATE) / cdbl(iOptToBaseRate)
	end if

	getRatePerQtyUoM = dRatePerQtyUoM	'RETURN RATE/QTY UOM

end function
%>

<%
function getUoMConvFactor(sOrgID,iClassCode,iItemCode,blnFlag)

Dim sSql,rsItem, sUoM,iToStoreRate,iToStoreOpr,sTemp
Set rsItem = Server.createobject("adodb.recordset")

If blnFlag = "STORE" Then
	sSql = "Select StoresUoM,0,0 from Inv_M_ItemOrgMaster Where Organisationcode='" & trim(sOrgID) & "' and ClassificationCode=" & trim(iClassCode) & " and ItemCode=" & trim(iItemCode) & ""
ElseIf blnFlag = "PUR" Then
	sSql = "Select PurchaseUoM, PurToStoreRate , PurToStoreOperator  from Inv_M_ItemOrgMaster Where Organisationcode='" & trim(sOrgID) & "' and ClassificationCode=" & trim(iClassCode) & " and ItemCode=" & trim(iItemCode) & ""
ElseIf blnFlag = "SALES" Then
	sSql = "Select SalesUoM,SaleToStoreRate,SaleToStoreOperator from Inv_M_ItemOrgMaster Where Organisationcode='" & trim(sOrgID) & "' and ClassificationCode=" & trim(iClassCode) & " and ItemCode=" & trim(iItemCode) & ""
End if

With rsItem
	.CursorLocation = 3
	.CursorType = 3
	.Source =  sSql
	.ActiveConnection = con
	.Open
End With
Set rsItem.ActiveConnection = nothing

If Not rsItem.EOF then
	sUoM = rsItem(0)
	iToStoreRate = rsItem(1)
	iToStoreOpr = rsItem(2)
Else
	sUoM = ""
	iToStoreRate = 0
	iToStoreOpr = 0
End if
rsItem.Close

sTemp = cstr(sUoM) + ":" + cstr(iToStoreRate) + ":" + cstr(iToStoreOpr)

getUoMConvFactor = sTemp

end function
%>
<% 
'Function Newly added on Feb 05 by Maheswari
Function getReceiptAs(sPara)
	 
	if cint(sPara) = cint("0") then
		Response.Write("<OPTION VALUE=""0"" Selected>Select</OPTION>" &vbcrlf)
	else	
		Response.Write("<OPTION VALUE=""0"">Select</OPTION>" &vbcrlf)
	end if
	if cint(sPara) = cint("1") then 
		Response.Write("<OPTION VALUE=""1"" Selected>Purchased</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""1"" >Purchased</OPTION>" &vbcrlf)
	end if	
	if cint(sPara) = cint("2") then	
		Response.Write("<OPTION VALUE=""2"" Selected>Job work</OPTION>" &vbcrlf)		
	else
		Response.Write("<OPTION VALUE=""2"" >Job work</OPTION>" &vbcrlf)		
	end if	
	if cint(sPara) = cint("3") then
		Response.Write("<OPTION VALUE=""3"" Selected>Subcontract</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""3"">Subcontract</OPTION>" &vbcrlf)
	end if
	if cint(sPara) = cint("4") then
		Response.Write("<OPTION VALUE=""4"" Selected>Returns</OPTION>" &vbcrlf)	
	else
		Response.Write("<OPTION VALUE=""4"">Returns</OPTION>" &vbcrlf)	
	end if
	if cint(sPara) = cint("41") then
		Response.Write("<OPTION VALUE=""41"" Selected>------Sales</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""41"" >------Sales</OPTION>" &vbcrlf)
	end if
	if cint(sPara) = cint("42") then
		Response.Write("<OPTION VALUE=""42"" Selected>------Purchase</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""42"">------Purchase</OPTION>" &vbcrlf)
	end if
	if cint(sPara) = cint("43") then
		Response.Write("<OPTION VALUE=""43"" Selected>------Transfer</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""43"">------Transfer</OPTION>" &vbcrlf)
	end if
	if cint(sPara) = cint("5") then
		Response.Write("<OPTION VALUE=""5"" Selected>Service</OPTION>" &vbcrlf) 
	else
		Response.Write("<OPTION VALUE=""5"">Service</OPTION>" &vbcrlf) 
	end if
	if cint(sPara) = cint("51") then
		Response.Write("<OPTION VALUE=""51"" Selected>------Material</OPTION>" &vbcrlf)	
	else
		Response.Write("<OPTION VALUE=""51"">------Material</OPTION>" &vbcrlf)
	end if
	if cint(sPara) = cint("52") then
		Response.Write("<OPTION VALUE=""52"" Selected>------Person</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""52"">------Person</OPTION>" &vbcrlf)
	end if
	if cint(sPara) = cint("6") then
		Response.Write("<OPTION VALUE=""6"" Selected>Captial / Assets</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""6"">Captial / Assets</OPTION>" &vbcrlf)
	end if
	if cint(sPara) = cint("7") then
		Response.Write("<OPTION VALUE=""7"" Selected>Transfer</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""7"">Transfer</OPTION>" &vbcrlf)
	end if
	if cint(sPara) = cint("71") then
		Response.Write("<OPTION VALUE=""71"" Selected >------Loan basis</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""71"" >------Loan basis</OPTION>" &vbcrlf)
	end if
	if cint(sPara) = cint("72") then
		Response.Write("<OPTION VALUE=""72"" Selected>------Depot</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""72"">------Depot</OPTION>" &vbcrlf)
	end if
	if cint(sPara) = cint("73") then
		Response.Write("<OPTION VALUE=""73"" Selected>------Consignment</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""73"">------Consignment</OPTION>" &vbcrlf)
	end if

	
End Function
%> 
<%
'Function Newly added on Feb 05 by Maheswari
Function getReferenceType(sPara)
	if cint(sPara) = cint("0") then	
		Response.Write("<OPTION VALUE=""0"" Selected>Select</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""0"">Select</OPTION>" &vbcrlf)
	end if
	
	if cint(sPara) = cint("01") then	
		Response.Write("<OPTION VALUE=""01"" Selected>Purchase Order</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""01"">Purchase Order</OPTION>" &vbcrlf)
	end if
	
	if cint(sPara) = cint("02") then	
		Response.Write("<OPTION VALUE=""02"" Selected>Purchase Return Rcpt.</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""02"">Purchase Return Rcpt.</OPTION>" &vbcrlf)
	end if
	'Response.Write("<OPTION VALUE=""03"" >Supplier Samples For Approval</OPTION>" &vbcrlf)
	
	if cint(sPara) = cint("04") then	
		Response.Write("<OPTION VALUE=""04"" Selected >Subcontract Orders</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""04"">Subcontract Orders</OPTION>" &vbcrlf)
	end if
	
	if cint(sPara) = cint("11") then	
		Response.Write("<OPTION VALUE=""11"" Selected>Job Work Order</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""11"" >Job Work Order</OPTION>" &vbcrlf)
	end if
	
	if cint(sPara) = cint("12") then		
		Response.Write("<OPTION VALUE=""12"" Selected>Service Gate Pass</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""12"" >Service Gate Pass</OPTION>" &vbcrlf)
	end if
	
	
	'Response.Write("<OPTION VALUE=""05"" >Job Order Rework/Replacement</OPTION>" &vbcrlf)
	
	'if cint(sPara) = cint("06") then	
	'	Response.Write("<OPTION VALUE=""06"" Selected>Customer Sample Items </OPTION>" &vbcrlf)
	'else
	'	Response.Write("<OPTION VALUE=""06"" >Customer Sample Items </OPTION>" &vbcrlf)
	'end if
	
	if cint(sPara) = cint("07") then	
		Response.Write("<OPTION VALUE=""07"" Selected>Sales Return Invoice</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""07"" >Sales Return Invoice</OPTION>" &vbcrlf)
	end if
		
	if cint(sPara) = cint("09") then	
		Response.Write("<OPTION VALUE=""09"" Selected>Inter-unit Transfer DC</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""09"" >Inter-unit Transfer DC</OPTION>" &vbcrlf)
	end if
	
	if cint(sPara) = cint("08") then	
		Response.Write("<OPTION VALUE=""08"" Selected>Transfer Return DC</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE=""08"">Transfer Return DC</OPTION>" &vbcrlf)
	end if
	'Response.Write("<OPTION VALUE=""10"" >Without Reference</OPTION>" &vbcrlf)
End Function
%>