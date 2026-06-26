<%@ Language=VBScript %>
<%	option explicit
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ParCreate_Edit_Entry.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 15,2010
	'Modified By				:
	'Modified By				:
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:
	'Procedures/Functions Used	:
	'Internal Variables			:
	'Database					:
	'Queries Used				:
	'Counters					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->
<!--#include file="../../include/populate.asp"-->
<%
dim sQuery,objRs,iParty,sCallTy,Temparr,Unitarr,sAction
Dim oDOM,MainNode,Root,sOrgName
Dim iAgentCount,iPerfCount,iLocationCount,iContactCount,iUnitCount,iRepCount
Dim sGroupParentName,sGroupPartyCode,sOrgCode

iParty = Request.QueryString("PartyCode")

sOrgCode = Session("organizationcode")
Set objRs = Server.CreateObject("ADODB.RecordSet")
sCallTy = Request("hCallTy")
iUnitCount = 0
sGroupPartyCode = 0
if trim(iParty)="" then
	sAction = "CREATE"
else
	sAction = "EDIT"
end if

set oDOM = Server.CreateObject("Microsoft.XMLDOM")

set Root = oDOM.createElement("Root")
oDOM.appendChild Root

sQuery = "SELECT OUDefinitionID, OrganizationUnitId, OrgUnitDescription, OrgUnitShortDescription, "&_
		 "isNull(Address1,''), isNull(Address2,''), isNull(PostCode,''), isNull(City,''), isNull(State,''), isNull(Country,0), isNull(PhoneNumber,''),isNull(FaxNumber,''), isNull(EmailID,''), "&_
		 "isNull(WeSiteURL,'') FROM  DCS_OrganizationUnitDefinitions "

With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = Con
	.Open
End With
Do While Not objRs.EOF
	Set MainNode = oDom.createElement("UNIT")
	MainNode.setAttribute "UnitID", objRs(0)
	MainNode.setAttribute "ID", objRs(1)
	MainNode.setAttribute "Desc", objRs(2)
	MainNode.setAttribute "ShortDesc", objRs(3)
	MainNode.setAttribute "Add1", objRs(4)
	MainNode.setAttribute "Add2", objRs(5)
	MainNode.setAttribute "PostCode", objRs(6)
	MainNode.setAttribute "City", objRs(7)
	MainNode.setAttribute "State", objRs(8)
	MainNode.setAttribute "Country", objRs(9)
	MainNode.setAttribute "Phone", objRs(10)
	MainNode.setAttribute "Fax", objRs(11)
	MainNode.setAttribute "EmailID", objRs(12)
	MainNode.setAttribute "Web", objRs(13)
	Root.appendChild MainNode
	objRs.MoveNext
loop
objRs.Close



sQuery = "Select OrganizationName From DCS_Organization "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sOrgName = objRs(0)
End IF
objRs.Close

if sAction="EDIT" then

sQuery = "Select PartyCode,PartyName from APP_M_PartyMaster where PartyCode "&_
		 "in (Select ParentPartyCode from APP_M_PartyMaster where"&_
		 " PartyCode = "& iParty &" and PartyCode<>0)"
'		 Response.Write sQuery
	objRs.Open sQuery,con
	if not objRs.EOF then
		sGroupPartyCode = objRs(0)
		sGroupParentName = trim(objRs(1))
	end if
	objRs.Close
end if


oDOM.Save server.MapPath("../Temp/Transaction/"&Session.SessionID&"-UNITDET.xml")


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/trim.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<XML ID="UNITDET" src="<%="../Temp/Transaction/"&Session.SessionID&"-UNITDET.xml"%>"></XML>
<XML ID="OutData" ></XML>
<XML id="PartyData"><Root/></XML>
<XML id="TempData"><Root/></XML>
<XML id="GroupData"><Root/></XML>
<SCRIPT LANGUAGE=vbscript>
'*********************************************
Function ViewData()
    if trim(document.formname.hPartyCode.value)<>"" then
	    document.formname.action = "ParDetailsView.asp?PartyCode="& document.formname.hPartyCode.value
	    document.formname.submit
	else
	    alert("Party Details Cannot view because Party is not available")
	    exit function
	end if
End Function
'**************************************************
Function ControlData()
    if Trim(document.formname.hPartyCode.value)<>"" then
        document.formname.action = "PartyControlData.asp?PartyCode="& document.formname.hPartyCode.value
        document.formname.submit
    else
        alert("Party Controls Cannot View because Party is not available")
        exit function
    end if
End Function
'**************************************************
Function GoToMain()
document.formname.action = "ParDisplayGrid.asp"
document.formname.submit
End Function
'***********************************
Function GetGroup()
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth

	if (document.formname.radGroupType(1).checked =true) or (document.formname.radGroupType(2).checked =true) then
	'	set OutValue = showModalDialog("ParentPartySelection.asp",GroupData,"Status:No;")
	'	sAct = UCase(trim(OutValue.getAttribute("Action")))
	'	sQuery = trim(OutValue.getAttribute("PassQuery"))
	'	if ucase(trim(sAct)) <> "CLOSE" then
	'		do while sAct <> "DONE"
	'			set OutValue=showModalDialog("ParentPartySelection.asp?" & sQuery,GroupData,"status:no")
	'			sAct = UCase(trim(OutValue.getAttribute("Action")))
	'			sQuery = trim(OutValue.getAttribute("PassQuery"))
	'			if ucase(Trim(sAct)) = "CLOSE" then exit do
	'		loop
	'	end if 'if ucase(trim(sAct)) <> "CLOSE" then

	 sTempValWindowSize = GetWindowSizeForPopup("2")
    sArrTempValWindowSize = split(sTempValWindowSize,":")
    sProgramName = sArrTempValWindowSize(0)
    sPopupHeight = sArrTempValWindowSize(1)
    sPopupWidth = sArrTempValWindowSize(2)
    sUnitID = document.formname.hUnitCode.value


    	Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="&sUnitID,GroupData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
		sAct = UCase(trim(OutValue.getAttribute("Action")))
		sQuery = trim(OutValue.getAttribute("PassQuery"))
		if ucase(trim(sAct)) <> "CLOSE" then
			do while sAct <> "DONE"
				set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,GroupData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
				sAct = UCase(trim(OutValue.getAttribute("Action")))
				if ucase(Trim(sAct)) = "CLOSE" then exit do
				sQuery = trim(OutValue.getAttribute("PassQuery"))
			loop
		end if
		if OutValue.hasChildNodes() then
			for each ndChild in OutValue.childNodes
				document.formname.hParentPartyCode.value  = ndChild.getAttribute("RetField1")
				document.formname.ParentPartyName.value = ndChild.getAttribute("RetField0")
			next
		end if
	else
		document.formname.hParentPartyCode.value  ="0"
	end if
End Function
'******************************************************
Function PopulatePartyTypes(sPartyCode,sAction)
	CheckUnit
    if not CheckForm then
		exit function
	else
	    SaveXML
	end if

	set TempRoot = showModalDialog("ParUnitPopup.asp?PartyCode="&sPartyCode&"&Action="&sAction,"","dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No")
	'alert(TempRoot.xml)
	if trim(TempRoot.xml)<>"" then
		document.formname.hParUnit.value = "Y"
		PartyData.loadXML(TempRoot.xml)
	end if 'if trim(TempRoot.xml)<>"" then
	'alert(PartyData.xml)
End Function
'************************************************
Function SaveXML()

	Dim sName,sShortName,sAddress1,sAddress2,sPincode,sCity,sState
	Dim sCountry,sPhone,sMobile,sFax,sEmail,sWebsite,sECCNo,sPanNo
	Dim sSalesLocal,sSalesCentral,sGroupFlag,sGroup,sTinNumber,sOwnUnit,sInActive

	Dim Root,newElem,objhttp,ndChild,ndRoot,ndUnits,ndPartyData,ndAgentData

	sName= document.formname.txtPartyName.value
	sShortName= document.formname.txtShortName.value
	sAddress1= document.formname.txtAddress1.value
	sAddress2= document.formname.txtAddress2.value
	sPincode= document.formname.txtPinCode.value
	sCity= document.formname.txtCity.value
	sState= document.formname.txtState.value
	sCountry= document.formname.txtCountry.value
	sPhone= document.formname.txtPhone.value
	sMobile=document.formname.txtMobileNo.value
	sFax= document.formname.txtFax.value
	sEmail= document.formname.txtEmail.value
	sWebsite= document.formname.txtWebsite.value
	sECCNo= document.formname.txtECCNo.value
	sPanNo= document.formname.txtPanNo.value
	sSalesLocal= document.formname.txtSalesLocal.value
	sSalesCentral= document.formname.txtSalesCentral.value
	iCreatedBy = document.formname.hCreatedBy.value

	if document.formname.chkGroupCompany.checked =true then
		sGroupFlag = document.formname.chkGroupCompany.value
	else
		sGroupFlag = ""
	end if

	'sGroupFlag= document.formname.chkGroupCompany.value
	if document.formname.radGroupType(0).checked=true then
		sGroup = document.formname.radGroupType(0).value
	elseif document.formname.radGroupType(1).checked=true then
		sGroup = document.formname.radGroupType(1).value
	elseif document.formname.radGroupType(2).checked=true then
		sGroup = document.formname.radGroupType(2).value
	end if
'	sGroup=trim(Request.Form("radGroupType"))
	sTinNumber =document.formname.txtTinNo.value
	sOwnUnit = document.formname.hOwnUnit.value
	sInActive =document.formname.hInActive.value

	Temp=trim(document.formname.hUnitCode.value)
	arrUnit=Split(Temp,":")

	'Response.Write Temp &"<br>"
	Temp=trim(document.formname.hUnitName.value)
	arrUnitName=Split(Temp,":")

	'Response.Write Temp &"<br>"
	iPartyCode = document.formname.hPartyCode.value

	if Trim(sGroupFlag)="" then
		sGroupFlag=0
		sGroup="N"
	end if
'	alert(sGroupFlag)
'	alert(sGroup)


	Set Root = PartyData.documentElement
	if Root.hasChildNodes then
		for each ndChild in Root.childNodes
			if ndChild.nodeName="ParCode" then
				ndChild.Text= iPartyCode
			elseif ndChild.nodeName="ParName" then
				ndChild.Text= sName
			elseif ndChild.nodeName="ShortName" then
				ndChild.Text= sShortName
			elseif ndChild.nodeName="Address1" then
				ndChild.Text= sAddress1
			elseif ndChild.nodeName="Address2" then
				ndChild.Text= sAddress2
			elseif ndChild.nodeName="PinCode" then
				ndChild.Text= sPincode
			elseif ndChild.nodeName="City" then
				ndChild.Text= sCity
			elseif ndChild.nodeName="State" then
				ndChild.Text= sState
			elseif ndChild.nodeName="Country" then
				ndChild.Text= sCountry
			elseif ndChild.nodeName="Phone" then
				ndChild.Text= sPhone
			elseif ndChild.nodeName="Mobile" then
				ndChild.Text= sMobile
			elseif ndChild.nodeName="Fax" then
				ndChild.Text= sFax
			elseif ndChild.nodeName="Email" then
				ndChild.Text= sEmail
			elseif ndChild.nodeName="Website" then
				ndChild.Text= sWebsite
			elseif ndChild.nodeName="ECCNo" then
				ndChild.Text= sECCNo
			elseif ndChild.nodeName="PANNo" then
				ndChild.Text= sPanNo
			elseif ndChild.nodeName="CreatedBy" then
				ndChild.Text= iCreatedBy
			elseif ndChild.nodeName="OwnUnit" then
				ndChild.Text= sOwnUnit
			elseif ndChild.nodeName="Sales" then
				ndChild.setAttribute "Local", sSalesLocal
				ndChild.setAttribute "Central", sSalesCentral
			elseif ndChild.nodeName="TINNumber" then
				ndChild.Text= sTinNumber
			elseif ndChild.nodeName="Group" then
				ndChild.setAttribute "Flag", sGroupFlag
				ndChild.setAttribute "Type", sGroup
				ndChild.setAttribute "ParentCompany", document.formname.hParentPartyCode.value
			elseif ndChild.nodeName="Units" then
				for each newChild in ndChild.childNodes
					ndChild.removeChild(newChild)
				next

				for iCounter=0 to UBound(arrUnit)
					if Trim(sOwnUnit) <> Trim(arrUnit(iCounter)) then
						Set newElem1 = PartyData.createElement("UN")
						newElem1.setAttribute "Code", arrUnit(iCounter)
						newElem1.setAttribute "Name", arrUnitName(iCounter)
						ndChild.appendChild newElem1
					end if
				next
			elseif ndChild.nodeName="Active" then
				ndChild.setAttribute "Flag",sInActive
			end if
		next
	else
		Set newElem = PartyData.createElement("ParCode")
		newElem.Text= iPartyCode
		Root.appendChild newElem

		Set newElem = PartyData.createElement("ParName")
		newElem.Text= sName
		Root.appendChild newElem

		Set newElem = PartyData.createElement("ShortName")
		newElem.Text= sShortName
		Root.appendChild newElem


		Set newElem = PartyData.createElement("Address1")
		newElem.Text= sAddress1
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Address2")
		newElem.Text= sAddress2
		Root.appendChild newElem

		Set newElem = PartyData.createElement("PinCode")
		newElem.Text= sPincode
		Root.appendChild newElem

		Set newElem = PartyData.createElement("City")
		newElem.Text= sCity
		Root.appendChild newElem

		Set newElem = PartyData.createElement("State")
		newElem.Text= sState
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Country")
		newElem.Text= sCountry
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Phone")
		newElem.Text= sPhone
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Mobile")
		newElem.Text= sMobile
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Fax")
		newElem.Text= sFax
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Email")
		newElem.Text= sEmail
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Website")
		newElem.Text= sWebsite
		Root.appendChild newElem

		Set newElem = PartyData.createElement("ECCNo")
		newElem.Text= sECCNo
		Root.appendChild newElem

		Set newElem = PartyData.createElement("PANNo")
		newElem.Text= sPanNo
		Root.appendChild newElem

		Set newElem = PartyData.createElement("CreatedBy")
		newElem.Text= iCreatedBy
		Root.appendChild newElem

		Set newElem = PartyData.createElement("OwnUnit")
		newElem.Text= sOwnUnit
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Sales")
		newElem.setAttribute "Local", sSalesLocal
		newElem.setAttribute "Central", sSalesCentral
		Root.appendChild newElem

		Set newElem = PartyData.createElement("TINNumber")
		newElem.Text= sTinNumber
		Root.appendChild newElem

		if Trim(sGroupFlag)="" then
			sGroupFlag=0
			sGroup="N"
		end if

		Set newElem = PartyData.createElement("Group")
		newElem.setAttribute "Flag", sGroupFlag
		newElem.setAttribute "Type", sGroup
		newElem.setAttribute "ParentCompany",document.formname.hParentPartyCode.value

		Root.appendChild newElem

		Set newElem = PartyData.createElement("Units")
		Root.appendChild newElem
		for iCounter=0 to UBound(arrUnit)
			if Trim(sOwnUnit) <> Trim(arrUnit(iCounter)) then
				Set newElem1 = PartyData.createElement("UN")
				newElem1.setAttribute "Code", arrUnit(iCounter)
				newElem1.setAttribute "Name", arrUnitName(iCounter)
				newElem.appendChild newElem1
			end if
		next

		set newElem = PartyData.createElement("Active")
		newElem.setAttribute "Flag",sInActive
		Root.appendChild newElem
	end if 'if Root.hasChildNodes then

	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.open "POST","XMLSaveParty.asp?Name=Party&Mod=Master",false
	objhttp.send PartyData.xmldocument

'	alert(document.formname.hAction.value)

	if trim(document.formname.hAction.value)="EDIT" then

		set objhttp = CreateObject("MSXML2.XMLHTTP")
		objhttp.open "GET","ParGetUnitDetails.asp",false
		objhttp.send
		if trim(objhttp.responseXML.xml)<>"" then
			TempData.loadXML(objhttp.responseXML.xml)
		else
			alert(objhttp.responseText)
		end if

		set Root = PartyData.documentElement
		if Root.hasChildNodes() then
			for each ndChild in Root.childNodes
				if ndChild.nodeName="Units" then
					set ndUnits = ndChild
				end if
			next
		end if
		set ndRoot = TempData.documentElement
		if ndRoot.hasChildNodes() then
			for each ndChild in ndRoot.childNodes
				if ndChild.nodeName="Partytype" then
					for each ndPartyData in ndUnits.childNodes
						if ndPartyData.getAttribute("Code")=ndChild.getAttribute("Unit") then
							'ndChild.removeAttribute("Unit")
							ndPartyData.appendChild ndChild
						end if
					next
				elseif ndChild.nodeName="Agent" then
					Root.appendChild ndChild
				end if
			next
		end if

		set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.open "POST","XMLSaveParty.asp?Name=Party&Mod=Master",false
		objhttp.send PartyData.xmldocument

	end if 'if trim(document.formname.hAction.value)="Edit" then

End Function
'********************************
Function SaveXMLFinal()

	Dim sName,sShortName,sAddress1,sAddress2,sPincode,sCity,sState
	Dim sCountry,sPhone,sMobile,sFax,sEmail,sWebsite,sECCNo,sPanNo
	Dim sSalesLocal,sSalesCentral,sGroupFlag,sGroup,sTinNumber,sOwnUnit,sInActive

	Dim Root,newElem,objhttp,ndChild,ndRoot,ndUnits,ndPartyData,ndAgentData

	sName= document.formname.txtPartyName.value
	sShortName= document.formname.txtShortName.value
	sAddress1= document.formname.txtAddress1.value
	sAddress2= document.formname.txtAddress2.value
	sPincode= document.formname.txtPinCode.value
	sCity= document.formname.txtCity.value
	sState= document.formname.txtState.value
	sCountry= document.formname.txtCountry.value
	sPhone= document.formname.txtPhone.value
	sMobile=document.formname.txtMobileNo.value
	sFax= document.formname.txtFax.value
	sEmail= document.formname.txtEmail.value
	sWebsite= document.formname.txtWebsite.value
	sECCNo= document.formname.txtECCNo.value
	sPanNo= document.formname.txtPanNo.value
	sSalesLocal= document.formname.txtSalesLocal.value
	sSalesCentral= document.formname.txtSalesCentral.value
	iCreatedBy = document.formname.hCreatedBy.value

	if document.formname.chkGroupCompany.checked =true then
		sGroupFlag = document.formname.chkGroupCompany.value
	else
		sGroupFlag = ""
	end if

	'sGroupFlag= document.formname.chkGroupCompany.value
	if document.formname.radGroupType(0).checked=true then
		sGroup = document.formname.radGroupType(0).value
	elseif document.formname.radGroupType(1).checked=true then
		sGroup = document.formname.radGroupType(1).value
	elseif document.formname.radGroupType(2).checked=true then
		sGroup = document.formname.radGroupType(2).value
	end if
'	sGroup=trim(Request.Form("radGroupType"))
	sTinNumber =document.formname.txtTinNo.value
	sOwnUnit = document.formname.hOwnUnit.value
	sInActive =document.formname.hInActive.value

	Temp=trim(document.formname.hUnitCode.value)
	arrUnit=Split(Temp,":")

	'Response.Write Temp &"<br>"
	Temp=trim(document.formname.hUnitName.value)
	arrUnitName=Split(Temp,":")

	'Response.Write Temp &"<br>"
	iPartyCode = document.formname.hPartyCode.value

	if Trim(sGroupFlag)="" then
		sGroupFlag=0
		sGroup="N"
	end if
'	alert(sGroupFlag)
'	alert(sGroup)


	Set Root = PartyData.documentElement
	if Root.hasChildNodes then
		for each ndChild in Root.childNodes
			if ndChild.nodeName="ParCode" then
				ndChild.Text= iPartyCode
			elseif ndChild.nodeName="ParName" then
				ndChild.Text= sName
			elseif ndChild.nodeName="ShortName" then
				ndChild.Text= sShortName
			elseif ndChild.nodeName="Address1" then
				ndChild.Text= sAddress1
			elseif ndChild.nodeName="Address2" then
				ndChild.Text= sAddress2
			elseif ndChild.nodeName="PinCode" then
				ndChild.Text= sPincode
			elseif ndChild.nodeName="City" then
				ndChild.Text= sCity
			elseif ndChild.nodeName="State" then
				ndChild.Text= sState
			elseif ndChild.nodeName="Country" then
				ndChild.Text= sCountry
			elseif ndChild.nodeName="Phone" then
				ndChild.Text= sPhone
			elseif ndChild.nodeName="Mobile" then
				ndChild.Text= sMobile
			elseif ndChild.nodeName="Fax" then
				ndChild.Text= sFax
			elseif ndChild.nodeName="Email" then
				ndChild.Text= sEmail
			elseif ndChild.nodeName="Website" then
				ndChild.Text= sWebsite
			elseif ndChild.nodeName="ECCNo" then
				ndChild.Text= sECCNo
			elseif ndChild.nodeName="PANNo" then
				ndChild.Text= sPanNo
			elseif ndChild.nodeName="CreatedBy" then
				ndChild.Text= iCreatedBy
			elseif ndChild.nodeName="OwnUnit" then
				ndChild.Text= sOwnUnit
			elseif ndChild.nodeName="Sales" then
				ndChild.setAttribute "Local", sSalesLocal
				ndChild.setAttribute "Central", sSalesCentral
			elseif ndChild.nodeName="TINNumber" then
				ndChild.Text= sTinNumber
			elseif ndChild.nodeName="Group" then
				ndChild.setAttribute "Flag", sGroupFlag
				ndChild.setAttribute "Type", sGroup
				ndChild.setAttribute "ParentCompany", document.formname.hParentPartyCode.value
			elseif ndChild.nodeName="Active" then
				ndChild.setAttribute "Flag",sInActive
			end if
		next
	else
		Set newElem = PartyData.createElement("ParCode")
		newElem.Text= iPartyCode
		Root.appendChild newElem

		Set newElem = PartyData.createElement("ParName")
		newElem.Text= sName
		Root.appendChild newElem

		Set newElem = PartyData.createElement("ShortName")
		newElem.Text= sShortName
		Root.appendChild newElem


		Set newElem = PartyData.createElement("Address1")
		newElem.Text= sAddress1
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Address2")
		newElem.Text= sAddress2
		Root.appendChild newElem

		Set newElem = PartyData.createElement("PinCode")
		newElem.Text= sPincode
		Root.appendChild newElem

		Set newElem = PartyData.createElement("City")
		newElem.Text= sCity
		Root.appendChild newElem

		Set newElem = PartyData.createElement("State")
		newElem.Text= sState
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Country")
		newElem.Text= sCountry
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Phone")
		newElem.Text= sPhone
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Mobile")
		newElem.Text= sMobile
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Fax")
		newElem.Text= sFax
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Email")
		newElem.Text= sEmail
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Website")
		newElem.Text= sWebsite
		Root.appendChild newElem

		Set newElem = PartyData.createElement("ECCNo")
		newElem.Text= sECCNo
		Root.appendChild newElem

		Set newElem = PartyData.createElement("PANNo")
		newElem.Text= sPanNo
		Root.appendChild newElem

		Set newElem = PartyData.createElement("CreatedBy")
		newElem.Text= iCreatedBy
		Root.appendChild newElem

		Set newElem = PartyData.createElement("OwnUnit")
		newElem.Text= sOwnUnit
		Root.appendChild newElem

		Set newElem = PartyData.createElement("Sales")
		newElem.setAttribute "Local", sSalesLocal
		newElem.setAttribute "Central", sSalesCentral
		Root.appendChild newElem

		Set newElem = PartyData.createElement("TINNumber")
		newElem.Text= sTinNumber
		Root.appendChild newElem

		if Trim(sGroupFlag)="" then
			sGroupFlag=0
			sGroup="N"
		end if

		Set newElem = PartyData.createElement("Group")
		newElem.setAttribute "Flag", sGroupFlag
		newElem.setAttribute "Type", sGroup
		newElem.setAttribute "ParentCompany",document.formname.hParentPartyCode.value

		Root.appendChild newElem

		Set newElem = PartyData.createElement("Units")
		Root.appendChild newElem
		for iCounter=0 to UBound(arrUnit)
			if Trim(sOwnUnit) <> Trim(arrUnit(iCounter)) then
				Set newElem1 = PartyData.createElement("UN")
				newElem1.setAttribute "Code", arrUnit(iCounter)
				newElem1.setAttribute "Name", arrUnitName(iCounter)
				newElem.appendChild newElem1
			end if
		next

		set newElem = PartyData.createElement("Active")
		newElem.setAttribute "Flag",sInActive
		Root.appendChild newElem
	end if 'if Root.hasChildNodes then

	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.open "POST","XMLSaveParty.asp?Name=Party&Mod=Master",false
	objhttp.send PartyData.xmldocument



	if trim(document.formname.hAction.value)="EDIT" and trim(document.formname.hParUnit.value="N")  then

		set objhttp = CreateObject("MSXML2.XMLHTTP")
		objhttp.open "GET","ParGetUnitDetails.asp",false
		objhttp.send
		if trim(objhttp.responseXML.xml)<>"" then
			TempData.loadXML(objhttp.responseXML.xml)
		else
			alert(objhttp.responseText)
		end if

		set Root = PartyData.documentElement
		if Root.hasChildNodes() then
			for each ndChild in Root.childNodes
				if ndChild.nodeName="Units" then
					set ndUnits = ndChild
				end if
			next
		end if
		set ndRoot = TempData.documentElement
		if ndRoot.hasChildNodes() then
			for each ndChild in ndRoot.childNodes
				if ndChild.nodeName="Partytype" then
					for each ndPartyData in ndUnits.childNodes
						if ndPartyData.getAttribute("Code")=ndChild.getAttribute("Unit") then
							ndPartyData.appendChild ndChild
						end if
					next
				elseif ndChild.nodeName="Agent" then
					Root.appendChild ndChild
				end if
			next
		end if

		set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.open "POST","XMLSaveParty.asp?Name=Party&Mod=Master",false
		objhttp.send PartyData.xmldocument

	end if 'if trim(document.formname.hAction.value)="Edit" then

End Function

'**********************************


Function Fun_Contact(sPartyCode)
    if trim(document.formname.hPartyCode.value)<>"" then
	    sValue=	showModalDialog("ParContactPopup.asp?PartyCode="&sPartyCode,"","dialogHeight:440px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No")
	    if trim(sValue)="Done" then
		    document.formname.submit
	    end if
    else
        alert("Save the Basic Party Details before entering Contact Details")
        exit function
    end if
End Function
'************************************
Function Fun_Location(sPartyCode)
    if trim(document.formname.hPartyCode.value)<>"" then
	    sValue = showModalDialog("ParLocationPopup.asp?PartyCode="&sPartyCode,"","dialogHeight:440px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No")
	    if trim(sValue)="Done" then
		    document.formname.submit
	    end if
	else
	    alert("Save the Basic Party Details before entering Location Details")
	    exit function
	end if
End Function
'*****************************************
Function Fun_Preference(sPartyCode)
    if trim(document.formname.hPartyCode.value)<>"" then

	    sValue = showModalDialog("ParPerferencePopup.asp?PartyCode="&sPartyCode,"","dialogHeight:440px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No")
	    if trim(sValue)="Done" then
		    document.formname.submit
	    end if
    else
        alert("Save the Basic Party Details before entering Preference Details")
        exit function
    end if
End Function
'************************************
Function Fun_Agent(sPartyCode)
    if trim(document.formname.hPartyCode.value)<>"" then
	    sValue = showModalDialog("ParAgentSelectPopup.asp?PartyCode="&sPartyCode,"","dialogHeight:440px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No")
	    if sValue="Done" then
		    document.formname.submit
	    end if
	else
	    alert("Save the Basic Party Details before entering Agent Details")
	    exit function
	end if
End Function
'*****************************
Function GetUnit()
		Dim sUnit, arrUnit,TempNode,Root,sExp

		if document.formname.ChkOwnUnit.checked then
			sUnit = showModalDialog("ParCreationUnitSelPopup.asp","","dialogHeight:240px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
			if Trim(sUnit) <> "" then

				arrUnit = Split(sUnit,":")
				document.formname.hOwnUnit.value = arrUnit(0)
				document.formname.txtPartyName.value = document.formname.horgName.value&" - "& arrUnit(1)
				document.formname.txtPartyName.readOnly = True
				Set Root = UNITDET.documentElement
				sExp = "//UNIT[@UnitID="&arrUnit(0)&"]"
				Set TempNode = Root.selectNodes(sExp)
				IF TempNode.length <> 0 Then
					document.formname.txtAddress1.value = TempNode.Item(0).Attributes.getNamedItem("Add1").Value
					document.formname.txtAddress2.value = TempNode.Item(0).Attributes.getNamedItem("Add2").Value
					document.formname.txtCity.value = TempNode.Item(0).Attributes.getNamedItem("City").Value
					document.formname.txtPinCode.value = TempNode.Item(0).Attributes.getNamedItem("PostCode").Value
					document.formname.txtState.value = TempNode.Item(0).Attributes.getNamedItem("State").Value
					document.formname.txtCountry.value = TempNode.Item(0).Attributes.getNamedItem("Country").Value
					document.formname.txtPhone.value = TempNode.Item(0).Attributes.getNamedItem("Phone").Value
					document.formname.txtFax.value = TempNode.Item(0).Attributes.getNamedItem("Fax").Value
					document.formname.txtEmail.value = TempNode.Item(0).Attributes.getNamedItem("EmailID").Value
					document.formname.txtWebsite.value = TempNode.Item(0).Attributes.getNamedItem("Web").Value
				End IF

			else
				document.formname.txtPartyName.readOnly = False
				Cleartxt()
			end if
		Else
			document.formname.txtPartyName.readOnly = False
			Cleartxt()
		end if
	End Function

	Function Cleartxt()
		document.formname.txtPartyName.value = ""
		document.formname.txtAddress1.value = ""
		document.formname.txtAddress2.value = ""
		document.formname.txtCity.value = ""
		document.formname.txtPinCode.value = ""
		document.formname.txtState.value = ""
		document.formname.txtCountry.value = ""
		document.formname.txtPhone.value = ""
		document.formname.txtFax.value = ""
		document.formname.txtEmail.value = ""
		document.formname.txtWebsite.value = ""

	End Function
'*******************************************************************
FUNCTION popPartyDet(sTemp)
Dim sUnit,iCtr,sPartyGType,sType,Temparr,iCount,iUnitrow,Pararr,iCounter,sUseable
	if trim(sTemp)<>"" then
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		'Msgbox sTemp
		objhttp.Open "GET","XMLGetPartyDet.asp?PartyCode=" &sTemp , false
		objhttp.send
		'alert objhttp.responseText
		iUnitrow = document.formname.hUnitrow.value
		if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
			Set Root = OutData.documentElement
			'Msgbox Root.xml
			For Each PartyNode In Root.childNodes
				document.formname.txtShortName.value=PartyNode.Attributes.getNamedItem("OrgnPartyCode").value
				document.formname.txtAddress1.value=PartyNode.Attributes.getNamedItem("AddressLine1").value
				document.formname.txtAddress2.value=PartyNode.Attributes.getNamedItem("AddressLine2").value
				document.formname.txtCity.value=PartyNode.Attributes.getNamedItem("City").value
				document.formname.txtState.value=PartyNode.Attributes.getNamedItem("State").value
				document.formname.txtCountry.value=PartyNode.Attributes.getNamedItem("Country").value
				document.formname.txtPhone.value=PartyNode.Attributes.getNamedItem("PhoneNos").value
				document.formname.txtMobileNo.value=PartyNode.Attributes.getNamedItem("MobileNos").value
				document.formname.txtFax.value=PartyNode.Attributes.getNamedItem("FaxNos").value
				document.formname.txtEmail.value=PartyNode.Attributes.getNamedItem("Email").value
				document.formname.txtWebsite.value=PartyNode.Attributes.getNamedItem("WebsiteURL").value
				document.formname.txtPinCode.value=replace(PartyNode.Attributes.getNamedItem("Pincode").value," ","")
				document.formname.txtECCNo.value=PartyNode.Attributes.getNamedItem("ExciseControlCode").value
				document.formname.txtSalesLocal.value=PartyNode.Attributes.getNamedItem("LocalSTNoandDT").value
				document.formname.txtSalesCentral.value=PartyNode.Attributes.getNamedItem("CentralSTNoandDT").value
				document.formname.txtPanNo.value=PartyNode.Attributes.getNamedItem("IncomeTaxPANNo").value
				document.formname.txtPartyName.value = PartyNode.Attributes.getNamedItem("PartyName").value
				sUnit = PartyNode.Attributes.getNamedItem("Units").value
				sPartyGType = PartyNode.Attributes.getNamedItem("PartyGroupCoyType").value
				sType = PartyNode.Attributes.getNamedItem("InTrans").value
				document.formname.txtTinNo.value = PartyNode.Attributes.getNamedItem("TINNumber").value
				sUseable = PartyNode.Attributes.getNamedItem("Useable").value
				'alert(sUseable)
				if sUseable="1" then
					document.formname.chkActive.checked = true
				end if

			next

			Dim arrName,iUnitCountSel
			'MsgBox sUnit
			Temparr = Split(sUnit,":")

			IF UBound(Temparr) = 0 Then
				IF CStr(iUnitrow)<>"" Then
					For iCount = 0 To iUnitrow
						Temparr = Split(Cstr(eval("document.formname.chkUnitZ"&iCount).value),":")
						IF Cstr(Trim(sUnit)) = Cstr(Trim(Temparr(0))) Then
							eval("document.formname.chkUnitZ"&iCount).checked = True
							Exit For
						End IF
					Next
				Else
					Temparr = Split(Cstr(document.formname.chkUnitZ0.value),":")
					IF Cstr(Trim(sUnit)) = Cstr(Trim(Temparr(0))) Then
						document.formname.chkUnitZ0.checked = True
					End IF
				End IF
			Else

		'	alert(iUnitrow)
		'	alert(UBound(Temparr)+1)

				If cint(iUnitrow) = cint(UBound(Temparr)+1) then
					document.formname.chkUnitZ0.checked = true
				Else
					For iCtr = 0 To UBound(Temparr)
						sUnit = Cstr(Trim(Temparr(iCtr)))
						For iCount = 0 To iUnitrow
							arrName =  Split(eval("document.formname.chkUnitZ"&iCount).value,":")
							IF (sUnit = Cstr(Trim(arrName(0)))) Then
								eval("document.formname.chkUnitZ"&iCount).checked = True
							End IF
						Next
					Next
				End If
			End IF

'			alert(sType)

			Pararr = Split(sType,":")


			IF CStr(sType) <> "0" Then
				IF CStr(iUnitrow) <> "" Then
					For iCount = 0 To iUnitrow
						For iCounter = 1 To UBound(Pararr)
							unitarr = Split(eval("document.formname.chkUnitZ"&iCount).value,":")
							IF Unitarr(0) = Pararr(iCounter) and eval("document.formname.chkUnitZ"&iCount).checked = True Then
								eval("document.formname.chkUnitZ"&iCount).disabled = True
							End IF
						Next
					Next
				Else
					unitarr = Split(document.formname.chkUnitZ0.value,":")
					IF document.formname.chkUnitZ0.checked = True Then
						document.formname.chkUnitZ0.disabled = True
					End IF
				End IF

			End IF
		end if

		IF CStr(sPartyGType) = "P" Then
			document.formname.chkGroupCompany.checked = True
			document.formname.radGroupType(0).checked = True
			document.formname.radGroupType(0).disabled = False
			document.formname.radGroupType(1).disabled = False
			document.formname.radGroupType(2).disabled = False
		Elseif CStr(sPartyGType) = "C" Then
			document.formname.chkGroupCompany.checked = True
			document.formname.radGroupType(1).checked = True
			document.formname.radGroupType(0).disabled = False
			document.formname.radGroupType(1).disabled = False
			document.formname.radGroupType(2).disabled = False
		Elseif CStr(sPartyGType) = "B" Then
			document.formname.chkGroupCompany.checked = True
			document.formname.radGroupType(2).checked = True
			document.formname.radGroupType(0).disabled = False
			document.formname.radGroupType(1).disabled = False
			document.formname.radGroupType(2).disabled = False
		End IF
	end if' if trim(sTemp)<>"" then
END FUNCTION

Function CheckUnit()
	Dim sTemp,iCtr,Temparr,sUnitid,sUnitName
	sTemp = document.formname.hUnitrow.value


	IF CStr(sTemp) <> "0" Then
		IF document.formname.chkUnitZ0.checked = True Then
			For iCtr = 1 To sTemp
				Temparr = Split(eval("document.formname.chkUnitZ"&iCtr).Value,":")
				sUnitid = sUnitid &":"&Temparr(0)
				sUnitName = sUnitName &":"&Temparr(1)
			Next
		Else
			For iCtr = 1 To sTemp
				IF eval("document.formname.chkUnitZ"&iCtr).checked = True Then
					Temparr = Split(eval("document.formname.chkUnitZ"&iCtr).Value,":")
					sUnitid = sUnitid &":"&Temparr(0)
					sUnitName = sUnitName &":"&Temparr(1)
				End IF
			Next
		End if
	Else
		IF document.formname.chkUnitZ0.checked = True Then
			For iCtr = 1 To sTemp
				Temparr = Split(eval("document.formname.chkUnitZ"&iCtr).Value,":")
				sUnitid = sUnitid &":"&Temparr(0)
				sUnitName = sUnitName &":"&Temparr(1)
			Next
		End IF
	End IF

	IF CStr(sUnitid) <> "" Then
		sUnitid = Right(sUnitid,Len(sUnitid)-1)
		sUnitName = Right(sUnitName,len(sUnitName)-1)

		document.formname.hUnitCode.value = sUnitid
		document.formname.hUnitName.value = sUnitName
	Else
		document.formname.hUnitCode.value = ""
		document.formname.hUnitName.value = ""
	End IF

	if document.formname.chkActive.checked = true then
		document.formname.hInActive.value = "1"
	end if

End Function

Function PageSubmit()
Dim sFlag,sUnitFlag
	sFlag = false
	sUnitFlag = false
	CheckUnit()
	IF Not CheckForm() Then
		Exit Function
	Else
		SaveXMLFinal

		set Root = PartyData.documentElement
	'	alert(Root.xml)
		If Root.hasChildNodes() then
			for each ndChild in Root.childNodes
				if ndChild.nodeName="Units" then
				    sUnitFlag = true
					for each ndChild1 in ndChild.childNodes
						if ndChild1.hasChildNodes() then
							sFlag = true
						else
							sFlag = false
						end if
					next
				end	if  'if ndChild.nodeName="Units" then
			next
		End If

		if not sFlag then
			alert("Enter Party Types for the Units")
			exit function
		end if

		IF CStr(document.formname.hUnitCode.value) = "0" Then
			MsgBox "Select any Unit "
			Exit Function
		Else
			if trim(document.formname.hPartyCode.value)<>"" then
				document.formname.action = "ParCreate_Edit_EntryInsert.asp?Action=Edit&PartyCode="&document.formname.hPartyCode.value
			else
				document.formname.action = "ParCreate_Edit_EntryInsert.asp?Action=Insert"
			end if  'if trim(document.formname.hPartyCode.value)<>"" then
			document.formname.B2.disabled = True
			document.formname.B3.disabled = True
			document.formname.submit
		End IF
	End IF
End Function

Function CheckForm()
	'alert(document.formname.hUnitCode.value)
	IF document.formname.txtShortName.value = "" Then
		MsgBox "Enter Party Code "
		document.formname.txtShortName.focus
		CheckForm = False
	Elseif document.formname.txtCity.value = "" Then
		MsgBox "Enter City "
		document.formname.txtCity.focus
		CheckForm = False
	Elseif document.formname.hUnitCode.value ="" then
		MsgBox "Selec the Units"
		CheckForm = False
	Else
		CheckForm = True
	End IF

End Function
'*************************************************
Function Fun_Rep(sPartyCode)
    if trim(document.formname.hPartyCode.value)<>"" then
	    sValue = showModalDialog("RepSelectionEntry.asp?PartyCode="&sPartyCode,"","dialogHeight:250px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
	    if sValue="Done" then
		    document.formname.submit
	    end if
	else
	    alert("Save the Basic Party Details before entering Rep. Details")
	    exit function
	end if
End Function
</SCRIPT>
<SCRIPT LANGUAGE=javascript>
<!--
function EnableGroup(objGroup)
{
	if (document.formname.chkGroupCompany.checked==true)
	{

		document.formname.radGroupType[0].disabled=false;
		document.formname.radGroupType[1].disabled=false;
		document.formname.radGroupType[2].disabled=false;
	}
	else
	{
		document.formname.radGroupType[0].disabled=true;
		document.formname.radGroupType[1].disabled=true;
		document.formname.radGroupType[2].disabled=true;
	}
}
function CheckSubmit()
{

	var i,bFalg;

	bFlag=true;
	if (trim(document.formname.txtShortName.value) =="")
	{
		alert("Enter Party Code");
		document.formname.txtShortName.select();
		return false;
	}
	if (trim(document.formname.txtCity.value) =="")
	{
		alert("Enter Party City");
		document.formname.txtCity.select();
		return false;
	}
	return true;
}
//-->
</SCRIPT>

<script language="javascript">
window.__itmsPopupCompat = { type: "partyCreateEditModals" };
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="popPartyDet('<%=iParty%>')">
<form method="POST" name="formname">
<input type="Hidden" name="hUnitName" value="">
<input type="Hidden" name="hUnitCode" value="<%=sOrgCode%>">
<input type="Hidden" name="hPartyCode" value="<%=iParty%>">
<input type="Hidden" name="hOwnUnit" value="">
<input type="Hidden" name="hAction" value="<%=sAction%>">
<input type="hidden" name="hInActive" value="0">
<input type="hidden" name="hCreatedBy" value="<%=getUserID%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hParentPartyCode" value="<%=sGroupPartyCode%>">
<input type="hidden" name="hParUnit" value="N">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">
		<%
			if trim(iParty)<>"" then
				Response.Write "Party Amendment"
			else
				Response.Write "Party Creation"
			end if
		%>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
			<TR>
					<td height="20px" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCurrentCell" valign="bottom" width="60px">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Details
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="60px">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)" height="13px">
										<tr>
											<td align="center">
												<a href="#" onClick="ControlData()">Control</a>
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="60px">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)" height="13px">
										<tr>
											<td align="center">
												<a href="#" onClick="ViewData()">View</a>
											</td>
										</tr>
									</table>
								</td>
								<!--<td class="TabCell" valign="bottom" align="center" width="60">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr>
									  <td align="center">Group</td>
									</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="72">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr>
									  <td align="center">Contact</td>
									</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="78">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr>
									  <td align="center">Location</td>
									</tr>
								  </table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="92">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								   <tr>
									  <td align="center">Preference</td>
									</tr>
								  </table>
								</td>-->
								<td class="TabCellEnd" valign="bottom" align="left">
                                &nbsp;
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<TR>
					<TD class="TabBody">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
							</tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=FieldCell width="115px"> Party Name</td>
															<td class="FieldCell">
															<Input type="text" size="71" name="txtPartyName" value="" class="FormElem">&nbsp;
															<input type="checkbox" name="chkActive" value="1" class="FormElem">&nbsp;In-Active
                                                                </td>
														</tr>
														<tr>
															<td class=FieldCell width="115px"> Party Code</td>
															<td class='FieldCell' valign=top><input type="text" name="txtShortName" size="12" maxlength="10" class="Formelem">
															<input type=CheckBox name=ChkOwnUnit class=Formelem onClick=GetUnit()>Own Unit
														</tr>
													</table>
								</td>
								<td align="center">
								</td>
							</tr>
							<tr>
								<td align="center"></td>
								<td valign="top" width="100%">
									<table border="0" cellspacing="0" cellpadding="0">
									                                                        <tr>
									                                                        	<td width="115px">&nbsp;</td>
									                                                    <td class="FieldCell" rowspan="3" valign="top">
									                                                    <input type="checkbox" name="chkGroupCompany" value="1" onClick="EnableGroup(this)" class="FormElem"> Group Company</td>
									                                                    <td class="FieldCellSub">
									                                                     <input type="radio" value="P" name="radGroupType" disabled="true" checked class="FormElem" onClick="GetGroup()"> Parent </td>

									                                                    <td class="FieldCellSub"><input type="radio" value="C" disabled="true" name="radGroupType" class="FormElem" onClick="GetGroup()" > Child </td>

									                                                    <td class="FieldCellSub"><input type="radio" value="B" disabled="true" name="radGroupType" class="FormElem" onClick="GetGroup()" > Parent / Child </td>
									                                                        </tr>
									                                                        <tr>
																								<td class="FieldCell" colspan=5>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" class="FormElemRead" name="ParentPartyName" size=75 readonly value="<%=sGroupParentName%>" >&nbsp;</span></td>
									                                                        </tr>
                                                            </table>
								</td>
							</td>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top">
												<center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class="GroupTitleLeft" width="10px">&nbsp;
                                                            </td>
															<td class="GroupTitle" width="60px"><p align="center">Address
                                                            </td>
												</center>
															<td class="GroupTitleRight"><p align="left">&nbsp;
                                                            </td>
														</tr>
													</table>
                                                        </td>
														</tr>
														<tr>
															<td class="GroupTable">
												<center>
                                                    <div align="left">
                                        <table cellpadding="0" cellspacing="0">
                                          <tr>
                                            <td class="MiddlePack" colspan="5"></td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub">Address
                                            </td>
                                            <td class="FieldCellSub" colspan="4"><input type="text" name="txtAddress1" size="81" class="FormElem">
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCellSub" colspan="4"><input type="text" name="txtAddress2" size="81" class="FormElem">
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub">City
                                            </td>
                                            <td class="FieldCellSub" colspan="4"><input type="text" name="txtCity" size="25" class="FormElem">
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub">PIN
                                            </td>
                                            <td class="FieldCellSub"><input type="text" name="txtPinCode" size="7"  maxlength="6" class="FormElem">
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub">Phone
                                          </td>
                                          <td class="FieldCellSub"><input type="text" name="txtPhone" size="18" class="FormElem">
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub">State
                                            </td>
                                            <td class="FieldCellSub"><input type="text" name="txtState" size="35" class="FormElem">
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub">Fax
                                          </td>
                                          <td class="FieldCellSub"><input type="text" name="txtFax" size="18" class="FormElem">
                                          </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub">Country
                                            </td>
                                            <td class="FieldCellSub"><input type="text" name="txtCountry" size="25" class="FormElem">
                                            </td>
                                            <td class="FieldCellSub">
                                            </td>
                                          <td class="FieldCellSub">Mobile
                                          </td>
                                          <td class="FieldCellSub"><input type="text" name="txtMobileNo" size="18" class="FormElem">
                                          </td>
                                          </tr>
                                          <tr>
                                          <td class="FieldCellSub">E-mail ID
                                          </td>
                                          <td class="FieldCellSub"><input type="text" name="txtEmail" size="35" class="FormElem">
                                          </td>
                                          <td class="FieldCellSub">
                                          </td>
                                          <td class="FieldCellSub">URL
                                          </td>
                                          <td class="FieldCellSub"><input type="text" name="txtWebsite" size="25" class="FormElem">
                                          </td>
                                          </tr>
                                        <tr>
                                          <td class="MiddlePack" colspan="5"></td>
                                        </tr>
                                        </table>
                                                    </div>
												</center>
                                                            </td>
														</tr>
													</table>
								</td>
								<td align="center">
								</td>
							</tr>
                                <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top">
                                    <table border="0" cellspacing="0" class="BodyTable" cellpadding="0" width="100%">
                                    	<tr>
                                           <td>
                                               <table border="0" cellpadding="0" cellspacing="0">
                                               <tr>
                                            		<td>
                                                            <table border="0" cellspacing="0" cellpadding="0">
                                                        		<tr>
																	<td class="FieldCellSub" width="165">Excise ECC Number</td>
																	<td class="FieldCellSub"> <input type="text" name="txtECCNo" size="17" class="FormElem"> </td>
																	<td class="FieldCellSub" >IT PAN No</td>
																	<td class="FieldCellSub"> <input type="text" name="txtPanNo" size="15" class="FormElem"> </td>
																</tr>
																<tr>
																	<td class="FieldCellSub" >Sales Tax Number - Local</td>
																	<td class="FieldCellSub"> <input type="text" name="txtSalesLocal" size="17" class="FormElem"> </td>
																	<td class="FieldCellSub" >TIN Number</td>
																	<td class="FieldCellSub"> <input type="text" name="txtTinNo" size="15" class="FormElem"> </td>
																</tr>
																<tr>
																	<td class="FieldCellSub" >Sales Tax Number - Central</td>
																	<td class="FieldCellSub"> <input type="text" name="txtSalesCentral" size="17" class="FormElem"> </td>
																	<td class="FieldCellSub" >&nbsp;</td>
																	<td class="FieldCellSub">

																	</td>
																</tr>

                                                            </table>
                                            		</td>
                                            		<!--td class="ClearPixel">
														<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                                            		</td>
													<td valign="top" align="right">
																	<table border="0" cellspacing="0" cellpadding="0">
																<tr>
															<td class="FieldCellSub" rowspan="3" valign="top">
															<input type="checkbox" name="chkGroupCompany" value="1" onClick="EnableGroup(this)" class="FormElem"> Group Company</td>
															<td class="FieldCellSub">
															 <input type="radio" value="P" name="radGroupType" disabled="true" checked class="FormElem"> Parent </td>
																</tr>
																<tr>
															<td class="FieldCellSub"><input type="radio" value="C" disabled="true" name="radGroupType" class="FormElem"> Child </td>
																</tr>
																<tr>
															<td class="FieldCellSub"><input type="radio" value="B" disabled="true" name="radGroupType" class="FormElem"> Parent / Child </td>
																</tr>
																	</table>
													</td-->
                                                </tr>
                                              	</table>
                                              </td>
                                           </tr>
                                     </table>
								</td>
								<td align="center">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
							</tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top">
									<table border="0" cellspacing="0" class="BodyTable" cellpadding="0" width="100%">
										<tr>
											<td class="ExcelHeaderCell" height="20px" >Transaction Units  (<b><a href="#" onClick="PopulatePartyTypes('<%=iParty%>','<%=sAction%>')" >Party Types</a></b>) </td>
											<td class="ExcelHeaderCell" >Contacts</td>
											<td class="ExcelHeaderCell" >Locations</td>
											<td class="ExcelHeaderCell" >Preference</td>
											<td class="ExcelHeaderCell" >Agent</td>
											<td class="ExcelHeaderCell" >Rep.</td>
										</tr>
										<tr>
											<td class="FieldCellSub">
												<%
													sQuery = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID "

													'Response.Write sQuery
													With objRs
														.CursorLocation = 3
														.CursorType = 3
														.Source = sQuery
														.ActiveConnection = Con
														.Open
													End With
													Set objRs.ActiveConnection = Nothing
												%>
												<input type="hidden" name="hUnitrow" value="<%=objRs.RecordCount%>">
												<input type="checkbox" class="FormElem" name="chkUnitZ0" value="0:All">All Units&nbsp;
													<br>
												<%
													Do while Not objRs.EOF
														iUnitCount = iUnitCount + 1
												%>
													<input type="checkbox" class="FormElem" name="chkUnitZ<%=iUnitCount%>" value="<%=objRs(0)%>:<%=Trim(objRs(1))%>"><%=Trim(objRs(1))%> &nbsp;
													<br>
												<%
													objRs.MoveNext
													Loop
													objRs.Close
												%>
											</td>
											<td class="FieldCellSub"><a href="#" onClick="Fun_Contact('<%=iParty%>')" >
											<%
												if iParty<>"" then
													sQuery = "Select count(PartyCode) from APP_M_PartyContactPersons where PartyCode = "& iParty
													objRs.Open sQuery,con
													if not objRs.EOF then
														iContactCount = objRs(0)
													end if
													objRs.Close
												else
													iContactCount = 0
												end if 'if iParty<>"" then

												if iContactCount=0 then
													Response.Write "0 Add"
												else
													Response.Write iContactCount &" Manage"
												end if

											%>
											</a>
											</td>
											<td class="FieldCellSub"><a href="#" onClick="Fun_Location('<%=iParty%>')">
											<%
												if iParty<>"" then
													sQuery= "Select Count(PartyCode) from APP_M_PartyLocations where PartyCode = "&iParty
													objRs.Open sQuery,con
													if not objRs.EOF then
														iLocationCount = objRs(0)
													end if
													objRs.Close
												else
													iLocationCount = 0
												end if 'if iParty<>"" then

												if iLocationCount=0 then
													Response.Write "0 Add"
												else
													Response.Write iLocationCount&" Manage"
												end if
											%>
											</a>
											</td>
											<td class="FieldCellSub"><a href="#" onClick="Fun_Preference('<%=iParty%>')">
											<%
												if iParty<>"" then
													sQuery = "Select Count(partyCode) from APP_R_OrgParty where (isNull(PrefTransporterCode,0)<>0 OR "&_
															 "isNull(PrefDespatchMode,0)<>0 or isNull(PrefCurrencyCode,0)<>0 or "&_
															 "isNull(PrefPaymentMode,0)<>0 Or isNull(PrefBasisOfPricing,0)<>0 Or "&_
															 "isNull(PrefPaymentTerms,0)<>0) and PartyCode = "& iParty
													objRs.Open sQuery,con
													if not objRs.EOF then
														 iPerfCount = objRs(0)
													end if
													objRs.Close
												end if ' if iParty<>"" then

												if iPerfCount = 0 then
													Response.Write "0 Add"
												else
													Response.Write "1 Manage"
												end if
											%>
											</a>
											</td>
											<td class="FieldCellSub"><a href="#" onClick="Fun_Agent('<%=iParty%>')">
											<%
												if iParty<>"" then
													sQuery = "Select Count(AgentCode) from APP_R_AgentOrgParty where PartyCode = "&iParty
													objRs.Open sQuery,con
													if not objRs.EOF then
														iAgentCount = objRs(0)
													end if
													objRs.Close
												else
													iAgentCount = 0
												end if 'if iParty<>"" then
												if iAgentCount=0 then
													Response.Write iAgentCount & "  Add"
												else
														Response.Write iAgentCount  &"  Manage"
												end if ' if iAgentCount=0 then
											%>
											</a>
											</td>
											<td class="FieldCellSub"><a href="#" onClick="Fun_Rep('<%=iParty%>')">
											<%
												if iParty<>"" then
													sQuery = "Select isNull(RepAreaCode,0),isNull(RepAgentEntryID,0) from APP_R_OrgParty where PartyCode = "& iParty
													objRs.Open sQuery,con
													if not objRs.EOF then
													    if Trim(objRs(0))<>"0" and Trim(objRs(1))<>"0" then
													       iRepCount  = 1
													    end if 'if Trim(objRs(0))<>"0" and Trim(objRs(1))<>"0" then
													end if
													objRs.Close
												else
													iRepCount = 0
												end if 'if iParty<>"" then
												if iRepCount=0 then
													Response.Write "Add"
												else
													Response.Write "Manage"
												end if ' if iAgentCount=0 then
											%>
											</a>
											</td>
										</td>
									</table>
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
							</tr>
							<tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td class="ActionCell">
                                                                <input type="button" value="Save" name="B2" class="ActionButton" onClick="PageSubmit()">
                                                                <input type="button" value="Close" name="B3" class="ActionButton"  onClick="GoToMain()">
                                                                <input type="button" value="Preview" name="btnPreveiw" class="ActionButton"  onClick="ViewData()">
                                                                <input type="reset" value="Reset" name="B1" class="ActionButton"  >
														</td>
													</tr>
												</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>
<%
set objRs=nothing
%>
