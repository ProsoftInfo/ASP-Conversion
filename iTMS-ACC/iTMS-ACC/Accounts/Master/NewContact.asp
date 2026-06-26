<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	NewContact.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Kalaiselvi R
	'Created On					:	Sep 27,2011
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
dim sQuery,objRs,iContactNo,sCallTy,Temparr,sAction



Set objRs = Server.CreateObject("ADODB.RecordSet")

iContactNo = Request.QueryString("ContactNo")
'Response.Write "<p>iContactNo = " & iContactNo

sCallTy = Request("hCallTy")

if trim(iContactNo)="" then
	sAction = "CREATE"
else
	sAction = "EDIT"
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/trim.js"></SCRIPT>

<XML ID="OutData" ></XML>
<XML id="ContactData"><Root/></XML>

<XML ID="PartyData"><Party/></XML>

<SCRIPT LANGUAGE=vbscript>
'-------------------------------------------------------------------------------------------
Function DelParty()
	document.formname.hPartyName.value = ""
	document.formname.hParCode.value = ""
	spParty.innerText = ""
End Function
'-------------------------------------------------------------------------------------------
Function SelPartyPopup()
	Dim Root1,objhttp

	set Root1 = PartyData.documentElement

	set objhttp = CreateObject("MSXML2.XMLHTTP")

	sOrgId = document.formname.hUnitId.value

	document.formname.hPartyName.value = ""
	spParty.innerText = ""

	sPartyType = "0"

	Set OutValue = showModalDialog("../Transaction/PartySelMultipleSubType.asp?orgID="&sOrgId&"&Party="&sPartyType&"&hSelectMode=S",PartyData,"","dialogHeight:500px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")

	sAct = UCase(trim(OutValue.getAttribute("Action")))
	sQuery = trim(OutValue.getAttribute("PassQuery"))
	if ucase(trim(sAct)) <> "CLOSE" then
		do while sAct <> "DONE"
			Set OutValue = showModalDialog("../Transaction/PartySelMultipleSubType.asp?"&sQuery,PartyData,"dialogHeight:500px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
			sAct = UCase(trim(OutValue.getAttribute("Action")))
			if ucase(Trim(sAct)) = "CLOSE" then exit do
			sQuery = trim(OutValue.getAttribute("PassQuery"))
		loop
	end if

	If  OutValue.hasChildNodes Then
		sExp ="//PartyDetails"
		Set PartyNode = OutValue.Selectnodes(sExp)
		if PartyNode.Length > 0 then
			for itr = 0 to PartyNode.Length - 1
				sParVal = ""
				sParTy = sParTy &",'"& PartyNode.Item(itr).Attributes.getNamedItem("ParType").value&"'"
				sParSubType = sParSubType &","&  PartyNode.Item(itr).Attributes.getNamedItem("ParSubType").value
				sParCode =sParCode  &","& PartyNode.Item(itr).Attributes.getNamedItem("ParCode").value
				sPartyName = sPartyName  &","& PartyNode.Item(itr).Attributes.getNamedItem("ParName").value
			Next
		End IF
		objhttp.Open "POST","XMLSave.asp?Name=PartyType&Mod=SAL", false
		objhttp.send PartyData.XMLDocument

		document.formname.hPartyName.value = mid(sPartyName,2)
		document.formname.hParCode.value = mid(sParCode,2)
		'sParVal = mid(sParTy,2) &"?"&mid(sParSubType,2)&"?"&mid(sPartyName,2)&"?"&mid(sParCode,2)
		spParty.innerText = mid(sPartyName,2)

	End If
End Function
'-------------------------------------------------------------------------------------------
Function ViewData()
    if trim(document.formname.hContactNumber.value)<>"" then
	    document.formname.action = "ParDetailsView.asp?PartyCode="& document.formname.hContactNumber.value
	    document.formname.submit
	else
	    alert("Party Details Cannot view because Party is not available")
	    exit function
	end if
End Function
'-------------------------------------------------------------------------------------------
Function GoToMain()
	document.formname.action = "ContactsList.asp"
	document.formname.submit
End Function
'-------------------------------------------------------------------------------------------
Function SaveXMLFinal()

	Dim sName,sDesignation,sAddress1,sAddress2,sPincode,sCity,sState
	Dim sCountry,sPhone,sMobile,sFax,sEmail,sWebsite,sInActive
	Dim iPartyCode

	Dim Root,newElem,objhttp,ndChild,ndRoot

	sName= document.formname.txtContactName.value
	sDesignation= document.formname.txtDesignation.value
	sContactPersonFor= document.formname.txtContactPersonFor.value
	iPartyCode = document.formname.hParCode.value

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
	iCreatedBy = document.formname.hCreatedBy.value
	sInActive = 0
	if document.formname.chkActive.checked then
		sInActive = 1
	end if

	iContactNoCode = document.formname.hContactNumber.value


	Set Root = ContactData.documentElement
	if Root.hasChildNodes then
		for each ndChild in Root.childNodes
			if ndChild.nodeName="ParCode" then
				ndChild.Text= iContactNoCode
			elseif ndChild.nodeName="ParName" then
				ndChild.Text= sName
			elseif ndChild.nodeName="Designation" then
				ndChild.Text= sDesignation
			elseif ndChild.nodeName="ContactPersonFor" then
				ndChild.Text= sContactPersonFor
			elseif ndChild.nodeName="PartyCode" then
				ndChild.Text= iPartyCode
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
			elseif ndChild.nodeName="CreatedBy" then
				ndChild.Text= iCreatedBy
			elseif ndChild.nodeName="Active" then
				ndChild.setAttribute "Flag",sInActive
			end if
		next
	else
		Set newElem = ContactData.createElement("ParCode")
		newElem.Text= iContactNoCode
		Root.appendChild newElem

		Set newElem = ContactData.createElement("ParName")
		newElem.Text= sName
		Root.appendChild newElem

		Set newElem = ContactData.createElement("Designation")
		newElem.Text= sDesignation
		Root.appendChild newElem

		Set newElem = ContactData.createElement("ContactPersonFor")
		newElem.Text= sContactPersonFor
		Root.appendChild newElem

		Set newElem = ContactData.createElement("PartyCode")
		newElem.Text= iPartyCode
		Root.appendChild newElem

		Set newElem = ContactData.createElement("Address1")
		newElem.Text= sAddress1
		Root.appendChild newElem

		Set newElem = ContactData.createElement("Address2")
		newElem.Text= sAddress2
		Root.appendChild newElem

		Set newElem = ContactData.createElement("PinCode")
		newElem.Text= sPincode
		Root.appendChild newElem

		Set newElem = ContactData.createElement("City")
		newElem.Text= sCity
		Root.appendChild newElem

		Set newElem = ContactData.createElement("State")
		newElem.Text= sState
		Root.appendChild newElem

		Set newElem = ContactData.createElement("Country")
		newElem.Text= sCountry
		Root.appendChild newElem

		Set newElem = ContactData.createElement("Phone")
		newElem.Text= sPhone
		Root.appendChild newElem

		Set newElem = ContactData.createElement("Mobile")
		newElem.Text= sMobile
		Root.appendChild newElem

		Set newElem = ContactData.createElement("Fax")
		newElem.Text= sFax
		Root.appendChild newElem

		Set newElem = ContactData.createElement("Email")
		newElem.Text= sEmail
		Root.appendChild newElem

		Set newElem = ContactData.createElement("Website")
		newElem.Text= sWebsite
		Root.appendChild newElem

		Set newElem = ContactData.createElement("CreatedBy")
		newElem.Text= iCreatedBy
		Root.appendChild newElem

		set newElem = ContactData.createElement("Active")
		newElem.setAttribute "Flag",sInActive
		Root.appendChild newElem
	end if 'if Root.hasChildNodes then

	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.open "POST","XMLSaveParty.asp?Name=Contact&Mod=Master",false
	objhttp.send ContactData.xmldocument

End Function
'-------------------------------------------------------------------------------------------
Function Cleartxt()
	document.formname.txtContactName.value = ""
	document.formname.txtDesignation.value = ""
	document.formname.txtContactPersonFor.value = ""
	document.formname.hParCode.value = ""

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
'-------------------------------------------------------------------------------------------
FUNCTION popPartyDet(sTemp)

Dim iCtr,Temparr,iCount,iCounter,sUseable

	'Msgbox sTemp

	if trim(sTemp)<>"" then
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		objhttp.Open "GET","XMLGetContactData.asp?ContactNo=" &sTemp , false
		objhttp.send

		'alert objhttp.responseText

		if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
			Set Root = OutData.documentElement
			'Msgbox Root.xml
			For Each PartyNode In Root.childNodes
				document.formname.txtDesignation.value=PartyNode.Attributes.getNamedItem("Designation").value
				document.formname.txtContactPersonFor.value=PartyNode.Attributes.getNamedItem("ContactPersonFor").value
				document.formname.hParCode.value=PartyNode.Attributes.getNamedItem("PartyCode").value
				spParty.innerText = PartyNode.Attributes.getNamedItem("PartyName").value

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
				document.formname.txtContactName.value = PartyNode.text
				sUseable = PartyNode.Attributes.getNamedItem("Useable").value
				'alert(sUseable)
				if sUseable="1" then
					document.formname.chkActive.checked = true
				end if

			next


		end if


	end if' if trim(sTemp)<>"" then
END FUNCTION
'-------------------------------------------------------------------------------------------
Function PageSubmit()
Dim sFlag
	sFlag = false


	IF Not CheckForm() Then
		Exit Function
	Else
		SaveXMLFinal



		if trim(document.formname.hContactNumber.value)<>"" then
			document.formname.action = "ContactInsert.asp?Action=Edit"
		else
			document.formname.action = "ContactInsert.asp?Action=Create"
		end if

		document.formname.B2.disabled = True
		document.formname.B3.disabled = True
		document.formname.submit

	End IF
End Function
'-------------------------------------------------------------------------------------------
Function CheckForm()

	IF document.formname.txtContactName.value = "" Then
		MsgBox "Enter Contact Name"
		document.formname.txtContactName.focus
		CheckForm = False
	ElseIf document.formname.txtDesignation.value = "" Then
		MsgBox "Enter Designation"
		document.formname.txtDesignation.focus
		CheckForm = False
	Elseif document.formname.txtCity.value = "" Then
		MsgBox "Enter City "
		document.formname.txtCity.focus
		CheckForm = False
	Else
		CheckForm = True
	End IF

End Function
'-------------------------------------------------------------------------------------------
</SCRIPT>

<script language="javascript">
window.__itmsPopupCompat = { type: "newContact" };
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="popPartyDet('<%=iContactNo%>')">
<form method="POST" name="formname">

<input type="Hidden" name="hContactNumber" value="<%=iContactNo%>">
<input type="Hidden" name="hAction" value="<%=sAction%>">
<input type="hidden" name="hCreatedBy" value="<%=getUserID%>">

<Input type="hidden" name="hPartyName" value="">
<Input type="hidden" name="hParCode" value="">
<Input type="hidden" name="hUnitId" value="<%=Session("organizationcode")%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td align="center" class="PageTitle">
		<%
			if trim(iContactNo)<>"" then
				Response.Write "Contact Amendment"
			else
				Response.Write "Contact Creation"
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
			<TABLE id="Table16" cellSpacing="0" cellPadding="0" border="0" width="100%">
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
															<td class="FieldCell" width="115px"> Contact Name</td>
															<td class="FieldCell">
																<Input type="text" size="71" name="txtContactName" value="" class="FormElem">&nbsp;
																<input type="checkbox" name="chkActive" value="1" class="FormElem">&nbsp;In-Active
                                                            </td>
														</tr>
														<tr>
															<td class="FieldCell" width="115px">Designation</td>
															<td class="FieldCell">
																<Input type="text" size="51" name="txtDesignation" value="" class="FormElem">
                                                            </td>
														</tr>

														<tr>
															<td class=FieldCell width="115px">ContactPerson For</td>
															<td class="FieldCell">
																<Input type="text" size="51" name="txtContactPersonFor" value="" class="FormElem">
                                                            </td>
														</tr>

														<tr>
															<td class=FieldCell width="115px"> Party Code</td>
															<td class="FieldCell" valign="bottom">
															<a href="Javascript:SelPartyPopup()"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Select Party" ></a>
															<span id="spParty" class="DataOnly"></span> &nbsp;
															<a style="width: 1em; height: 1em;" href="#" onclick="DelParty()" >
															<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Click here to delete the Party" width="12px" height="12px">
															</a>
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
                                            <td class="FieldCellSub">Address</p>
                                            </td>
                                            <td class="FieldCellSub" colspan="4"><input type="text" name="txtAddress1" size="81" class="FormElem">
                                            </td>
                                          </tr>
                                          <tr>
                                            <td class="FieldCellSub"><p align="left"></td>
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td class="ActionCell">
                                                                <input type="button" value="Save"		name="B2"			class="ActionButton"  onClick="PageSubmit()">
                                                                <input type="button" value="Close"		name="B3"			class="ActionButton"  onClick="GoToMain()">
                                                                <input type="button" value="Preview"	name="btnPreveiw"	class="ActionButton"  onClick="ViewData()">
                                                                <input type="reset"  value="Reset"		name="B1"			class="ActionButton"  >
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
