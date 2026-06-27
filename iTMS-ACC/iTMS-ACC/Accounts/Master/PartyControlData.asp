<%@ Language=VBScript %>
<%option explicit%>
<%
	'Program Name				:	PartyControlData.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Sep 12,2011
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
Dim oDOM,oDOMParty,MainNode,Root,sOrgName,objRs1,ndRoot,ndChild
Dim iAgentCount,iPerfCount,iLocationCount,iContactCount,iUnitCount
Dim sGroupParentName,sGroupPartyCode,iSNo,sChildPartyName
Dim sCreditLimit,sCreditDays

iParty = Request.QueryString("PartyCode")
Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs1 = Server.CreateObject("ADODB.RecordSet")
sCallTy = Request("hCallTy")
iUnitCount = 0
sGroupPartyCode = 0
if trim(iParty)="" then
	sAction = "CREATE"
else
	sAction = "EDIT"
end if

set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set oDOMParty = Server.CreateObject("Microsoft.XMLDOM")

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

	sQuery = "Select PartyName from APP_M_PartyMaster where ParentPartyCode = "& iParty
	objRs.Open sQuery,con
	if not objRs.EOF then
		do while not objRs.EOF
			sChildPartyName	= sChildPartyName & ","& objRs(0)
			objRs.MoveNext
		loop
	end if
	objRs.Close
	if sChildPartyName<>"" then
		sChildPartyName = mid(sChildPartyName,2)
	end if

end if
oDOM.Save server.MapPath("../Temp/Transaction/"&Session.SessionID&"-UNITDET.xml")

set ndRoot = oDOMParty.createElement("Root")
ndRoot.setAttribute "PartyCode",iParty
oDOMParty.appendChild ndRoot
if trim(iParty)<>"" then
sQuery = "Select OUDefinitionID,R.PartyType,R.PartySubType,PartyCode,CreditLimit, "&_
         " CreditDays,SubTypeName from APP_R_OrgParty R,APP_M_PartyTypes M  "&_
         " where R.PartyType=M.PartyType and R.PartySubType = M.PartySubType and PartyCode = "& iParty
'Response.write sQuery
objRs.Open sQuery,con
if not objRs.EOF then
    do while not objRs.EOF
        sCreditLimit = objRs(4)
        sCreditDays = objRs(5)
        if Trim(sCreditDays)="" or IsNull(sCreditDays) then sCreditDays = ""
        if Trim(sCreditLimit)="" or IsNull(sCreditLimit) then sCreditLimit = ""

        set ndChild = oDOMParty.createElement("Party")
            ndChild.setAttribute "Unit",objRs(0)
            ndChild.setAttribute "Type",objRs(1)
            ndChild.setAttribute "SubType",objRs(2)
            ndChild.setAttribute "Code",objRs(3)
            ndChild.setAttribute "CreditLimit",sCreditLimit
            ndChild.setAttribute "CreditDays",sCreditDays
            ndChild.setAttribute "SubTypeName",objRs(6)
            ndRoot.appendChild ndChild
        objRs.MoveNext
    loop
end if
objRs.Close
end if 'if trim(iParty)<>"" then
oDOMParty.Save server.MapPath("../Temp/Transaction/PartyDetails_"&Session.SessionID&".xml")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/trim.js"></SCRIPT>
<XML ID="UNITDET" src="<%="../Temp/Transaction/"&Session.SessionID&"-UNITDET.xml"%>"></XML>
<XML ID="OutData" ></XML>
<XML id="PartyData" src="<%="../Temp/Transaction/PartyDetails_"&Session.SessionID&".xml"%>"></XML>
<XML id="TempData"><Root/></XML>
<XML id="GroupData"><Root/></XML>
<SCRIPT LANGUAGE=vbscript>
'***************************************
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
Function PrintFun()
	Dim  objhttp

	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.open "GET","ParPrintDetailsPopulate.asp?Action=FIND",false
	objhttp.send
	if trim(objhttp.responseText)="T" then
		objhttp.open "POST","ParPrintDetailsPopulate.asp?PartyCode="&document.formname.hPartyCode.value ,false
		objhttp.send
		if trim(objhttp.responseText)<>"" then
			alert(objhttp.responseText)
		else
			'alert("../temp/master/PartyPrinting_"& document.formname.hPartyCode.value &".xml")
			window.open "../temp/master/PartyPrinting_"& document.formname.hPartyCode.value &".xml","","Status:No"
		end if
	elseif trim(objhttp.responseText)="F" then
		alert("Please Create a Print Setup File")
		PrintSetup
	else
		alert(objhttp.responseText)
	end if

End Function
'****************************************
Function PrintSetup()
	showModalDialog "ParPrintSetup.asp","","dialogWidth:600px;dialogHeight:400;Status:No"
'	document.formname.action = "ParPrintSetup.asp"
'	document.formname.submit
End Function
'*********************************************
Function DetailsData()
	document.formname.action = "ParCreate_Edit_Entry.asp?PartyCode="& document.formname.hPartyCode.value
	document.formname.submit
End Function
'**************************************************
Function GoToMain()
document.formname.action = "ParDisplayGrid.asp"
document.formname.submit
End Function
'***********************************
'*******************************************************************
FUNCTION popPartyDet(sTemp)
Dim sUnit,iCtr,sPartyGType,sType,Temparr,iCount,iUnitrow,Pararr,iCounter,sUseable
	if trim(sTemp)<>"" then
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		'Msgbox sTemp
		objhttp.Open "GET","XMLGetPartyDet.asp?PartyCode=" &sTemp , false
		objhttp.send
		'alert objhttp.responseText

		if objhttp.responseXML.xml <> "" then
			OutData.loadXML objhttp.responseXML.xml
			Set Root = OutData.documentElement
			'Msgbox Root.xml
			For Each PartyNode In Root.childNodes
			    document.formname.txtShortName.value=PartyNode.Attributes.getNamedItem("OrgnPartyCode").value
			    document.formname.txtPartyName.value = PartyNode.Attributes.getNamedItem("PartyName").value
			    sUnit = PartyNode.Attributes.getNamedItem("Units").value
			next
            document.formname.hUnits.value = sUnit

			Dim arrName,iUnitCountSel

			Temparr = Split(sUnit,":")

			For iCnt = 0 to UBound(Temparr)
                set ndRoot = UNITDET.documentElement

                if ndRoot.hasChildNodes() then
                    for each ndChild in ndRoot.childNodes
                        if ndChild.getAttribute("UnitID")=Temparr(iCnt) then
                            sUnitName = ndChild.getAttribute("Desc")
                            exit for
                        end if
                    next
                end if

			    set oRow = document.all.tblCredit.insertRow(document.all.tblCredit.rows.length)
			    set iCell = oRow.insertCell
			    iCell.colspan= "3"
			    iCell.className="ExcelDisplayCell"
			    iCell.innerHtml = "<b>"& sUnitName&"</b>"

			    set ndPartyRoot = PartyData.documentElement
			    if ndPartyRoot.hasChildNodes() then
			        For each ndChildParty in ndPartyRoot.childNodes
			            if ndChildParty.getAttribute("Unit")=Temparr(iCnt) then
			                sPartyUnit      = ndChildParty.getAttribute("Unit")
			                sPartyType      = ndChildParty.getAttribute("Type")
			                sSubType        = ndChildParty.getAttribute("SubType")
			                sCreditLimit    = ndChildParty.getAttribute("CreditLimit")
			                sCreditDays     = ndChildParty.getAttribute("CreditDays")
			                sSubTypeName    = ndChildParty.getAttribute("SubTypeName")

			                set oRow = document.all.tblCredit.insertRow(document.all.tblCredit.rows.length)
			                set iCell = oRow.insertCell
			                iCell.className="ExcelDisplayCell"
			                iCell.innerText = sSubTypeName

			                set iCell = oRow.insertCell
			                iCell.className="ExcelDisplayCell"
			                iCell.innerHtml = "Rs. "
			                iCell.innerHtml = iCell.innerHtml & "<input type=text name=txtCreditLimitZ"&trim(sPartyUnit)&"Z"&trim(sPartyType)&"Z"&trim(sSubType)&" class=FormElem style=text-align:right value="& sCreditLimit&">"
			                iCell.align = "Center"

			                set iCell = oRow.insertCell
			                iCell.className="ExcelDisplayCell"
			                iCell.innerHtml = "<input type=text name=txtCreditDaysZ"&trim(sPartyUnit)&"Z"&trim(sPartyType)&"Z"&trim(sSubType)&" class=FormElem style=text-align:right size=5 value="& sCreditDays &">"
			                iCell.align = "Center"

			            end if
			        Next
			    end if
			Next
		end if
	end if' if trim(sTemp)<>"" then
END FUNCTION
'********************************************
Function CheckSubmit()
    Dim CheckFlag
    sUnits = document.formname.hUnits.value
    Temparr = Split(sUnits,":")
    CheckFlag = false
	For iCnt = 0 to UBound(Temparr)
       set ndPartyRoot = PartyData.documentElement
	    if ndPartyRoot.hasChildNodes() then
	        For each ndChildParty in ndPartyRoot.childNodes
	            if ndChildParty.getAttribute("Unit")=Temparr(iCnt) then
	                sPartyUnit      = ndChildParty.getAttribute("Unit")
	                sPartyType      = ndChildParty.getAttribute("Type")
	                sSubType        = ndChildParty.getAttribute("SubType")
	                sSubTypeName    = ndChildParty.getAttribute("SubTypeName")

	                sCreditLimit = eval("document.formname.txtCreditLimitZ"&trim(sPartyUnit)&"Z"&trim(sPartyType)&"Z"&trim(sSubType)).value
	                sCreditDays  = eval("document.formname.txtCreditDaysZ"&trim(sPartyUnit)&"Z"&trim(sPartyType)&"Z"&trim(sSubType)).value

	                ndChildParty.setAttribute "CreditLimit",sCreditLimit
	                ndChildParty.setAttribute "CreditDays",sCreditDays
	                if (Trim(sCreditLimit)<>"") or (Trim(sCreditDays)<>"") then
	                    CheckFlag = true
	                end if
	            end if
	        Next
	    end if
	Next
	'alert(PartyData.xml)
	if CheckFlag = true then
	    set objhttp = CreateObject("Microsoft.XMLHTTP")
	    objhttp.open "POST","PartyControlDataInsert.asp",false
	    objhttp.send PartyData.XMLDocument
	    if Trim(objhttp.responseText)<>"" then
	        alert(objhttp.responseText)
	    else
	        document.formname.submit
	    end if
	else
	    exit function
	end if 'if CheckFlag = true then

End Function
</SCRIPT>

<script language="javascript">
window.__itmsPopupCompat = { type: "partyControlData" };
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>

</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="popPartyDet('<%=iParty%>')">
<form method="POST" name="formname">
<input type="Hidden" name="hUnitName" value="">
<input type="Hidden" name="hUnitCode" value="" >
<input type="Hidden" name="hPartyCode" value="<%=iParty%>">
<input type="Hidden" name="hOwnUnit" value="">
<input type="Hidden" name="hAction" value="<%=sAction%>">
<input type="hidden" name="hInActive" value="0">
<input type="hidden" name="hCreatedBy" value="<%=getUserID%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hParentPartyCode" value="<%=sGroupPartyCode%>">
<input type="hidden" name="hParUnit" value="N">
<input type="hidden" name="hUnits" value="">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">
		<%
				Response.Write "Party Details"
		%>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
			<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="60">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center"><a href="#" onClick="DetailsData()">Details</a>
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="60">
									<table border="0" cellpadding="0" cellspacing="0" width="100%"  class="TabCurrentTable" height="13">
										<tr>
											<td align="center">Control
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" width="60">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center"><a href="#" onClick="ViewData()">View</a>
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
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
							    <td align="center">
								</td>
								<td class=FieldCellSub> Party Code&nbsp;&nbsp;&nbsp;<input type="text" name="txtShortName" size="12" maxlength="10" class="FormElemRead" Readonly ></td>
							</tr>
							<tr>
							    <td align="center">
								</td>
								<td class=FieldCellSub> Name&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								    <Input type="text" size="60" name="txtPartyName" value="" class="FormElemRead" ReadOnly>&nbsp;
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
							    	<table id="tblCredit"  cellpadding="0" cellspacing="1" class="ExcelTable" width="100%">
									<tr>
									    <td class="ExcelHeaderCell" align="Center">Unit Name</td>
									    <td class="ExcelHeaderCell" align="Center">Credit Limit</td>
									    <td class="ExcelHeaderCell" align="Center">Credit Days</td>
									</tr>
									</table>
			                    </td>
							</tr>
							<tr>
							    <td height=5></td>
							</tr>
							<tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															    <input type="button" value="Save" name="btnSave" class="ActionButtonX"  onClick="CheckSubmit()">
																<input type="button" value="Close" name="btnClose" class="ActionButtonX"  onClick="GoToMain()">
                                                               <!-- <input type="button" value="Save" name="B2" class="ActionButton" onClick="PageSubmit()">

                                                                <input type="button" value="Preview" name="btnPreveiw" class="ActionButton"  onClick="ViewData()">
                                                                <input type="reset" value="Reset" name="B1" class="ActionButton"  >-->
														</td>
													</tr>
												</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
