<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	IssSubConProcessDetailsPop.asp
	'Module Name				:	Inventory(Issue for Subcontract)
	'Author Name				:	Ragavendran R
	'Created On					:	Oct 10,2013
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!-- #include File="../../include/sessionVerify.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/purpopulate.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS-SubContract Process Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="ItemAddData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="Data"><Root></Root></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></script>

<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
set objTemp = window.dialogArguments
Function checkNumbers(val)
	dim valid,temp,i
	valid = "0123456789"
	for i=1 to len(val)
		temp = mid(val,i,1)
		if Instr(1,valid,temp) > 0 then
			checkNumbers = true
		else
			checkNumbers = false
			exit for
		end if
	next
end Function
'**************************************
Function SelectItem()
Dim	sIType, iItem,iClass,objhttp
	sUnit = document.formname.hUnit.value
	iStock = "Y"
    sTempValWindowSize = GetWindowSizeForPopup("1")
	sArrTempValWindowSize = split(sTempValWindowSize,":")
	sProgramName = sArrTempValWindowSize(0)
	sPopupHeight = sArrTempValWindowSize(1)
	sPopupWidth = sArrTempValWindowSize(2)

	set ResData = showModalDialog("../../Common/"&sProgramName&"?orgID="& sUnit &"&sIType=" & sIType & "&Stock=" & iStock & "&hSelectMode=M&Flag="+cstr(nFlag)&"&hDispButt=Y&PartyType=CR&CallFrom=PUR" ,Data,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	sAct = UCase(trim(ResData.getAttribute("Action")))
	sQuery = trim(ResData.getAttribute("PassQuery"))
	if ucase(trim(sAct)) <> "CLOSE" then
		do while sAct <> "DONE"
			set ResData = showModalDialog("../../Common/"&sProgramName&"?"&sQuery ,Data,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			sAct = UCase(trim(ResData.getAttribute("Action")))
			if ucase(Trim(sAct)) = "CLOSE" then exit do
			sQuery = trim(ResData.getAttribute("PassQuery"))
		loop
	end if
	Set Root = Data.documentElement
	If not Root.hasChildNodes Then 	exit function

	sItemNames = ""
	sListOfItemsSelected = ""
	if Root.hasChildNodes() then
		For Each HeaderNode In Root.childNodes
			sItemCode = Trim(HeaderNode.getAttribute("ItemCode"))
			sClassCode = Trim(HeaderNode.getAttribute("ClassCode"))
			sItemNames   = trim(sItemNames) & "," & Trim(HeaderNode.getAttribute("ItemName"))
			sListOfItemsSelected = trim(sListOfItemsSelected) & "," & trim(sItemCode)&":"& sClassCode
			sAttributeSel = trim(sAttributeSel)&","& Trim(HeaderNode.getAttribute("AttributeList"))
		next
		if trim(sItemNames) <> "" then
			sItemNames = mid(sItemNames,2)
			sListOfItemsSelected = mid(sListOfItemsSelected,2)
			sAttributeSel = mid(sAttributeSel,2)
		end if
		SpnMaterialToBeReceived.innerHTML   = sItemNames
		document.formname.cmbMatRecdAs.value = sListOfItemsSelected
	end if
end Function
'----------------------------------------------------------------------------------------
function saveXML()
Dim newElem,SCProcess,MatRecdAs
set rtData = objTemp.documentElement
MatRecdAs = trim(document.formname.cmbMatRecdAs.value)
	Lvalue = document.formname.txtLabourCharge.value
	if(not checkNumbers(Lvalue)) then
		msgbox "Enter Only Numberic",0,"Numerals"
		document.formname.txtLabourCharge.select()
		exit function
	elseIf  trim(document.formname.cmbSCProcess.value)= "0"  then
		Msgbox "Select Subcontracting Process"
		document.formname.cmbSCProcess.focus()
		exit function
'	elseIf  trim(document.formname.cmbMatRecdAs.value)= "0"  then
'		Msgbox "Select Material Received As"
'		'document.formname.cmbMatRecdAs.focus()
'		exit function
	elseIf trim(document.formname.txtInstruct.value)= ""  then
		Msgbox "Enter Instruction"
		document.formname.txtInstruct.focus()
		exit function
	elseif Trim(MatRecdAs)="0" then
	    MsgBox "Select Material Received As"
	    exit function
	end if

	set objProcess = document.formname.cmbSCProcess

	SCProcess = trim(objProcess(objProcess.selectedIndex).value)
	SCProcessName = trim(objProcess(objProcess.selectedIndex).text)
	MatRecdAs = trim(document.formname.cmbMatRecdAs.value)
	sMatRecdAsDescr = SpnMaterialToBeReceived.innerHTML
	sHardWaste = trim(document.formname.txtHardWaste.value)
	sInvWaste = trim(document.formname.txtInvWaste.value)
	if Trim(sHardWaste)="" or IsNull(sHardWaste) then sHardWaste = "0"
	if Trim(sInvWaste)="" or IsNull(sInvWaste) then sInvWaste ="0"

	if rtData.hasChildNodes() then
	    for each ndChild in rtData.childNodes
	        if ndChild.nodeName="SubContract" then
                rtData.removeChild ndChild
                exit for
	        end if
	    next
	end if
	Set newElem = objTemp.createElement("SubContract")
        newElem.setAttribute "SCProcess",  SCProcess
        newElem.setAttribute "Instruct",document.formname.txtInstruct.value
        newElem.setAttribute "LabourCharge",document.formname.txtLabourCharge.value
        newElem.setAttribute "Currency",document.formname.cmbCurrency.value
        newElem.setAttribute "HardWaste",sHardWaste
        newElem.setAttribute "InvWaste",sInvWaste
        newElem.setAttribute "ProcessName",SCProcessName
        rtData.appendChild newElem
	    if trim(MatRecdAs)<>"" and trim(MatRecdAs)<>"0" then
		    sArrMatRecdAs =  split(MatRecdAs,",")
		    sArrItemName = split(sMatRecdAsDescr,",")
		    For iCnt = 0 to UBound(sArrMatRecdAs)
		        sArrItemCode =split(sArrMatRecdAs(iCnt),":")
		        sItemCode = sArrItemCode(0)
		        sClassCode = sArrItemCode(1)

		        set newElem1 = objTemp.createElement("Details")
		            newElem1.setAttribute "MatRecdAsItem",sItemCode
	                newElem1.setAttribute "MatRecdAsCode",sClassCode
	                newElem1.setAttribute "MatRecdAsDescr",sArrItemName(iCnt)
	                newElem.appendChild newElem1
	        Next 'For iCnt = 0 to UBound(sArrMatRecdAs)
	    end if

	    set window.returnValue = objTemp.documentElement
	    window.close
End Function
'*****************************
Function window_onunload()
set window.returnValue = objTemp.documentElement
End Function
'************************************
Function Init()
    set ndRoot = objTemp.documentElement
    if ndRoot.hasChildNodes() then
        for each ndPriItem in ndRoot.childNodes
            if trim(ndPriItem.nodeName)="SubContract" then
                sProcess = ndPriItem.getAttribute("SCProcess")
                sInstruct = ndPriItem.getAttribute("Instruct")
                sLabourCharge = ndPriItem.getAttribute("LabourCharge")
                sCurr = ndPriItem.getAttribute("Currency")
                sHardWaste = ndPriItem.getAttribute("HardWaste")
                sInvWaste = ndPriItem.getAttribute("InvWaste")

                document.formname.txtInstruct.value = sInstruct
                document.formname.txtLabourCharge.value = sLabourCharge
                document.formname.txtHardWaste.value = sHardWaste
                document.formname.txtInvWaste.value = sInvWaste

                for iCnt = 0 to document.formname.cmbSCProcess.length - 1
                    if trim(document.formname.cmbSCProcess(iCnt).value) = trim(sProcess) then
	                    document.formname.cmbSCProcess.selectedIndex = iCnt
	                    exit for
                    end if
                next

                for iCnt = 0 to document.formname.cmbCurrency.length - 1
                    if trim(document.formname.cmbCurrency(iCnt).value) = trim(sCurr) then
                        document.formname.cmbCurrency.selectedIndex = iCnt
                        exit for
                    end if
                next

                for each ndSubContDet in ndPriItem.childNodes
                    if ndSubContDet.nodeName="Details" then
                        sItemCode = ndSubContDet.getAttribute("MatRecdAsItem")
                        sClassCode = ndSubContDet.getAttribute("MatRecdAsCode")
                        sItemName = ndSubContDet.getAttribute("MatRecdAsDescr")
                        sMatRecdAs = sMatRecdAs &","& sItemCode&":"&sClassCode
                        sMatDesc = sMatDesc &","& sItemName
                    end if
                next

	            SpnMaterialToBeReceived.innerHTML   =mid(sMatDesc,2)
	            document.formname.cmbMatRecdAs.value = mid(sMatRecdAs,2)
            end if
        next
    end if
End Function
</Script>

<%
Dim objRs,sDrgNo,sItemType,sDesc,OrderFor,sUnit,sSupp,sSql,iClassCode,iItemCode,sSupplier,bFlag
Dim sClassDesc,sItemDesc,saTemp,sDrawVerNo,sSource,sPRNo,sPRNoStr,iItemRecdAs,iSCProcess
Dim indrwgStoreNo,iItemRecdAt,sInstruct,iTempItemCode,sDNo,sCallFrom,iForPurNo,sItemMode
Dim sReturnable,sReturnItem,iEntryNo

Dim sRequest,sArrTemp
Set objRs = server.CreateObject("Adodb.recordset")
sUnit = Request.QueryString("Unit")


%>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init()">
<form method="POST" name="formname" action="">
<input type="hidden" name="hUnit" value="<%=sUnit%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
    <tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">
		Additional Details
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
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                   <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td colspan="2">
                                                <div id="tblAddDet" style="display:block;">
                                                    <table width="100%">
                                                        <tr>
                                                        <td class="FieldCell">Sub Contracting Process</td>
                                                        <td class="FieldCell" colspan="4">
                                                            <select size="1" name="cmbSCProcess" class="FormElem" >
										                    <option value="0">Select</option>
											                <%
											                'function to populate SC process
											                 popSubContractProcess sItemType%>
									                    </select></td>
                                                       </tr>

                                                        <tr>
												            <td class="FieldCell">Materials to be&nbsp;received as</td>
												            <td class="FieldCell" colspan="4">
												                <input type="hidden" name="cmbMatRecdAs" value="0">
														        <a class="ExcelDisplayLink" href="#" onclick="SelectItem()">
															        <img border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" alt="Select Item" width="15" height="15">
														        </a>
														        <span id="SpnMaterialToBeReceived" class="DataOnly">&nbsp;</span>
												            </td>
												        </tr>
                                                            <tr>
												            <td class="FieldCell" valign="top">Instruction</td>
                                                        <td class="FieldCell" colspan="4">
											            <textarea rows="4" name="txtInstruct" cols="60" class="FormElem"></textarea>
                                                        </td>
                                                            </tr>
                                                            <tr>
											            <td class="FieldCellsub" valign="top">Labour Charges </td>
											            <td class="FieldCell" >
												            <Input type="text" name="txtLabourCharge" value="" class=formelem size="10" style="text-align: right">
												            <SELECT size="1" class="Formelem" NAME = "cmbCurrency">
												            <%
													               'To populate currency
														            popSelCurrency(0)
												            %>
												            </select>
											            </td>
                                                    </tr>
                                                       <tr>
											            <td class="FieldCellsub" valign="top">Hard Waste</td>
											            <td class="FieldCell" >
												            <Input type="text" name="txtHardWaste" value="" class=formelem size="10" style="text-align: right">%
											            </td>
                                                    </tr>
                                                       <tr>
											            <td class="FieldCellsub" valign="top">Invisible Waste</td>
											            <td class="FieldCell" >
												            <Input type="text" name="txtInvWaste" value="" class=formelem size="10" style="text-align: right">%
											            </td>
                                                    </tr>
                                                    </table>
                                                </div>
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
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
								<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td valign="middle" class="ActionCell">
                                            <input type="button" value="Done" name="B4" onClick="saveXML()" class="ActionButton" tabindex="3">
                                            <input type="button" value="Close" name="B6" onClick="javascript:window.close();" class="ActionButton" tabindex="3">
                                            <input type="reset" value="Reset" name="B5" class="ActionButton" tabindex="3">
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
</Html>

<%
function popSubContractProcess(ItemType)
Dim PrcID,PrcName

sSql ="Select SubConProcessID,SubConProcessName from App_M_SubContractProcess"
  '" Where ItemTypeId='" & ItemType & "'"
 'Response.Write sSql
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSql
	.ActiveConnection = con
	.Open
End With
Set objRs.ActiveConnection = nothing

Set PrcID = objRs(0)
Set PrcName = objRs(1)

If not objRs.EOF then
	Do While Not objRs.EOF
			Response.Write("<OPTION VALUE="""&trim(PrcID)&""">"&trim(PrcName)&"</OPTION>" &vbcrlf)
		objRs.MoveNext
	Loop
end if
objRs.Close

End function
%>

<%
'' to fetch the Unit No ,Unit Name
Function popUnitOrgRel(selUnit)
Dim sUnitNo,sUnitName
	sSql  = "Select OUDefinitionID,OrgUnitShortDescription from DCS_OrganizationUnitDefinitions where Len(OUDefinitionID)> 4"
	With objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	End With
	Set objRs.ActiveConnection = nothing

	Set sUnitNo = objRs(0)
	Set sUnitName = objRs(1)

	If not objRs.EOF then
		Do While Not objRs.EOF
			if trim(sUnitNo) = trim(selUnit) then
				Response.Write("<OPTION VALUE="""&trim(sUnitNo)&""" selected>"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			Else
				Response.Write("<OPTION VALUE="""&trim(sUnitNo)&""">"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			end if
			objRs.MoveNext
		Loop
	end if
	objRs.Close
End function
%>

