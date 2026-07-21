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
<HTML><HEAD><TITLE>iTMS- PO </TITLE>
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

	set ResData = showModalDialog("../../Common/"&sProgramName&"?orgID="& sUnit &"&sIType=" & sIType & "&Stock=" & iStock & "&hSelectMode=S&Flag="+cstr(nFlag)&"&hDispButt=Y&PartyType=CR&CallFrom=PUR" ,Data,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
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
		document.formname.hAttribute.value = sAttributeSel
	end if
end Function
'----------------------------------------------------------------------------------------
Function SelectAddItem()
Dim	sIType, iItem,iClass,objhttp
	sUnit = document.formname.hUnit.value
	iStock = "Y"
    sTempValWindowSize = GetWindowSizeForPopup("1")
	sArrTempValWindowSize = split(sTempValWindowSize,":")
	sProgramName = sArrTempValWindowSize(0)
	sPopupHeight = sArrTempValWindowSize(1)
	sPopupWidth = sArrTempValWindowSize(2)
	
	set ResData = showModalDialog("../../Common/"&sProgramName&"?orgID="& sUnit &"&sIType=" & sIType & "&Stock=" & iStock & "&hSelectMode=M&Flag="+cstr(nFlag)&"&hDispButt=Y&PartyType=CR&CallFrom=PUR" ,ItemAddData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
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
	Set Root = ItemAddData.documentElement
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
		SpnAdditionalMaterials.innerHTML   = sItemNames
		document.formname.hAddMatAs.value = sListOfItemsSelected
		document.formname.hAddAttribute.value = sAttributeSel
	end if
end Function
'----------------------------------------------------------------------------------------
function saveXML()
Dim newElem,SCProcess,MatRecdAs
set rtData = objTemp.documentElement


set objMat = eval("document.formname.radType")

if objMat(0).checked then
    sMatType = objMat(0).value
elseif objMat(1).checked then
    sMatType = objMat(1).value
elseif objMat(2).checked then
    sMatType = objMat(2).value
end if

if trim(sMatType)="P" then
	Lvalue = document.formname.txtLabourCharge.value
	if(not checkNumbers(Lvalue)) then
		msgbox "Enter Only Numberic",0,"Numerals"
		document.formname.txtLabourCharge.select()
		exit function
	End If

	if document.formname.hOrderFor.value = "C" then
		If  trim(document.formname.cmbSCProcess.value)= "0"  then
			Msgbox "Select Subcontracting Process"
			document.formname.cmbSCProcess.focus()
			exit function
		elseIf  trim(document.formname.cmbMatRecdAs.value)= "0"  then
			Msgbox "Select Material Received As"
			'document.formname.cmbMatRecdAs.focus()
			exit function
		End if
		SCProcess = trim(document.formname.cmbSCProcess.value)
		MatRecdAs = trim(document.formname.cmbMatRecdAs.value)
		sMatRecdAsDescr = SpnMaterialToBeReceived.innerHTML
		sMatRecdAsItemType = ""'document.formname.selItmType.value
	    sAttributeList = document.formname.hAttribute.value
	    iEntryNo = document.formname.hEntryNo.value
	else
		SCProcess = ""
		MatRecdAs = ""
		sMatRecdAsDescr = ""
		sMatRecdAsItemType = ""
		sAttributeList = ""
		iEntryNo = ""
	end if

	If trim(document.formname.txtInstruct.value)= ""  then
		Msgbox "Enter Instruction"
		document.formname.txtInstruct.focus()
		exit function
	End if
	if rtData.hasChildNodes() then
	    for each ndChild in rtData.childNodes
	        if ndChild.nodeName="PRIMARYADDITIONALDET" then
	            if ndChild.getAttribute("PItemCode") = document.formname.hItemCode.value and ndChild.getAttribute("EntryNo")=iEntryNo then
	                rtData.removeChild ndChild
	                exit for
	            end if 'if ndChild.getAttribute("PItemCode") = document.formname.hItemCode.value and ndChild.getAttribute("EntryNo")=iEntryNo then
	        end if
	    next
	end if

		if trim(MatRecdAs)<>"" then
		    sArrMatRecdAs =  split(MatRecdAs,",")
		    sArrItemName = split(sMatRecdAsDescr,",")
		    sArrAttList = split(sAttributeList,",")
		    For iCnt = 0 to UBound(sArrMatRecdAs)
		        sArrItemCode =split(sArrMatRecdAs(iCnt),":")
		        sItemCode = sArrItemCode(0)
		        sClassCode = sArrItemCode(1)
		        if trim(sAttributeList)<>"" then
		            if trim(sArrAttList(iCnt))<>"" then
		                sAttList = split(sArrAttList(iCnt),":")(0)
		                sAttList = split(sAttList,"#")
		                if UBound(sAttList)=1 then
		                    sAttID = sAttList(1)
		                else
		                    sAttID = sAttList(0)
		                end if
		            end if
		        end if 'if trim(sAttributeList)<>"" then
		        
		        if trim(sAttID)="0" then sAttID=""


	            Set newElem = objTemp.createElement("PRIMARYADDITIONALDET")
	            newElem.setAttribute "MatRecdAT",""
	            newElem.setAttribute "SCProcess",  SCProcess
	            newElem.setAttribute "MatRecdAsItem",sItemCode
	            newElem.setAttribute "MatRecdAsCode",sClassCode
	            newElem.setAttribute "Instruct",document.formname.txtInstruct.value
	            newElem.setAttribute "LabourCharge",document.formname.txtLabourCharge.value
	            newElem.setAttribute "Currency",document.formname.cmbCurrency.value
	            newElem.setAttribute "MatRecdAsDescr",sArrItemName(iCnt)
	            newElem.setAttribute "MatRecdAsItemType",  sMatRecdAsItemType
	            newElem.setAttribute "AttributeList",sAttID
	            newElem.setAttribute "PItemCode",document.formname.hItemCode.value
	            newElem.setAttribute "PClassCode",document.formname.hClassCode.value
	            newElem.setAttribute "MatType",sMatType
	            newElem.setAttribute "EntryNo",iEntryNo
	            rtData.appendChild NewElem
	        Next 'For iCnt = 0 to UBound(sArrMatRecdAs)
	    end if
else

iEntryNo = document.formname.hEntryNo.value
    Set newElem = objTemp.createElement("PRIMARYADDITIONALDET")
        newElem.setAttribute "MatRecdAT",""
        newElem.setAttribute "SCProcess",""
        newElem.setAttribute "MatRecdAsItem",""
        newElem.setAttribute "MatRecdAsCode",""
        newElem.setAttribute "Instruct",""
        newElem.setAttribute "LabourCharge",""
        newElem.setAttribute "Currency",""
        newElem.setAttribute "MatRecdAsDescr",""
        newElem.setAttribute "MatRecdAsItemType",""
        newElem.setAttribute "AttributeList",""
        newElem.setAttribute "PItemCode",document.formname.hItemCode.value
        newElem.setAttribute "PClassCode",document.formname.hClassCode.value
        newElem.setAttribute "MatType",sMatType
        newElem.setAttribute "EntryNo",iEntryNo
        rtData.appendChild NewElem
end if 
	    set window.returnValue = objTemp.documentElement
	    window.close
End Function
'*****************************
Function window_onunload()
set window.returnValue = objTemp.documentElement
End Function
'************************************
Function ShowAdd()
    if document.formname.radType(0).checked then
        sType =  document.formname.radType(0).value
    elseif document.formname.radType(1).checked then
        sType =  document.formname.radType(1).value
    elseif document.formname.radType(2).checked then
        sType =  document.formname.radType(2).value
    end if
    if trim(sType)="P" then
        tblAddDet.style.display="block"
    else
        tblAddDet.style.display="none"
    end if
End Function
'******************************
Function Init()
    sPItemCode = document.formname.hItemCode.value
    sPClassCode = document.formname.hClassCode.value
    sPEntryNo = document.formname.hEntryNo.value
    set ndRoot = objTemp.documentElement
    if ndRoot.hasChildNodes() then
        for each ndPriItem in ndRoot.childNodes
            if trim(ndPriItem.nodeName)="PRIMARYADDITIONALDET" then
                sItemCode = ndPriItem.getAttribute("PItemCode")
                sClassCode = ndPriItem.getAttribute("PClassCode")
                sEntryNo = ndPriItem.getAttribute("EntryNo")
                if trim(sPItemCode)=trim(sItemCode) and trim(sPClassCode)=trim(sClassCode) and trim(sPEntryNo)=trim(sEntryNo) then
                    SpnMaterialToBeReceived.innerHTML   =ndPriItem.getAttribute("MatRecdAsDescr")
		            document.formname.cmbMatRecdAs.value = ndPriItem.getAttribute("MatRecdAsItem")&":"& ndPriItem.getAttribute("MatRecdAsCode")
		            sMatType = ndPriItem.getAttribute("MatType")
		            sProcess = ndPriItem.getAttribute("SCProcess")
		            if trim(sMatType)="P" then
		                document.formname.radType(0).checked =true
		                
		                for iCnt = 0 to document.formname.cmbSCProcess.length - 1
		                    if trim(document.formname.cmbSCProcess(iCnt).value) = trim(sProcess) then
    		                    document.formname.cmbSCProcess.selectedIndex = iCnt
    		                    exit for
		                    end if
		                next
		                document.formname.txtInstruct.value = ndPriItem.getAttribute("Instruct")
		                document.formname.txtLabourCharge.value = ndPriItem.getAttribute("LabourCharge")
		            elseif trim(sMatType)="C" then
		                document.formname.radType(1).checked =true
		            elseif trim(sMatType)="A" then
		                document.formname.radType(2).checked =true
		            end if
		                
                end if
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

sRequest = Request.QueryString("sTemp")

sArrTemp = split(sRequest,"|")
iClassCode =sArrTemp(0)
iItemCode = sArrTemp(1)
sUnit = sArrTemp(2)
sSupplier = sArrTemp(3)
sItemMode = sArrTemp(5)
sReturnable = sArrTemp(6)
sReturnItem = sArrTemp(7)
iEntryNo = sArrTemp(8)
OrderFor = "C"

bFlag = false

if sSource = "W" Then
	sDNo = indrwgStoreNo
End If

sDesc  = GetItemName(iItemCode,iClassCode)

if trim(sItemMode)="" or IsNull(sItemMode) then sItemMode = "P"

%>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init()">
<form method="POST" name="formname" action="">
<input type="hidden" name="hOrderFor" value="<%=OrderFor%>">
<input type="hidden" name="hUnit" value="<%=sUnit%>">
<input type="hidden" name="hAttribute" value="">
<input type="hidden" name="hAddMatAs" value="">
<input type="hidden" name="hAddAttribute" value="">
<input type="hidden" name="hItemCode" value="<%=iItemCode%>" />
<input type="hidden" name="hClassCode" value="<%=iClassCode%>" />
<input type="hidden" name="hRequest" value="<%=sRequest%>" />
<input type="hidden" name="hEntryNo" value="<%=iEntryNo%>" />


<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
    <tr>
		<td align="center" class="TopPack">
		</td>
    </tr>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
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
					<TD class=TabBodyWithTopLine>
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
                                            <td class="FieldCell">Materials&nbsp;sent as</td>
                                            <td class="FieldCell" colspan="4">
                                            <span class="DataOnly"><%=sDesc%>&nbsp;</span>
                                            </td>
                                        </tr>
                                         <tr>
                                            <td class="FieldCell">Type</td>
                                            <td class="FieldCellSub">
                                                <input type="radio" name="radType" value="P" onclick="ShowAdd()" <%if trim(sItemMode)="P" then response.write "Checked"%>>Primary
                                                <input type="radio" name="radType" value="C" onclick="ShowAdd()" <%if trim(sItemMode)="C" then response.write "Checked"%>>Consumable
                                                <input type="radio" name="radType" value="A" onclick="ShowAdd()" <%if trim(sItemMode)="A" then response.write "Checked"%>>Accessories
											</td>
                                        </tr>
                                        <tr>
                                            <td colspan="2">
                                                <div id="tblAddDet" style="display:block;"> 
                                                    <table width="100%">
                                                        <% if trim(OrderFor) = "C" then %>
                                                        <tr>
                                                        <td class="FieldCell">Sub Contracting Process</td>
                                                        <td class="FieldCell" colspan="4">
                                                        <% if bFlag then %>
											                <select size="1" name="cmbSCProcess" class="FormElem" disabled>
										                <% else %>
											                <select size="1" name="cmbSCProcess" class="FormElem" >
										                <% End if%>
											                <option value="0">Select</option>
											                <%
											                'function to populate SC process
											                 popSubContractProcess sItemType%>
									                    </select></td>
                                                       </tr>
                                                 
                                                        <tr>
												            <td class="FieldCell">Materials to be&nbsp;received as</td>
												            <td class="FieldCell" colspan="4">
												            <%if trim(sReturnable)="Y" and trim(sReturnItem)="S" then %>
												                <span id="SpnMaterialToBeReceived" class="DataOnly"><%=sDesc%>&nbsp;</span>
												                <input type="hidden" name="cmbMatRecdAs" value="<%=iItemCode%>:<%=iClassCode%>">
												            <%else %>
												                <input type="hidden" name="cmbMatRecdAs" value="0">
														        <a class="ExcelDisplayLink" href="#" onclick="SelectItem()">
															        <img border="0" src="../../assets/images/iTMS%20Icons/Entry.gif" alt="Select Item" width="15" height="15">
														        </a>
														        <span id="SpnMaterialToBeReceived" class="DataOnly">&nbsp;</span>
														    <%end if  %>

												            </td>
												        </tr>
                                                       <%end if%>
                                                            <tr>
                                                        <%' if trim(OrderFor) = "E" then %>
												            <td class="FieldCell" valign="top">Instruction</td>
                                                        <%' else %>
												            <!--td class="FieldCell" valign="top">Subcontracting Instruction</td-->
                                                        <%'end if%>
                                                        <td class="FieldCell" colspan="4">
											            <% if bFlag then %>
												            <textarea rows="4" name="txtInstruct" cols="60" class="FormElem" readonly><%=sInstruct%></textarea>
											            <% else %>
												            <textarea rows="4" name="txtInstruct" cols="60" class="FormElem"></textarea>
											            <% end if%>
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
                                            <p align="center">
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

