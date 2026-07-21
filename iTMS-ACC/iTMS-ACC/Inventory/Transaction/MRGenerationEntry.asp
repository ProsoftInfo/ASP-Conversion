<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	MRGenerationEntry.asp
	'Module Name				:	Inventory (Transaction)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	June 27, 2005
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>Material Requisition Creation</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<script type="application/xml" data-itms-xml-island="1" id="PurTypeData"></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemData"></script>
<script type="application/xml" data-itms-xml-island="1" id="UoMData" data-src="../../inventory/xmldata/Uom.xml"></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><ROOT></ROOT></script>
<script type="application/xml" data-itms-xml-island="1" id="OutSelectData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="OutCost"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemTypeData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root/></script>
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/RoundOff.js"></script>
<script LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/TempItem.js"></script>
<script Language="javascript" Src="../../scripts/RefTypePop.js"></script>
<script language="javascript" src="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
dim j,iCounter1,ssupcode,iser,iTempCode,Partyname,iRcptNo
j = 0
iser = 0
iCounter1 = 0
iTempCode = 0
'*********************************************
Function GetDetails()
    Dim sRefType,sOrgID,sItemType,sPartyCode,iStock,nFlag,bAddButton
    sRefType = document.formname.selRefName(document.formname.selRefName.selectedIndex).value
    sorgID = document.formname.hUnit.value
    sItemType = ""
    sPartyCode = ""
    iStock = "Y"
    nFlag=1
    bAddButton = "N"


    if sUsage = "PRD" and sIssFor="M" and sRefType<>"14" then
        alert("Select Mix Code Reference")
        document.formname.selRefName.focus
        exit function
    end if
sDispItem = 0

sCallFrom = "MR"
    'RefTypeSelection sRefType,sOrgID,sItemType,sPartyCode,iStock,nFlag,bAddButton,sDispItem
    RefTypeSelection sRefType,sOrgID,sPartyCode,iStock,nFlag,bAddButton,sDispItem,sCallFrom

        Set ndRoot = OutData.documentElement
        Set Root = OutSelectData.documentElement

    if trim(sRefType)<>"N" then
'        alert(ndRoot.xml)
        if ndRoot.hasChildNodes() then
            for each ndChild in ndRoot.childNodes
                if ndChild.nodeName="Reference" then
                    sRefCode = trim(ndChild.getAttribute("ReferenceCode"))
                    sRefDate = sRefDate &","& ndChild.getAttribute("ReferenceDate")
                    sRefNo = sRefNo &","& ndChild.getAttribute("ReferenceNo")
                    sRefDt = ndChild.getAttribute("ReferenceDate")
                    sRefCodeDate =  sRefCodeDate &","& sRefCode &" - "& sRefDt
                end if
            next

            sRefCodeDate = mid(sRefCodeDate,2)
            sRefNo = mid(sRefNo,2)
            sRefDate = mid(sRefDate,2)

            RefNoDate.innerHTML = sRefCodeDate
            document.formname.hRefNo.value = sRefNo
            document.formname.hRefDate.value = sRefDate

            if sRefNo <>"" then
                set objhttp = CreateObject("Microsoft.XMLHTTP")
                    objhttp.open "GET","InvGetItemDetForRefType.asp?RefType="&sRefType &"&RefCodes="&sRefNo&"&orgID="& document.formname.hUnit.value,false
                    objhttp.send
'                    alert(Objhttp.responseText)
                    if trim(objhttp.responseXML.xml)<>"" then
                        ItemData.loadXML(objhttp.responseXML.xml)
                    else
                        alert(Objhttp.responseText)
                    end if
                    set ndIRoot = ItemData.documentElement
                    if ndIRoot.hasChildNodes() then
                        for each ndIChild in ndIRoot.childNodes
                            if ndIChild.nodeName="Item" then

                                if trim(document.formname.hItemType.value)="" then
                                    sItemCode = ndIChild.getAttribute("ItemCode")
                                    set objhttp=CreateObject("Microsoft.XMLHTTP")
                                    objhttp.open "GET","../../Include/GetItemTypeForItem.asp?ItemCode="&sItemCode ,false
                                    objhttp.send
                                    if trim(objhttp.responseXML.xml)<>"" then
                                        ItemTypeData.loadXML(objhttp.responseXML.xml)
                                        set iRoot = ItemTypeData.documentElement
                                        document.formname.hItemType.value = iRoot.getAttribute("ItemType")
                                    else
                                        alert(objhttp.responseText)
                                    end if
                                end if

                                Set newElem = OutSelectData.createElement("ITEMDETAILS")
			                        newElem.setAttribute "ENTRYNO",ndIChild.getAttribute("EntryNo")
			                        newElem.setAttribute "ITEMCODE", ndIChild.getAttribute("ItemCode")
			                        newElem.setAttribute "CLASSCODE", ndIChild.getAttribute("ClassCode")
			                        newElem.setAttribute "UNIT", document.formname.hUnit.value
			                        newElem.setAttribute "ITEMNAME", ndIChild.getAttribute("ItemName")
			                        newElem.setAttribute "UOM", ndIChild.getAttribute("StoresUoM")
			                        newElem.setAttribute "DECIMAL", ndIChild.getAttribute("Decimal")
			                        newElem.setAttribute "DISPALYED", "N"
			                        newElem.setAttribute "QTY", "0"
			                        newElem.setAttribute "REQUIREDBY", ""
			                        newElem.setAttribute "REQUIREDVALUE", ""
			                        newElem.setAttribute "ATTRIBUTELIST",ndIChild.getAttribute("AttributeList")
			                        newElem.setAttribute "REMARKS",""
			                        newElem.setAttribute "RefNo",ndIChild.getAttribute("No")
			                        Root.appendChild newElem
			                end if  'if ndChild.nodeName="Item" then
                        next
                    end if
            end if 'if sRefNo <>"" then
        end if
    else
        if ndRoot.hasChildNodes() then
            for each ndChild in ndRoot.childNodes
                if ndChild.nodeName="Item" then
                    Set newElem = OutSelectData.createElement("ITEMDETAILS")
			            newElem.setAttribute "ENTRYNO",ndChild.getAttribute("EntryNo")
			            newElem.setAttribute "ITEMCODE", ndChild.getAttribute("ItemCode")
			            newElem.setAttribute "CLASSCODE", ndChild.getAttribute("ClassCode")
			            newElem.setAttribute "UNIT", document.formname.hUnit.value
			            newElem.setAttribute "ITEMNAME", ndChild.getAttribute("ItemName")
			            newElem.setAttribute "UOM", ndChild.getAttribute("StoresUoM")
			            newElem.setAttribute "DECIMAL", ndChild.getAttribute("Decimal")
			            newElem.setAttribute "DISPALYED", "N"
			            newElem.setAttribute "QTY", "0"
			            newElem.setAttribute "REQUIREDBY", ""
			            newElem.setAttribute "REQUIREDVALUE", ""
			            newElem.setAttribute "ATTRIBUTELIST",ndChild.getAttribute("AttributeList")
			            newElem.setAttribute "REMARKS",""
			            newElem.setAttribute "RefNo",""
			            Root.appendChild newElem
			    end if  'if ndChild.nodeName="Item" then
			next 'for each ndChild in ndRoot.childNodes
        end if 'if ndRoot.hasChildNodes() then
    end if
    'alert(Root.xml)
    DisplayTable date
End Function
'********************************************************************************************
Function init()
Dim Root,PartyNode,RefNode,subNode,sRefName,RefNo,IssName,sIssCode,sUsage

	IF DateValue(Date()) <  DAteValue(document.formname.hToDate.value) and   DateValue(Date()) > dateValue(document.formname.hFrmDate.value) then
		document.formname.ctlCDDate.SetDate = date()
	Else
	  document.formname.ctlCDDate.SetDate = document.formname.hToDate.value
	End IF

End Function

'********************************************************************************************
Function MinDate()

  	Dim sMinDate,sFinPeriod,sSelDate,sMaxDate
  	sMinDate = document.formname.hFrmDate.value
  	sMaxDate = document.formname.hToDate.value
  	dDate = document.formname.ctlCDDate.getdate

  	'alert(RngFrom &"="& sMinDate)
  	If DateValue(dDate) < DateValue(sMinDate) or  DateValue(dDate) > DateValue(sMaxDate) then
  		Alert("Date Should be With in the Range "& sMinDate & " to " & sMaxDate)
  		document.formname.ctlCDDate.Setdate = sMaxDate
  		Exit function
  	End If

End Function
'********************************************************************************************
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

Function ClearTable()
	dim i
	for i=1 to document.all.tblLot.rows.length - 1
		document.all.tblLot.deleteRow(1)
	next
	j = 0
end Function

Function DeleteItems()
	set root = OutSelectData.DocumentElement

	sExp ="//ITEMDETAILS"
	Set ItemNode = Root.Selectnodes(sExp)
	if ItemNode.Length > 0 then
		for itr = 0 to ItemNode.Length - 1
			iItemDel = ItemNode.Item(itr).Attributes.getNamedItem("ITEMCODE").value
			iClassDel = ItemNode.Item(itr).Attributes.getNamedItem("CLASSCODE").value
			iEntDel   = ItemNode.Item(itr).Attributes.getNamedItem("ENTRYNO").value

			set objSel = eval("document.formname.chkDeleteA"&CStr(iItemDel)&"A"&CStr(iClassDel)&"A"&CStr(iEntDel))
			set objSer = eval("document.formname.txtSerA"&CStr(iItemDel)&"A"&CStr(iClassDel)&"A"&CStr(iEntDel))
			if objSel.checked then DeleteItem objSel,objSer.value + 1
		next
	end if
End Function

Function DeleteItem(sobj,iRow)
	dim itr
	itr = 0

	'alert(window.event.srcElement.sourceIndex)
	arrTemp = split(sobj.name,"A")
	iItem = arrTemp(1)
	iClass = arrTemp(2)
	iEntNo = arrTemp(3)
	set root = OutSelectData.DocumentElement
	if root.hasChildNodes then
		sExp ="//ITEMDETAILS [@ENTRYNO = "&iEntNo&" and  @ITEMCODE = "&iItem&" and @CLASSCODE = "&iClass&"]"
		Set DeleteNode = Root.Selectnodes(sExp)
		if DeleteNode.Length > 0 then
			Set oNode = root.RemoveChild(DeleteNode.Item(0))
			document.all.tblLot.deleteRow(iRow - 1)
		end if

		sExp ="//ITEMDETAILS"
		Set ItemNode = Root.Selectnodes(sExp)

		if ItemNode.Length > 0 then
			iser = 0
			for itr = 0 to ItemNode.Length - 1
				iser = iser + 1
				iItemDel = ItemNode.Item(itr).Attributes.getNamedItem("ITEMCODE").value
				iClassDel = ItemNode.Item(itr).Attributes.getNamedItem("CLASSCODE").value
				iEntDel   = ItemNode.Item(itr).Attributes.getNamedItem("ENTRYNO").value

				set objSer = eval("document.formname.txtSerA"&CStr(iItemDel)&"A"&CStr(iClassDel)&"A"&CStr(iEntDel))
				objSer.value = iser
			next
		else
			for i=2 to document.all.tblLot.rows.length - 1
				document.all.tblLot.deleteRow(2)
			next
			iCounter1 = 0
			j = 0
			iser = 0
		end if
	end if
	'alert(OutData.xml)
End Function

Function DisplayTable(todaysDate)
	iCounter1 = 0
	set rootData = ItemData.DocumentElement
	set root = OutSelectData.DocumentElement
	set oNodTaxRoot = PurTypeData.DocumentElement
	set rootUoM = UoMData.DocumentElement

	sUnit = document.formname.hUnit.value
	'alert(rootData.xml)
	'alert(Root.xml)
	if root.hasChildNodes then
		sExp ="//ITEMDETAILS [ @DISPALYED = 'N']"
		Set CheckNode = Root.Selectnodes(sExp)
		if CheckNode.Length > 0 then
			iCounter = 0
			'iser = 0
			For iCounter = 0 to CheckNode.Length - 1
				iEntNo = CheckNode.Item(iCounter).Attributes.getNamedItem("ENTRYNO").value
				iItem = CheckNode.Item(iCounter).Attributes.getNamedItem("ITEMCODE").value
				iClass = CheckNode.Item(iCounter).Attributes.getNamedItem("CLASSCODE").value
				sItemDesc = replace(replace(CheckNode.Item(iCounter).Attributes.getNamedItem("ITEMNAME").value,"'","~"),Chr(34),"~~")
				sCheck = CheckNode.Item(iCounter).Attributes.getNamedItem("DECIMAL").value
				sItemName = replace(replace(CheckNode.Item(iCounter).Attributes.getNamedItem("ITEMNAME").value,"'","~"),Chr(34),"~~")
				sAttributeList = CheckNode.Item(iCounter).Attributes.getNamedItem("ATTRIBUTELIST").value
				j = j + 1
				iser = iser + 1
				set oRow = document.all.tblLot.insertRow(document.all.tblLot.rows.length - cint(iCtr))

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtSerA"&CStr(iItem)&"A"&CStr(iClass)&"A"&CStr(iEntNo)&""" value="""&iser&""" size=""1"" class=""FormElemRead"" style=""text-align=center"" READONLY>" )
				headerCell.appendChild(oText)
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""checkbox"" name=""chkDeleteA"&CStr(iItem)&"A"&CStr(iClass)&"A"&CStr(iEntNo)&""" value="""&iser&""" class=""FormElem"" style=""text-align=right"">" )
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelInputCell"

				set headerCell=oRow.insertCell()
				'if len(sItemDesc) > 20 then
				'	sItemDesc=mid(sItemDesc,1,20)&"..."
				'else
					sItemDesc=sItemDesc
				'end if
				headerCell.innerHTML=sItemDesc
				headerCell.className="ExcelDisplayCell"
				headerCell.align="left"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtQtyZ"&CStr(iItem)&"Z"&CStr(iClass)&"Z"&CStr(iEntNo)&""" size=""12"" class=""FormElem"" style=""text-align=right"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelInputCell"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<SELECT name=""selUoMZ"&CStr(iItem)&"Z"&CStr(iClass)&"Z"&CStr(iEntNo)&""" class=""FormElem"">" )

				set oText1 = document.createElement("<Option>" )
				oText1.Text = CheckNode.Item(iCounter).Attributes.getNamedItem("UOM").value
				oText1.Value = CheckNode.Item(iCounter).Attributes.getNamedItem("UOM").value
				oText.Options.Add(oText1)
				headerCell.className="ExcelFieldCell"
				headerCell.align="center"
				headerCell.appendChild(oText)

				set headerCell=oRow.insertCell()
				'set oText = document.createElement("<SELECT name=""selSchZ"&CStr(iItem)&"Z"&CStr(iClass)&"Z"&CStr(iEntNo)&"Z"&CStr(sAttributeList)&""" class=""FormElem"" onChange=""CheckSch(this,'"&todaysDate&"','"&CheckNode.Item(iCounter).Attributes.getNamedItem("UOM").value&"')"">" )
				set oText = document.createElement("<SELECT name=""selSchZ"&CStr(iItem)&"Z"&CStr(iClass)&"Z"&CStr(iEntNo)&""" class=""FormElem"" onChange=""CheckSch(this,'"&todaysDate&"','"&CheckNode.Item(iCounter).Attributes.getNamedItem("UOM").value&"','"&sAttributeList&"')"">" )

				set oText1 = document.createElement("<Option>" )
				oText1.Text = "Select"
				oText1.Value = "select"
				oText.Options.Add(oText1)

				set oText1 = document.createElement("<Option>" )
				oText1.Text = "Immediate"
				oText1.Value = "I"
				oText.Options.Add(oText1)

				set oText1 = document.createElement("<Option>" )
				oText1.Text = "Within x Days"
				oText1.Value = "W"
				oText.Options.Add(oText1)

				set oText1 = document.createElement("<Option>" )
				oText1.Text = "Specific Date"
				oText1.Value = "D"
				oText.Options.Add(oText1)

				set oText1 = document.createElement("<Option>" )
				oText1.Text = "Scheduled"
				oText1.Value = "S"
				oText.Options.Add(oText1)

				headerCell.className="ExcelFieldCell"
				headerCell.align="center"
				headerCell.appendChild(oText)

			'	set headerCell=oRow.insertCell()
			'	set oText = document.createElement("<img border=""0"" src=""../../assets/images/iTMS%20Icons/EntryIcon.gif"" style=""cursor:hand"" alt=""Additional Specs"" width=""11"" height=""11"" onClick=""GetAddDetails('"&iItem&"','"&iClass&"','"&sUnit&"','"&iEntNo&"','"&sAttributeList&"')"">" )
			'	headerCell.appendChild(oText)
			'	headerCell.align="center"
			'	headerCell.className="ExcelFieldCell"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<img border=""0"" src=""../../assets/images/iTMS%20Icons/EntryIcon.gif"" style=""cursor:hand"" alt=""Stock Details"" width=""11"" height=""11"" onClick=""DisplayStock('"&iItem&"','"&iClass&"','"&sUnit&"','"&iEntNo&"','"&sItemName&"')"">" )
				headerCell.appendChild(oText)
				headerCell.align="center"
				headerCell.className="ExcelFieldCell"

				set headerCell=oRow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""txtRemZ"&CStr(iItem)&"Z"&CStr(iClass)&"Z"&CStr(iEntNo)&""" size=""12"" class=""FormElem"" maxlength=""100"">")
				headerCell.appendChild(oText)
				headerCell.width = "10"
				headerCell.className="ExcelInputCell"

				CheckNode.Item(iCounter).Attributes.getNamedItem("DISPALYED").value = "Y"
			Next
		end if

	end if

End Function

Function GetAddDetails(sItem,sClass,sOrg,iEntNo,iAttribList)

	set Q = eval("document.formname.txtQtyZ"&sItem&"Z"&sClass&"Z"&iEntNo)
	if trim(Q.value) = "" then
		alert("Enter Quantity")
		exit function
	end if

	if cdbl(Q.value) = 0 then exit function

	sTempValues = sClass&"|"&sItem&"|"&sOrg&"|"&sUsage&"|"&trim(Q.value)&"|"&iEntNo&"|"&iAttribList

	if sUsage = "PRD" then
		set OutValue = showModalDialog("DirectIssueMixEntry.asp?sTemp="&sTempValues,OutSelectData,"dialogHeight:300px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No")
	elseif sUsage = "PAC" then
		set OutValue = showModalDialog("DirectIssuePackingEntry.asp?sTemp="&sTempValues,OutSelectData,"dialogHeight:400px;dialogWidth:325px;center:Yes;help:No;resizable:No;status:No")
	elseif sUsage = "WIP" or sUsage = "MAT" then
		set OutValue = showModalDialog("AddEntryDetails.asp?sTemp="&sTempValues,OutSelectData,"dialogHeight:370px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
		'window.open "AddEntryDetails.asp?sTemp="&sTempValues,"OutData","",""
	'	set OutValue = showModalDialog("DirectIssueAddEntry.asp?sTemp="&sTempValues,OutData,"dialogHeight:370px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
	end if
end Function

Function DisplayStock(sItem,sClass,sOrg,iEntNo,sItemName)
	showModalDialog "../master/itmStockDetailsPop.asp?EntNo="&iEntNo&"&sItem="&sItem&"&sClass="&sClass&"&ItemName="&sItemName&"&sOrg="&sOrg,"Stock","dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No"
end Function

Function CheckSch(obj,todaysdate,sUoM,sAttribList)
	Set Root = OutSelectData.documentElement
'alert(OutData.xml)
	dim sItem,sClass,a
	arrTemp = split(obj.name,"Z")

	sItem = arrTemp(1)
	sClass = arrTemp(2)
	iEntNo =  arrTemp(3)
	'iAttribList = arrTemp(4)
	iAttribList = sAttribList

	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	if trim(iAttribList) <> "" then
	 	'sOptName = FunAttribName(sAttributeList)
	 	objhttp.Open "GET","XMLGetAttributeName.asp?Para="&iAttribList,false
		objhttp.send
		sOptName = objhttp.responsetext
	else
		sOptName = ""
	end if

	For Each HeaderNode In Root.childNodes
		if strcomp(HeaderNode.nodeName,"ITEMDETAILS")=0 then
			if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem  and HeaderNode.Attributes.Item(2).nodeValue = sClass then
				if HeaderNode.HaschildNodes() then
					For Each HNode In HeaderNode.childNodes
						if StrComp(Trim(HNode.NodeName),"Schedule") = 0 or StrComp(Trim(HNode.NodeName),"ScheduleDetails") = 0 then
							set a = HeaderNode.removeChild(HNode)
						end if
					next
				end if
			end if
		end if ' if strcomp(HeaderNode.nodeName,"ITEMDETAILS")=0 then
	Next

	if (obj.selectedIndex = "1") then
		iSchEntNo =  1
		Set Root = OutSelectData.documentElement
		For Each HeaderNode In Root.childNodes
			if strcomp(HeaderNode.nodeName,"ITEMDETAILS")=0 then
				if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem  and HeaderNode.Attributes.Item(2).nodeValue = sClass then
					Set newElem = OutSelectData.createElement("Schedule")
					newElem.setAttribute "STYPE", trim(obj(obj.selectedIndex).value)
					newElem.setAttribute "SVALUE", document.formname.ctlCDDate.GetDate
					newElem.setAttribute "ITEMCODE", sItem
					newElem.setAttribute "CLASSCODE", sClass
					newElem.setAttribute "SCHENTRYNO",	iSchEntNo
					HeaderNode.appendChild newElem
				end if
			end if 'if strcomp(HeaderNode.nodeName,"ITEMDETAILS")=0 then
				iSchEntNo = iSchEntNo + 1
		Next
	end if
	if (obj.selectedIndex = "2") then
		value = prompt("Enter No of Days","0")
		if (isNull(value)) then
			obj.selectedIndex=0
			exit function
		elseif (trim(value)="") then
			obj.selectedIndex=0
			exit function
		else
			if(trim(value)="") then
				msgbox "Enter Number of Days",0,"Number of Days"
				obj.selectedIndex=0
				exit function
			else
				if(not checkNumbers(value)) then
					msgbox "Enter Numerals Only",0,"Numerals"
					obj.selectedIndex=0
					exit function
				else
					iSchEntNo = 1
					Set Root = OutSelectData.documentElement
					For Each HeaderNode In Root.childNodes
						if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem  and HeaderNode.Attributes.Item(2).nodeValue = sClass then
							Set newElem = OutSelectData.createElement("Schedule")
							newElem.setAttribute "STYPE", trim(obj(obj.selectedIndex).value)
							newElem.setAttribute "SVALUE", trim(value)
							newElem.setAttribute "ITEMCODE", sItem
							newElem.setAttribute "CLASSCODE", sClass
							newElem.setAttribute "SCHENTRYNO",iSchEntNo
							HeaderNode.appendChild newElem
						end if
						iSchEntNo = iSchEntNo + 1
					Next
				end if
			end if
		end if
	end if
	if (obj.selectedIndex = "3") then
		value=prompt("Enter the Date","")
		if (isNull(value)) then
			obj.selectedIndex=0
			exit function
		elseif (trim(value)="") then
			objType.selectedIndex=0
			objValue.value=	""
			exit function
		else
			if (not vd(value,todaysdate)) then
				MsgBox "Invalid Date",0,"Invalid Date"
				obj.selectedIndex=0
				Exit Function
			end if
			if (DateDiff("d",todaysdate,value) < 0) then
				MsgBox "Date should be greater or equal to Today's Date",0,"Invalid Date"
				obj.selectedIndex=0
				Exit Function
			else
				iSchEntNo = 1
				Set Root = OutSelectData.documentElement
				For Each HeaderNode In Root.childNodes
					if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem  and HeaderNode.Attributes.Item(2).nodeValue = sClass then
						Set newElem = OutSelectData.createElement("Schedule")
						newElem.setAttribute "STYPE", trim(obj(obj.selectedIndex).value)
						newElem.setAttribute "SVALUE", trim(value)
						newElem.setAttribute "ITEMCODE", sItem
						newElem.setAttribute "CLASSCODE", sClass
						newElem.setAttribute "SCHENTRYNO",	iSchEntNo
						HeaderNode.appendChild newElem
					end if
						iSchEntNo = iSchEntNo + 1
				Next
			end if
		end if
	end if
	if (obj.selectedIndex = "4") then
		dim qty
		set qty = eval("document.formname.txtQtyZ"+cstr(sItem)+"Z"+cstr(sClass)+"Z"+cstr(iEntNo))
		if (trim(qty.value)="") then
			MsgBox "Enter Quantity",0,"Quantity"
			qty.focus()
			obj.selectedIndex=0
			exit function
		elseif(not checkNumbers(qty.value)) then
			msgbox "Enter Numerals Only",0,"Numerals"
			qty.focus()
			obj.selectedIndex=0
			exit function
		else
			iSchEntNo = 1
			Set Root = OutSelectData.documentElement
			For Each HeaderNode In Root.childNodes
				if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem  and HeaderNode.Attributes.Item(2).nodeValue = sClass then
					Set newElem = OutSelectData.createElement("Schedule")
					newElem.setAttribute "STYPE", trim(obj(obj.selectedIndex).value)
					newElem.setAttribute "SVALUE", ""
					newElem.setAttribute "ITEMCODE", sItem
					newElem.setAttribute "CLASSCODE", sClass
					newElem.setAttribute "SCHENTRYNO",	iSchEntNo
					HeaderNode.appendChild newElem
				end if
					iSchEntNo = iSchEntNo + 1
			Next

			sTempValues = qty.value&":"&sItem&":"&sClass&":"&iEntNo&":"&document.formname.hUnit.value&":"&sUoM&":"&sOptName

			Set OutDataValue = showModalDialog("MRGenSchedulePoP.asp?sTemp="&sTempValues,OutSelectData,"dialogHeight:510px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No")

			'window.open "MRGenSchedulePoP.asp?sTemp="&sTempValues,"OutData",""
		end if
	end if
end Function

Function popCC()
	document.formname.selCC.options.length = 1

	dim Root,HeaderNode
	sOrgID = document.formname.hUnit.value
	set objhttp = CreateObject("MSXML2.XMLHTTP")

	sIType=trim(sIType)
	objhttp.Open "GET","XMLSelectCostCenter.asp?sOrgID="& sOrgID, false

	objhttp.send
	if objhttp.responseXML.xml <> "" then
		OutCost.loadXML objhttp.responseXML.xml
		Set Root = OutCost.documentElement
		if Root.HaschildNodes() then
			For Each HeaderNode In Root.childNodes
				document.formname.selCC.length = document.formname.selCC.length+1
				document.formname.selCC.options(document.formname.selCC.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
				document.formname.selCC.options(document.formname.selCC.length-1).Value = cstr(HeaderNode.Attributes.Item(0).nodeValue)
			next
		end if
	else
		'alert("No Cost Center for the Unit Selected")
		'Exit Function
	end if
End Function

Function CreateNew(obj)
	if obj.selectedIndex = "1" then
		idConsumption.style.display = "block"
	end if
End Function

Function hideDiv()
	idConsumption.style.display = "none"
	document.formname.selAccHead.selectedIndex = 0
End Function

Function CheckSubmit(todaysdate)
Dim sAddSpcsCount
sRequstTo = document.formname.selIssueTo
sRequstFor = document.formname.cmbIssType(document.formname.cmbIssType.selectedIndex).value

set root = OutSelectData.DocumentElement

	if(datediff("d",todaysdate,document.formname.ctlCDDate.GetDate)) > 0 then
		alert("Created On should be less than or equal to Today's Date")
		exit function
	elseif len(trim(document.formname.txtRemarks.value)) > 200 then
		alert("Remarks should be less than 200 characters")
		document.formname.txtRemarks.select
		exit function
	elseif document.formname.selIssueTo.selectedIndex<=0 then
	    alert("Select Requested By")
	    document.formname.selIssueTo.focus
	    exit function
	elseif trim(sRequstFor)="SEL" then
	    alert("Select Requested For")
	    document.formname.cmbIssType.focus
	    exit function
	else
		itr = 0
		sAddSpcsCount = 0

		if not root.hasChildNodes then
		    alert("Select the Item(s)")
		    exit function
		end if

		if root.hasChildNodes then
			sExp ="//ITEMDETAILS"

			Set ItemNode = Root.Selectnodes(sExp)
			for itr = 0 to ItemNode.Length - 1
				iEntNo = ItemNode.Item(itr).Attributes.getNamedItem("ENTRYNO").value
				iItem = ItemNode.Item(itr).Attributes.getNamedItem("ITEMCODE").value
				iClass = ItemNode.Item(itr).Attributes.getNamedItem("CLASSCODE").value
				iAtributeList = ItemNode.Item(itr).Attributes.getNamedItem("ATTRIBUTELIST").value
				Set objQty = eval("document.formname.txtQtyZ"&iItem&"Z"&iClass&"Z"&iEntNo)
				'Set objReq = eval("document.formname.selSchZ"&iItem&"Z"&iClass&"Z"&iEntNo&"Z"&iAtributeList)
				Set objReq = eval("document.formname.selSchZ"&iItem&"Z"&iClass&"Z"&iEntNo)
				Set objUoM = eval("document.formname.selUoMZ"&iItem&"Z"&iClass&"Z"&iEntNo)
				Set objRem = eval("document.formname.txtRemZ"&iItem&"Z"&iClass&"Z"&iEntNo)

				if objQty.value = "" or objQty.value = "0" then
					alert("Enter Quantity")
					objQty.select
					exit function
				elseif objUoM.value = "select" then
					alert("Select Unit Or Measurement")
					objUoM.focus
					exit function
				elseif objReq.selectedIndex = "0" then
					alert("Select Required By")
					objReq.focus
					exit function
				end if
				ItemNode.Item(itr).Attributes.getNamedItem("UOM").value = objUoM.value
				ItemNode.Item(itr).Attributes.getNamedItem("QTY").value = objQty.value
				ItemNode.Item(itr).Attributes.getNamedItem("REQUIREDBY").value = objReq.value
				ItemNode.Item(itr).Attributes.getNamedItem("REMARKS").value = trim(objRem.value)

				sExp1 ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&"]/AddDet"
				Set ADNode = Root.Selectnodes(sExp1)

				if ADNode.Length <= 0 then
				    sAddSpcsCount = cdbl(sAddSpcsCount) + 1
				end if
			next
		end if

	'	if sAddSpcsCount > 0 then
	'	    if Not confirm("Add.Spec Details are not Specified for all Items. Do you want to Save?") then
	'		    exit function
	'	    end if
	'	end if

		sExp ="//HEADER"
		Set CheckNode = Root.Selectnodes(sExp)
		if CheckNode.Length = 0 then
			Set newElem = OutSelectData.createElement("HEADER")

			newElem.setAttribute "FORUNIT", document.formname.hUnit.value
			newElem.setAttribute "CREATEDON", document.formname.ctlCDDate.GetDate
			newElem.setAttribute "REMARKS", trim(document.formname.txtRemarks.value)
			newElem.setAttribute "APPROVER", document.formname.selApprover.value
			newElem.setAttribute "CREATEDBY", document.formname.hCreatedBy.value
			newElem.setAttribute "RECEIPTNO", iRcptNo
			newElem.setAttribute "COSTCENTER", trim(document.formname.selCC.value)
			newElem.setAttribute "ITEMTYPE", trim(document.formname.hItemType.value)
			newElem.setAttribute "ACCHEAD", trim(document.formname.selAccHead.value)
			newElem.setAttribute "REFTYPE", ""
			sRefType = document.formname.selRefName(document.formname.selRefName.selectedIndex).value
		    if trim(sRefType)<>"N" then
		        newElem.setAttribute "AppRefType", sRefType
		    else
		        newElem.setAttribute "AppRefType",""
		    end if
		    newElem.setAttribute "AppRefNo", document.formname.hRefNo.value
		    newElem.setAttribute "AppRefDate", document.formname.hRefDate.value
		    newElem.setAttribute "CallFrom", "MR"
		    newElem.setAttribute "RedirectTo","MRSMGMTLIST.ASP?HCHECK=M"

		    sApprover = document.formname.selApprover(document.formname.selApprover.selectedIndex).value
		    if trim(sApprover)="IM" then
		        newElem.setAttribute "ImmediateApprover","Y"
		    else
		        newElem.setAttribute "ImmediateApprover","N"
		    end if
            newElem.setAttribute "MRNo",""
            newElem.setAttribute "RequestedByUnit",document.formname.hRequestedByUnit.value
            newElem.setAttribute "ISSTOTYPE",document.formname.hIssueToType.value
            newElem.setAttribute "ISSTOCODE",document.formname.hIssueToCode.value
            newElem.setAttribute "ISSTOSUBCODE",document.formname.hIssueToSubCode.value
            newElem.setAttribute "ISSUETYPECODE",sRequstFor
			Root.appendChild newElem
		elseif CheckNode.Length > 0 then
			CheckNode.Item(0).Attributes.getNamedItem("FORUNIT").value = document.formname.hUnit.value
			CheckNode.Item(0).Attributes.getNamedItem("CREATEDON").value = document.formname.ctlCDDate.GetDate
		'	if document.formname.chkReqType.checked = true then
		'		CheckNode.Item(0).Attributes.getNamedItem("TYPE").value = "0"
		'	else
		'		CheckNode.Item(0).Attributes.getNamedItem("TYPE").value = "1"
		'	end if

			CheckNode.Item(0).Attributes.getNamedItem("REMARKS").value = trim(document.formname.txtRemarks.value)
			CheckNode.Item(0).Attributes.getNamedItem("APPROVER").value = document.formname.selApprover.value
			CheckNode.Item(0).Attributes.getNamedItem("CREATEDBY").value = document.formname.hCreatedBy.value
			CheckNode.Item(0).Attributes.getNamedItem("RECEIPTNO").value = iRcptNo
			CheckNode.Item(0).Attributes.getNamedItem("COSTCENTER").value = trim(document.formname.selCC.value)
			CheckNode.Item(0).Attributes.getNamedItem("ITEMTYPE").value = trim(document.formname.hItemType.value)
			CheckNode.Item(0).Attributes.getNamedItem("ACCHEAD").value = trim(document.formname.selAccHead.value)
			CheckNode.Item(0).Attributes.getNamedItem("ISSUEFOR").value =""
			'CheckNode.Item(0).Attributes.getNamedItem("REFTYPE").value =  trim(document.formname.SelRefType.value)
			CheckNode.Item(0).Attributes.getNamedItem("REFTYPE").value =  ""
			CheckNode.Item(0).Attributes.getNamedItem("PARTYCODE").value = trim(document.formname.hPartyCode.value)
			sRefType = document.formname.selRefName(document.formname.selRefName.selectedIndex).value
		    if trim(sRefType)<>"N" then
		        CheckNode.Item(0).Attributes.getNamedItem("AppRefType").value = sRefType
		    else
		        CheckNode.Item(0).Attributes.getNamedItem("AppRefType").value = ""
		    end if
		    CheckNode.Item(0).Attributes.getNamedItem("AppRefNo").value = document.formname.hRefNo.value
		    CheckNode.Item(0).Attributes.getNamedItem("AppRefDate").value = document.formname.hRefDate.value
		    CheckNode.Item(0).Attributes.getNamedItem("CallFrom").value = "MR"
		    CheckNode.Item(0).Attributes.getNamedItem("RedirectTo").value = "MRSMGMTLIST.ASP?HCHECK=M"

		    sApprover = document.formname.selApprover(document.formname.selApprover.selectedIndex).value
		    if trim(sApprover)="IM" then
		        CheckNode.Item(0).Attributes.getNamedItem("ImmediateApprover").value = "Y"
		    else
		        CheckNode.Item(0).Attributes.getNamedItem("ImmediateApprover").value = "N"
		    end if
		    CheckNode.Item(0).Attributes.getNamedItem("MRNo").value = ""
		    CheckNode.Item(0).Attributes.getNamedItem("RequestedByUnit").value = document.formname.hRequestedByUnit.value
		    CheckNode.Item(0).Attributes.getNamedItem("ISSTOTYPE").value = document.formname.hIssueToType.value
            CheckNode.Item(0).Attributes.getNamedItem("ISSTOCODE").value =document.formname.hIssueToCode.value
            CheckNode.Item(0).Attributes.getNamedItem("ISSTOSUBCODE").value = document.formname.hIssueToSubCode.value
            CheckNode.Item(0).Attributes.getNamedItem("ISSUETYPECODE").value = sRequstFor
		end if
	end if

	'alert(OutData.xml)
	'exit function
	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLSave.asp?Name=MRS&SessionFlag=true", false
	objhttp.send OutSelectData.XMLDocument

	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	document.formname.action = "MRGenerationInsert.asp"
	document.formname.submit
end Function

Function PopDone()
	Set Root = OutData.documentElement

	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","ConsumptionHeadInsert.asp", false
	objhttp.send OutData.XMLDocument
'	alert(objhttp.responseText)
	if Root.HaschildNodes() then
		For Each HeaderNode In Root.childNodes
			if HeaderNode.Attributes.Item(3).nodeValue = "N" then
				document.formname.selAccHead.length = document.formname.selAccHead.length+1
				document.formname.selAccHead.options(document.formname.selAccHead.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
				document.formname.selAccHead.options(document.formname.selAccHead.length-1).Value = cstr(HeaderNode.Attributes.Item(0).nodeValue)
			end if
		next
	end if
	hideDiv
end Function

Function CheckEntry()
	if trim(document.formname.txtCHead.value) = "" then
		alert("Enter Consumption Head")
		document.formname.txtCHead.select()
		exit Function
	elseif document.formname.selAcc.selectedIndex = "0" then
		alert("Select Account Head")
		document.formname.selAcc.focus()
		exit Function
	else
		sAHead = document.formname.selAcc.value
		sDesc = trim(document.formname.txtCHead.value)
		sIssTo = document.formname.hIssueToCode.value

		Set Root = OutData.documentElement
		sExp ="//AccountHead [ @ACCHEAD = "&sAHead&" and @CONSUM = '"&sDesc&"' and @ISSFOR = '"&sIssFor&"']"
		Set CheckNode = Root.Selectnodes(sExp)
		if CheckNode.Length > 0 then
			alert("Already Consumption Head exits")
			document.formname.txtCHead.select()
			exit function
		else
			Set newElem = OutData.createElement("AccountHead")
			newElem.setAttribute "ACCHEAD", sAHead
			newElem.setAttribute "CONSUM", sDesc
			newElem.setAttribute "ISSFOR", sIssTo
			newElem.setAttribute "SRC", "N"

			Root.appendChild newElem

			set oRow = document.all.tblData.insertRow(document.all.tblData.rows.length)

			set headerCell=oRow.insertCell()
			headerCell.innerHTML= trim(document.formname.txtCHead.value)
			headerCell.className="ExcelDisplayCell"
			headerCell.align="left"

			set headerCell=oRow.insertCell()
			headerCell.innerHTML= document.formname.selAcc(document.formname.selAcc.selectedIndex).text
			headerCell.className="ExcelDisplayCell"
			headerCell.width="10"

			document.formname.selAcc.selectedIndex = "0"
			document.formname.txtCHead.value = ""

		end if
	end if
end Function

</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<%
	dim sUnit,sIType,sCreatedBy,iuserid,rsUser
	dim sFinPeriod,Arr,dFrmDate,dToDate,sQuery,dcrs

	sUnit = Request.QueryString("sOrg")
	sIType = Request.QueryString("sIType")

	if trim(sUnit)="" or IsNull(sUnit) then sUnit = session("organizationcode")

	sFinPeriod = session("Finperiod")
	Arr = split(sFinPeriod,":")
	dFrmDate = "01/04/"& Arr(0)
	dToDate = "31/03/"& Arr(1)

	iuserid = Session("userid")
	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	set dcrs = Server.CreateObject("ADODB.Recordset")

    sCreatedBy = session("username")


%>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onLoad="popCC();Init()">
<form method="POST" name="formname">
	<input type=hidden name="hCreatedBy" value="<%=Session("userid")%>">
	<input type=hidden name="hFrmDate" value="<%=dFrmDate%>">
	<input type=hidden name="hToDate" value="<%=dToDate%>">
	<input type=hidden name="hUnit" value="<%=sUnit%>">
	<input type=hidden name="hItemType" value="">
	<input type=hidden name="hIssTo" value="">
	<input type=hidden name="hIssForType" value="">
	<input type=hidden name="hPartyCode" value="">
	<input type=hidden name="hRefType" value="">
	<input type=hidden name="hRefNo" value="">
	<input type=hidden name="hRefDate" value="">
	<input type="hidden" name="hRequestedByUnit" value="<%=sUnit%>">

	<input type="hidden" name="hIssueToType" value="">
	<input type="hidden" name="hIssueToCode" value="">
	<input type="hidden" name="hIssueToSubCode" value="">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Material Requisition
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
				    		    <tr>
						<td height="20" valign="bottom">
							<table border="0" cellpadding="0" cellspacing="0" >
								<tr>
								   	<td class="TabCell" valign="bottom" align="center" width="50">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr><a href="MRSMGMTList.asp">
												<td align="center">List
												</td></a>
											</tr>
										</table>
									</td>
									<td class="TabCurrentCell" valign="bottom" width="90">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr><a href="MRGENERATIONENTRY.asp">
												<td align="center">Basic
												</td></a>
											</tr>
										</table>
									</td>
									<td class="TabCell" valign="bottom" width="145">
									    <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										    <tr><a href="MRApprovalEntry.asp">
											    <td align="center">Edit/Approval
											    </td></a>
										    </tr>
									    </table>
								    </td>
									<td class="TabCellEnd" valign="bottom" align="left">
										    &nbsp;
								    </td>
								</tr>
							</table>
						</td>
                	</tr>
					<tr>
						<td class="TabBody">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="4" class="MiddlePack">
									</td>
								</tr>
								<tr>
									<td align="center">
									</td>
									<td width="100%" colspan="2">
										<div align="left">
											<table border="0" cellspacing="0" cellpadding="0" width="100%">
											    <tr>
											        <td class="FieldCellSub" style="width:125px" >Requested By</td>
													<td class="FieldCellSub">
													    <select name="selIssueTo" class="FormElem" onChange="popIssueTo()">
													        <option value="0">Select</option>
													        <%
														        populateIssueToSel(sUnit)
													        %>
													        </select>
													    <span id="txtParty" class="DataOnly"></span>
													</td>
													<td class="FieldCellSub">MR Date</td>
													<td class="FieldCellSub" valign="middle">
														<object id="ctlCDDate" onBlur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"    codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="FormElem" viewastext>
															<param name="_ExtentX" value="2355">
															<param name="_ExtentY" value="529">
														</object>
													</td>
											    </tr>
												<tr>
                                                    <td class="FieldCellSub" style="width:125px">Reference Name</td>
													<td class="FieldCellSub">
														<select name="selRefName" class="FormElem">
														<%
														    RefTypePop 2,4
														%>
														</select>
													    <a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click Here to Edit Usage Information" width="11" height="11" onClick="GetDetails()"></a>
													</td>
													<td class="FieldCellSub">Requested For</td>
                                                    <td class="FieldCellSub">
                                                        <select id="cmbIssType" class="FormElem">
                                                            <option value="SEL">Select</option>
                                                            <%
                                                                sQuery = "Select ReceiptIssueTypeCode,ReceiptIssueTypeDesc from APP_M_ReceiptIssueTypes where ApplicableFor in ('B','I')"
                                                                dcrs.open sQuery,con
                                                                if not dcrs.eof then
                                                                    do while not dcrs.eof
                                                                            response.write "<option value="& trim(dcrs(0)) &">"& trim(dcrs(1)) &"</option>"
                                                                        dcrs.movenext
                                                                    loop
                                                                end if
                                                                dcrs.close
                                                            %>
                                                        </select>&nbsp;
                                                    </td>
												</tr>
                                                   <tr>
                                                    <td class="FieldCellSub" style="width:125px">Reference No - Date</td>
													<td class="FieldCellSub">

														<span class="DataOnly" align=center id="RefNoDate">NA</span>
												    </td>
												    <td class="FieldCellSub">Created By</td>
													<td class="FieldCellSub">
														<span class="dataonly"><%=sCreatedBy%></span>
													</td>
												</tr>
												<tr>
													<td class="FieldCellSub">Acc. Head</td>
													<td class="FieldCellSub">
														<select size="1" name="selAccHead" class="FormElem" onChange="CreateNew(this)">
															<option value="select">Select</option>
															<option value="NEW">< NEW ></option>
														</select>
														<!--Div to display the Consumption head mapping pop up layer-->
														<div class=frmbody id="idConsumption" style="Z-INDEX: 1; POSITION: absolute" style="width=350;display: none" >
															<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
															<tr>
																<td valign="top">
																	<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
																		<TR>
																			<TD class="TabBodyWithTopLine">
																				<table border="0" cellpadding="0" cellspacing="0" width="100%">
																					<tr><td height="1px"></td></tr>
																					<tr>
																						<td colspan=2 class="PageTitle">Consumption Head Configuration</td>
																					</tr>
																					<tr>
																						<td align="center"></td>
																						<td width="100%">
																							<table border="0" cellpadding="0" cellspacing="0">
																								<tr>
																									<td class="FieldCell">Consumption Head</td>
																									<td class="FieldCellSub">
																										<input type="text" name="txtCHead" size="20" maxlength=50 class="FormElem">
																									</td>
																								</tr>
																								<tr>
																									<td class="FieldCell" valign="top">Account Head</td>
																									<td class="FieldCellSub">
																										<select size="1" name="selAcc" class="FormElem">
																											<option value="select">Select</option>
																											<%	'Calling the Function which populates Account Head List
																												populateAccountHead
																											%>
																										</select>
																										<input type="button" value=" Add " name="B3" class="AddButtonX" onClick="CheckEntry()">
																									</td>
																								</tr>
																							</table>
																						</td>
																						<td align="center"></td>
																					</tr>
																					<tr>
																						<td align="center" colspan="3" class="MiddlePack"></td>
																					</tr>
																					<tr>
																						<td align="center"></td>
																						<td width="100%">
																							<div class="frmbody" id="frm2" style="width: 100%; height:130;">
																								<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
																									<tr>
																										<td class="ExcelHeaderCell" align="center">Consumption</td>
																										<td class="ExcelHeaderCell" align="center" width="100">Account Head</td>
																									</tr>
																								</table>
																							</div>
																						</td>
																						<td align="center"></td>
																					</tr>

																					<tr>
																						<td align="center" colspan="3" class="MiddlePack"></td>
																					</tr>

																					<tr>
																						<td align="center">
																						    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
																						</td>
																						<td valign="top">
																							<table border="0" cellpadding="0" cellspacing="0" width="100%">
																								<tr>
																									<td valign="middle" class="ActionCell">
																										    <input type="button" value="Done" name="B1" class="ActionButton" onClick="PopDone()">
																										    <input type="button" value="Cancel" name="B2" class="ActionButton" onClick="hideDiv()">
																									</td>
																								</tr>
																							</table>
																						</td>
																						<td align="center">
																						    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
																						</td>
																					</tr>
																					<tr>
																						<td align="center" colspan="3" class="BottomPack"></td>
																					</tr>
																				</table>
																			</td>
																		</tr>
																	</table>
																</td>
															</tr>
															</table>
														</div>
													</td>
													<td class="FieldCellSub" style="width:125px">Cost Center</td>
													<td class="FieldCellSub" valign="top">
														<select size="1" name="selCC" class="FormElem">
															<option value="select">Select</option>
														<%	'Calling the Function which populates Cost Center List
															populateCostCenter
														%>
														</select>
													</td>

													<!--<td class="FieldCellSub">Type</td>
													<td class="FieldCell" valign="top">
														<input type="checkbox" name="chkReqType" class="FormElem" value="Returnable">Returnable
													</td>-->
												</tr>
											</table>
										</div>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="MiddlePack"></td>
								</tr>

								<tr>
									<td align="center"></td>
									<td width="100%" colspan="2">
										<div class="frmBody" id="frm1" style="width: 100%; height:280;">
											<table border="0" cellspacing="1" class="ExcelTable" width="100%" id=tblLot>
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10">
														<p align="center">S.No.
													</td>
													<td class="ExcelHeaderCell" align="center">
														<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15" alt="Delete's the Selected Item (s)" height="15" onClick="DeleteItems()"></a>
													</td>
													<td class="ExcelHeaderCell" align="center" >
														Item Description
														<!--<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click here to Select Item (s)" width="11" height="11" onClick="GetItems('<%=FormatDate(date)%>')"></a>-->
													</td>
													<td class="ExcelHeaderCell" align="center">Quantity</td>
													<td class="ExcelHeaderCell" align="center" >UoM</td>
													<td class="ExcelHeaderCell" align="center" >Required By</td>
													<!--<td class="ExcelHeaderCell" align="center" >Add Spec</td>-->
													<td class="ExcelHeaderCell" align="center" >Stock</td>
													<td class="ExcelHeaderCell" align="center" >Remarks</td>
												</tr>
											</table>
										</div>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="MiddlePack">
									</td>
								</tr>

								<tr>
								    <td align="center"></td>
									<td class="FieldCellSub"> Approver</td>
									<td class="FieldCellSub">
                                      <select size="1" name="selApprover" class="FormElem">
											<option value="0">Select</option>
											<option value="IM">Immediate Approver</option>
											<%	'Calling the Function which populates the User list
												populateEmployee
											%>
										</select>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
								    <td align="center"></td>
									<td class="FieldCellSub">Remarks</td>
									<td class="FieldCellSub">
										<textarea name="txtRemarks" cols="100" class="FormElem"></textarea>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top" colspan="2">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<input type="button" value="Save" name="BtnSubmit" class="ActionButton" onClick="CheckSubmit('<%=FormatDate(date)%>')">
 													<input type="reset" value="Reset" name="B1" class="ActionButton">
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="BottomPack">
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
</body>
</html>

<%
	' Function to populate Usage
	Function populateUsage()
		' Declaration of variables
		Dim dcrs,sUsageCode,sUsageDesc
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISSUEDFORCODE,ISSUEDFORDESCRIPTION FROM INV_M_ISSUEDFOR WHERE ISSUEDFORCODE <> 'INV' ORDER BY ISSUEDFORCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sUsageCode = dcrs(0)
		set sUsageDesc = dcrs(1)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUsageCode)&""">"&trim(sUsageDesc)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>
<%
	' Function to populate the Cost Center list
	Function populateCostCenter()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT COSTCENTERHEAD,CCACCOUNTDESCRIPTION FROM VWORGCOSTCENTER WHERE OUDEFINITIONID = " & Pack(sUnit) & " AND USEABLE = 1 ORDER BY COSTCENTERHEAD"
			.ActiveConnection = con
			.Open
		end with
		set stypID = dcrs(0)
		set stypName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function
%>

<%
	' Function to populate the Account Head list
	Function populateAccountHead()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT DISTINCT ACCOUNTHEAD,ACCOUNTDESCRIPTION,ACCOUNTHEADCODE FROM VWORGGLHEADS WHERE OUDEFINITIONID = " & Pack(sUnit) & " AND ACCOUNTHEAD IN (SELECT ACCOUNTHEAD FROM ACC_R_GLACCAPPLICATIONS WHERE AVAILABLEINAPPLN IN (4,5,6) AND OUDEFINITIONID = " & Pack(sUnit) & ") ORDER BY 2"
			.ActiveConnection = con
			.Open
		end with
		set stypID = dcrs(0)
		set stypName = dcrs(2)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function

	Function Issue()
		MsgBox "ok"
	End Function
%>
