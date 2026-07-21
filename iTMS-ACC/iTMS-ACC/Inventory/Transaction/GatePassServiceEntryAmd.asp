<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GatePassServiceEntry.asp
	'Module Name				:	Gate Pass - Service
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	APRIL 05,2010
	'Modified On				:	Jan 06,2011
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
<!-- #include File="../../include/CheckPrevFinYear.asp" -->
<!-- #include File="../../include/populate.asp" -->
<%
	Dim oDOM,Root,nodeDetail,objRs,objRs1,hRoot,suppNode
	Dim sQuery,sRemarks,sTransport,sDelivery,sTakenBy,sForUnit,sItemType,sSupAgent,sDesc
	Dim nGatePassNo
	Dim sItemCode,sClassCode,sItemDesc,sOtherDesc,sReason
	nGatePassNo= Request.QueryString("GatePassNo")
	sForUnit = Session("organizationcode")

	set oDOM = CreateObject("Microsoft.XMLDOM")
	set objRs = server.CreateObject("ADODB.Recordset")
	set objRs1 = server.CreateObject("ADODB.Recordset")

	sQuery = "Select GATEPASSNO,ORGANISATIONCODE,INVOICETYPE,PARTYCODE,isNull(TYPEOFITEMS,''),APPLICATIONCODE,"&_
			 "MARKEDON,isNull(REMARKS,''),STATUS,isNull(Transport,''),isNull(TakenBy,''),isNull(DeliveryBy,''),DCCODE from  FORGATEPASSHEADER "&_
			 "WHERE GatePassNo ="& nGatePassNo

	'Response.Write sQuery
	set Root = oDOM.createElement("Root")
			oDOM.appendChild Root
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = sQuery
		.Open
	end with
	if not objrs.EOF then
		set hRoot = oDOM.createElement("HEADER")
			hRoot.setAttribute "ITEMTYPE",objRs(4)
			hRoot.setAttribute "FORUNIT",objrs(1)
			hRoot.setAttribute "REMARKS",objrs(7)
			hRoot.setAttribute "SUPPAGENT",objrs(3)
			hRoot.setAttribute "Transport",objRs(9)
			hRoot.setAttribute "TakenBy",objrs(10)
			hRoot.setAttribute "DeliveryBy",objrs(11)
			Root.appendChild hRoot
			sSupAgent = objrs(3)
	end if
	objrs.Close
	'sQuery = "Select OrgnPartyCode,SupplierCode,PartyName,PartyCode,PartyType,PartySubType from vwSupplierAddress where PartyCode = "&sSupAgent
	sQuery = "Select OrgnPartyCode,B.PartyCode,PartyName,A.PartyCode,PartyType,PartySubType from APP_M_PartyMaster A,APP_R_OrgParty B where A.PartyCode = B.PartyCode and A.PartyCode = "& sSupAgent
	'Response.Write sQuery
	with objrs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source =sQuery
		.Open
	end with
	if not objrs.EOF then
		set suppNode = oDOM.createElement("Supplier")
			suppNode.setAttribute "SuppShortCode",objrs(0)
			suppNode.setAttribute "SuppCode",objrs(1)
			suppNode.setAttribute "SuppName",objrs(2)
			suppNode.setAttribute "AgentCode","N"
			suppNode.setAttribute "AgentName",""
			suppNode.setAttribute "PartyCode",objrs(3)
			suppNode.setAttribute "PartyType",objrs(4)
			suppNode.setAttribute "PartySubType",objrs(5)
		Root.appendChild suppNode
	else
			with objRs1
				.CursorLocation = 3
				.CursorType = 3
				.ActiveConnection = con
				.Source ="Select PartyType,PartySubType,PartyCode,PartyName from vwOrgParty where PartyCode = "& sSupAgent
				.Open
			end with
			if not objRs1.EOF then
				set suppNode = oDOM.createElement("Supplier")
					suppNode.setAttribute "SuppShortCode",""
					suppNode.setAttribute "SuppCode",""
					suppNode.setAttribute "SuppName",objrs1(3)
					suppNode.setAttribute "AgentCode","N"
					suppNode.setAttribute "AgentName",""
					suppNode.setAttribute "PartyCode",objrs1(2)
					suppNode.setAttribute "PartyType",objrs1(0)
					suppNode.setAttribute "PartySubType",objrs1(1)
				Root.appendChild suppNode
			end if
			objRs1.Close
	end if
	objrs.Close
	sQuery = "Select GATEPASSNO,ENTRYNO,isNull(ITEMCODE,0),isNull(CLASSIFICATIONCODE,0),QUANTITY,isNull(DESCRIPTION,''),"&_
			 "INVOICEDUOM,isNull(ItemValue,0),isNull(FormJJ,'N'),isNull(Reason,'') from FORGATEPASSDETAILS Where GatePassNo= "& nGatePassNo

	'Response.Write sQuery
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = sQuery
		.Open
	end with
	if not objrs.EOF then
		do while not objrs.EOF
			SET nodeDetail = oDOM.createElement("DETAILS")
				sItemCode = objrs(2)
				sClassCode = objrs(3)
				sOtherDesc = objrs(5)
				sReason = objRs(9)
				if sItemCode <>"" and sClassCode<>"" then
					with objRs1
						.CursorLocation = 3
						.CursorType = 3
						.ActiveConnection = con
						.Source = "Select ItemDescription from vwItem where ItemCode ="& sItemCode  &"  and ClassificationCode = "&sClassCode
						.Open
					end with
					if not objRs1.EOF then
						sItemDesc = objrs1(0)
						sItemDesc = Replace(sItemDesc,"'","")
					end if
					objRs1.Close
				end if 'if sItemCode <>"" and sClassCode<>"" then
				if sItemDesc <>"" and sOtherDesc<>"" then
					sDesc= sItemDesc  &"-"& sOtherDesc
				elseif sItemDesc <>"" and trim(sOtherDesc)="" then
					sDesc= sItemDesc
				else
					sDesc = sOtherDesc
				end if
			nodeDetail.setAttribute "OTHERDESC",objrs(5)
			nodeDetail.setAttribute "QTY",objrs(4)
			nodeDetail.setAttribute "ITEMCODE",sItemCode
			nodeDetail.setAttribute "CLASSCODE",sClassCode
			nodeDetail.setAttribute "UOM",objrs(6)
			nodeDetail.setAttribute "VALUE",objrs(7)
			nodeDetail.setAttribute "FORMJJ",objrs(8)
			nodeDetail.setAttribute "DESC",sDesc
			nodeDetail.setAttribute "ITEMDESC",sItemDesc
			nodeDetail.setAttribute "REASON",sReason
			Root.appendChild nodeDetail
			objrs.MoveNext
		loop
	end if
	objrs.Close

oDOM.save server.MapPath("../temp/transaction/GatePassServiceAmd.xml")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>Gate Pass</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<script type="application/xml" data-itms-xml-island="1" id="UoMData" data-src="../../inventory/xmldata/Uom.xml"></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData">
	<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="Data"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="newData" data-src="../temp/transaction/GatePassServiceAmd.xml"></script>
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
dim j,iCounter1,ssupcode,iser,iTempCode,Partyname,iRcptNo
j = 0
iser = 0
iCounter1 = 0
'---------------------------------------------------------------------

Function DelItem()

	dim nTempCtr,nRows,sRoot,NodeChk,HeaderNode,count,stemp
	Dim ret,schqty,itemqty,totpackqty
	count=0
	nRows=document.formname.hRows.value
	for nTempCtr=nRows to 1 step -1
		set objChk	= eval("document.formname.chkboxDel"&cstr(nTempCtr))
		if objchk.checked=true then

			stemp= split(eval("document.formname.chkboxDel"&cstr(nTempctr)).value,":")

		document.formname.hRows.value=document.formname.hRows.value-1
		count=count+1
			set sRoot=outData.documentElement
			for each HeaderNode in sRoot.ChildNodes
				if strcomp(HeaderNode.nodename,"DETAILS")=0 then
					if strcomp(HeaderNode.getAttribute("ITEMCODE"),stemp(0))=0 and strcomp(HeaderNode.getAttribute("CLASSCODE"),stemp(1))=0 and (strcomp(replace(HeaderNode.getAttribute("OTHERDESC")," ",""),replace(stemp(2)," ",""))=0 or strcomp(replace(HeaderNode.getAttribute("ITEMDESC")," ",""),replace(stemp(2)," ",""))=0)  then
						sRoot.removechild HeaderNode
					end if
				end if
			next
		end if
	next

	if count=0 then
		exit function
	end if
	ClearTableItem()
	DispItem()
End Function


'-----------------------------------------------------------------------
Function popSuppAgentOLD()

	OutValue = showModalDialog("../../sales/transaction/PartySelectionTrans.asp","","dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(OutValue,":")

	while UBound(arrTemp) = 0
		OutValue = showModalDialog("../../sales/transaction/PartySelectionTrans.asp?"&OutValue,"","dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
	wend
	if UBound(arrTemp) <= 1 then exit function

	ssupcode = arrTemp(1)
	Partyname = arrTemp(0)

	document.formname.txtRefName.value = Partyname

End Function

Function popSuppAgent()

Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth
    
    sTempValWindowSize = GetWindowSizeForPopup("2")
    sArrTempValWindowSize = split(sTempValWindowSize,":")
    sProgramName = sArrTempValWindowSize(0)
    sPopupHeight = sArrTempValWindowSize(1)
    sPopupWidth = sArrTempValWindowSize(2)

	sForUnit = document.formname.hForUnit.value+":O"
	nFlag = 2
	'note : value for hSelectMode is M( Multiple ) / S(Single)
'	set OutValue=showModalDialog("SupplierSelect.asp?OrgId="+sForUnit+"&hSelectMode=S&Flag="+cstr(nFlag) & "&OrderTo=S",OutData,"status:no")
'	'msgbox OutValue.xml
'
'
'	sAct = UCase(trim(OutValue.getAttribute("Action")))
'	sQuery = trim(OutValue.getAttribute("PassQuery"))
'	if ucase(trim(sAct)) <> "CLOSE" then
'		do while sAct <> "DONE"
'			set OutValue=showModalDialog("SupplierSelect.asp?" & sQuery,OutData,"status:no")
'			sAct = UCase(trim(OutValue.getAttribute("Action")))
'			sQuery = trim(OutValue.getAttribute("PassQuery"))
'
'			if ucase(Trim(sAct)) = "CLOSE" then exit do
'		loop
'	end if 'if ucase(trim(sAct)) <> "CLOSE" then
	'alert(OutValue.xml)
	
	set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="&sForUnit,OutData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sQuery = OutValue.getAttribute("PassQuery")
	    if OutValue.getAttribute("Action")="CLOSE" then exit function

		while OutValue.getAttribute("Action")<>"Done"
		set	OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,OutData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
		    sQuery = OutValue.getAttribute("PassQuery")
	        if OutValue.getAttribute("Action")="CLOSE" then exit function
		wend
		
	
	

	If not OutValue.hasChildNodes Then 	exit function


	For each Node2 in OutValue.childNodes
		if ucase(Node2.nodename) = ucase("Entry") then

			sSName		= Node2.getAttribute("RetField0")
			'ssupcode	= trim(Node2.getAttribute("SuppCode")) & "#" & trim(Node2.getAttribute("AgentCode"))
			ssupcode	= trim(Node2.getAttribute("RetField1"))
			document.formname.txtRefName.value = sSName


		end if
	Next
End Function

Function ClearTable()
	dim i
	for i=1 to document.all.tblLot.rows.length - 1
		document.all.tblLot.deleteRow(1)
	next
	j = 0
end Function

Function AddDetails()
		Dim sDesc
		'alert(document.formname.selUOM.disabled)

		set Root = OutData.documentElement
		if trim(document.formname.txtDesc.value) = "" and trim(document.formname.txtItemDesc.value) = "" then
			alert("Enter Description or Select Item...!")
			document.formname.txtDesc.select
			exit function
		elseif trim(document.formname.txtQty.value) = "" then
			alert("Enter Quantity")
			document.formname.txtQty.select
			exit function
		elseif not document.formname.selUOM.disabled and ( document.formname.selUOM.value = "select" or trim(document.formname.selUOM.value) = "" ) then
			alert("Select UOM")
			document.formname.selUOM.focus
			exit function
		else
			'alert(document.formname.txtItemDesc.value)
			If trim(document.formname.txtItemDesc.value) <> "" Then
				sDesc = ucase(Trim(document.formname.txtItemDesc.value))
			End If

			If trim(document.formname.txtDesc.value) <> "" Then
				If Trim(sDesc) = "" Then
					sDesc = ucase(trim(document.formname.txtDesc.value))
				Else
					sDesc = sDesc & " - " & ucase(trim(document.formname.txtDesc.value))
				End If
			End If

			sApplFormJJ = "N"
			if document.formname.ChkFormJJ.checked then
				sApplFormJJ = "Y"
			end if

			sExp = "//ROOT/DETAILS [@DESC = '" & replace(sDesc,"'"," ") & "']"

			set Node1 = Root.selectNodes(sExp)
			'alert Node1.length
			if Node1.length > 0 then
				alert("Description Already entered")
				document.formname.txtDesc.select
				exit function
			end if


			Set Node = OutData.createElement("DETAILS")
			Node.setAttribute "DESC",sDesc
			Node.setAttribute "QTY",document.formname.txtQty.value
			Node.setAttribute "ITEMCODE", document.formname.hItem.value
			Node.setAttribute "CLASSCODE", document.formname.hClass.value
			Node.setAttribute "ITEMDESC", ucase(Trim(document.formname.txtItemDesc.value))
			Node.setAttribute "OTHERDESC", ucase(trim(document.formname.txtDesc.value))
			Node.setAttribute "UOM", ucase(trim(document.formname.SelUOM.value))
			Node.setAttribute "VALUE",document.formname.txtValue.value
			Node.setAttribute "FORMJJ",sApplFormJJ
			Node.setAttribute "REASON",trim(document.formname.txtReason.value)
			Root.appendChild Node

		end if

		ClearTableItem()
		dispitem
End Function
'============================================
Function ClearTableItem()
	Dim iNum,sRoot,NewElem
	For iNum = 1 to document.all.tblDetails.rows.length - 1
		document.all.tblDetails.deleteRow(1)
	Next
End Function
'=============================================
Function CheckSubmit(todaysdate)
'	if document.formname.selUnit.selectedIndex = "0" then
'		alert("Select For Unit")
'		document.formname.selUnit.focus
'		exit function
'	else
'	if document.formname.selItmType.selectedIndex = "0" then
'		alert("Select Item Type")
'		document.formname.selItmType.focus
'		exit function
'	else
    if document.formname.txtRefName.value = "" then
		alert("Select Supplier")
		exit function
	elseif len(trim(document.formname.txtRemarks.value)) > 200 then
		alert("Remarks should be less than 200 characters")
		document.formname.txtRemarks.select
		exit function
	else
		itr = 0
		set root = OutData.DocumentElement
		sExp ="//DETAILS"
		Set ItemNode = Root.Selectnodes(sExp)
		if ItemNode.Length = 0 then
			alert("Enter Details")
			exit function
		else

			for nTempCtr = 0 to ItemNode.Length -1
				set ObjChk1 = eval("document.formname.chkbox" + trim(nTempCtr+1) )
				if ObjChk1.checked then
					ItemNode.item(nTempCtr).Attributes(8).value ="Y"
				else
					ItemNode.item(nTempCtr).Attributes(8).value ="N"
				end if
			next
		end if
		for each node in root.childNodes
			if strcomp(node.nodeName,"HEADER")= 0 then
				root.removechild node
			end if
		next

		Set newElem = OutData.createElement("HEADER")
		newElem.setAttribute "FORUNIT", document.formname.hForUnit.value
		newElem.setAttribute "ITEMTYPE",""' document.formname.selItmType.value
		newElem.setAttribute "SUPPAGENT", ssupcode
		newElem.setAttribute "REMARKS", trim(document.formname.txtRemarks.value)

		newElem.setAttribute "Transport", trim(document.formname.txtTransport.value)
		newElem.setAttribute "TakenBy", trim(document.formname.txtTakenBy.value)
		newElem.setAttribute "DeliveryBy", trim(document.formname.txtDeliveryBy.value)

		Root.appendChild newElem
	end if

	'alert(OutData.xml)
	'exit function

	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","GatePassServiceUpdate.asp?GatePassNo="&document.formname.hGatePassNo.value, false
	objhttp.send OutData.XMLDocument
	'alert(objhttp.responseText)
	'exit function

	if objhttp.responseText = "" then
		if confirm("Gate Pass for Service has been created. Do you want to create another one?") then
			window.location.href "../../Inventory/Transaction/GATEPASSSELECTION.ASP?SelSent=Y&InvoiceType=V"
		else
			window.location.href "../../Inventory/welcome_Inventory.asp"
		end if
	else
		alert(objhttp.responseText)
	end if

end Function

Function SelectItem()
Dim	sIType, iItem,iClass,objhttp
'	if document.formname.selUnit.value = "" then
'		alert("Select Unit")
'		document.formname.selUnit.focus
'		exit function
'	else
'	if document.formname.selItmType.value = "select" then
'		alert("Select Item Type")
'		document.formname.selItmType.focus
'		exit function
'	end if

Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth
    
    sTempValWindowSize = GetWindowSizeForPopup("1")
    sArrTempValWindowSize = split(sTempValWindowSize,":")
    sProgramName = sArrTempValWindowSize(0)
    sPopupHeight = sArrTempValWindowSize(1)
    sPopupWidth = sArrTempValWindowSize(2)


iStock="N"
nFlag="1"
	sUnit = document.formname.hForUnit.value
	sIType = ""'document.formname.selItmType.value
	'set ResData=showModalDialog("../transaction/ItemSelect.asp?orgID=" & sUnit & "&iType=" & sIType & "&hSelectMode=M&Flag="+cstr(nFlag),Data,"dialogHeight:650px;dialogWidth:730px;center:Yes;help:No;resizable:No;status:No")
	'set ResData=showModalDialog("../transaction/ItemSelect.asp?orgID=" & sUnit & "&iType=" & sIType & "&hSelectMode=R&Flag="+cstr(nFlag),Data,"dialogHeight:650px;dialogWidth:730px;center:Yes;help:No;resizable:No;status:No")
	set ResData=showModalDialog("../../Common/"&sProgramName&"?orgID=" & sUnit & "&Stock=" & iStock & "&iType=" & sIType & "&hSelectMode=R&Flag="+cstr(nFlag)&"&hDispButt=Y",Data,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	'alert(ResData.xml)
	sAct = UCase(trim(ResData.getAttribute("Action")))
	sQuery = trim(ResData.getAttribute("PassQuery"))
	if ucase(trim(sAct)) <> "CLOSE" then
		do while sAct <> "DONE"
			'set OutValue=showModalDialog("../transaction/ItemSelect.asp?" & sQuery,Data,"dialogHeight:650px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No")
			set OutValue=showModalDialog("../../Common/"&sProgramName&"?" & sQuery,Data,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			sAct = UCase(trim(ResData.getAttribute("Action")))
			if ucase(Trim(sAct)) = "CLOSE" then exit do
			sQuery = trim(ResData.getAttribute("PassQuery"))
		loop
	end if
	Set Root = Data.documentElement
	If not Root.hasChildNodes Then 	exit function
	if Root.hasChildNodes() then
		For Each HeaderNode In Root.childNodes
			'alert(HeaderNode.xml)
			document.formname.txtItemDesc.value = Trim(HeaderNode.getAttribute("ItemName"))
			document.formname.hItem.Value =  HeaderNode.getAttribute("ItemCode")
			document.formname.hClass.value =  HeaderNode.getAttribute("ClassCode")
			document.formname.SelUOM.value = Trim(HeaderNode.getAttribute("StoresUoM"))
			document.formname.selUOM.disabled = true
		next
	end if


end Function
'=====================================
Function Init()

Dim	Root,subNode,outRoot,outsubNode,j,sItemCode,sClassCode,suppCode
	j=0
	set Root = newData.documentElement
	set outRoot = outData.documentElement
	for each subNode in Root.childNodes
		if strcomp(subNode.nodeName,"HEADER")=0 then
			'document.formname.hForUnit.value = subNode.getAttribute("FORUNIT")
			'document.formname.selItmType.value = subNode.getAttribute("ITEMTYPE")
			document.formname.txtRemarks.value = subNode.getAttribute("REMARKS")
			document.formname.txtTransport.value = subNode.getAttribute("Transport")
			document.formname.txtDeliveryBy.value = subNode.getAttribute("TakenBy")
			document.formname.txtTakenBy.value = subNode.getAttribute("DeliveryBy")

			set  outsubNode = outData.createElement("HEADER")
			outsubNode.setAttribute "FORUNIT",subNode.getAttribute("FORUNIT")
			outsubNode.setAttribute "ITEMTYPE",subNode.getAttribute("ITEMTYPE")
			outsubNode.setAttribute "REMARKS",subNode.getAttribute("REMARKS")
			outsubNode.setAttribute "Transport",subNode.getAttribute("Transport")
			outsubNode.setAttribute "TakenBy",subNode.getAttribute("TakenBy")
			outsubNode.setAttribute "DeliveryBy",subNode.getAttribute("DeliveryBy")
			outRoot.appendChild outsubNode

		elseif strcomp(subNode.nodeName,"DETAILS")=0 then

			set  outsubNode = outData.createElement("DETAILS")
				outsubNode.setAttribute "OTHERDESC",subNode.getAttribute("OTHERDESC")
				outsubNode.setAttribute "QTY",subNode.getAttribute("QTY")
				outsubNode.setAttribute "ITEMCODE",subNode.getAttribute("ITEMCODE")
				outsubNode.setAttribute "CLASSCODE",subNode.getAttribute("CLASSCODE")
				outsubNode.setAttribute "UOM",subNode.getAttribute("UOM")
				outsubNode.setAttribute "VALUE",subNode.getAttribute("VALUE")
				outsubNode.setAttribute "FORMJJ",subNode.getAttribute("FORMJJ")
				outsubNode.setAttribute "ITEMDESC",subNode.getAttribute("ITEMDESC")
				outsubNode.setAttribute "DESC",subNode.getAttribute("DESC")
				outsubNode.setAttribute "REASON",subNode.getAttribute("REASON")
				sItemCode = subNode.getAttribute("ITEMCODE")
				sClassCode = subNode.getAttribute("CLASSCODE")
			outRoot.appendChild outsubNode

		elseif StrComp(subNode.nodeName,"Supplier")=0 then
			set outsubNode = outData.CreateElement("Supplier")
				outsubNode.setAttribute "SuppShortCode",subNode.getAttribute("SuppShortCode")
				outsubNode.setAttribute "SuppCode",subNode.getAttribute("SuppCode")
				outsubNode.setAttribute "SuppName",subNode.getAttribute("SuppName")
				outsubNode.setAttribute "AgentCode",subNode.getAttribute("AgentCode")
				outsubNode.setAttribute "AgentName",subNode.getAttribute("AgentName")
				outsubNode.setAttribute "PartyCode",subNode.getAttribute("PartyCode")
				outsubNode.setAttribute "PartyType",subNode.getAttribute("PartyType")
				outsubNode.setAttribute "PartySubType",subNode.getAttribute("PartySubType")
				suppCode = subNode.getAttribute("SuppCode")
			outRoot.appendChild outsubNode
		end if
	next
	for each subNode in outRoot.childNodes
		if strcomp(subNode.nodeName,"HEADER")=0 then
			ssupcode = suppCode
			subNode.setAttribute "SUPPAGENT",suppCode
		end if
	next
	dispitem()
End Function
'===============================================================
Function dispitem()

Dim outRoot,subNode,j,sDesct,sApplFormJJ,sReason
j =0
	set outRoot = outData.documentElement
	for each subnode in outRoot.childNodes
		if strcomp(subNode.nodeName,"DETAILS")=0 then
			set oRow = document.all.tblDetails.insertRow(j+1)
			sItemCode = subNode.getAttribute("ITEMCODE")
			sClassCode = subNode.getAttribute("CLASSCODE")
			sApplFormJJ = subNode.getAttribute("FORMJJ")
			sDesct = subNode.getAttribute("DESC")
			sReason = subNode.getAttribute("REASON")

				set headerCell=oRow.insertCell()
				headerCell.innerHTML=j+1
				headerCell.className="ExcelSerial"
				headerCell.align="center"

				set headerCell=oRow.insertCell()
				headerCell.className="ExcelDisplayCell"
				set oText = document.createElement("<input type=""checkbox"" value="&sItemCode&":"&sClassCode&":"&replace(sDesct," ","")&" name=""chkboxDel"& j+1 &""" class=""FormElem"">" )
				headerCell.appendChild(oText)

				set headerCell=oRow.insertCell()
				headerCell.innerHTML= sDesct
				headerCell.className="ExcelDisplayCell"
				headerCell.align="left"

				set headerCell=oRow.insertCell()
				headerCell.innerHTML=subNode.getAttribute("QTY")
				headerCell.className="ExcelDisplayCell"
				headerCell.align="right"


				set headerCell=oRow.insertCell()
				headerCell.innerHTML=subNode.getAttribute("UOM")
				headerCell.className="ExcelDisplayCell"
				headerCell.align="right"

				set headerCell=oRow.insertCell()
				headerCell.innerHTML=subNode.getAttribute("VALUE")
				headerCell.className="ExcelDisplayCell"
				headerCell.align="right"

				set headerCell=oRow.insertCell()
				if sApplFormJJ = "Y" then
					set oText = document.createElement("<input type=""checkbox"" value=""Y"" name=""chkbox"& j+1 &""" class=""FormElem"" checked >" )
				else
					set oText = document.createElement("<input type=""checkbox"" value=""Y"" name=""chkbox"& j+1 &""" class=""FormElem"" >" )
				end if
				headerCell.appendChild(oText)

				headerCell.className="ExcelDisplayCell"
				headerCell.align="right"

				set headercell=oRow.insertCell()
				headercell.innerText = sReason
				headercell.className="ExcelDisplayCell"
				headercell.align="left"

				j = j + 1

				document.formname.txtDesc.value = ""
				document.formname.txtQty.value = ""

				document.formname.txtItemDesc.value = ""
				document.formname.hItem.Value =  ""
				document.formname.hClass.value =  ""
				document.formname.selUOM.value  = "select"
				document.formname.txtReason.value ="SENT FOR REPAIRS - TO BE RETURNED"
				document.formname.txtValue.value = ""
				document.formname.ChkFormJJ.checked = false
		elseif strcomp(subNode.nodeName,"Supplier")=0 then
			document.formname.txtRefName.value = subNode.getAttribute("SuppName")
		end if ' if strcomp(subNode.nodeName,"DETAILS")=0 then
	next
	document.formname.hRows.value=j
End Function
'=========================================================
</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onLoad="Init()">
	<form method="POST" name="formname">
	<input type=hidden name="hItem" value="">
	<input type=hidden name="hClass" value="">
	<input type=hidden name="hForUnit" value="<%=sForUnit%>">
	<input type=hidden name="hItemType" value="<%=sItemType%>">
	<input type=hidden name="hRemarks" value="<%=sRemarks%>">
	<input type=hidden name="hTransport" value="<%=sTransport%>">
	<input type=hidden name="hDeliveryBy" value="<%=sDelivery%>">
	<input type=hidden name="hTakenBy" value="<%=sTakenBy%>">
	<input type=hidden name="hPartyCode" value="<%=sSupAgent%>">
	<input type=hidden name="hRows" value="">
	<input type=hidden name="hGatePassNo" value="<%=nGatePassNo%>">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Gate Pass - Service
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
						<td class="TabBodyWithTopLine">
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
											<table border="0" cellspacing="0" cellpadding="0" width="572">
												<!--<tr>
													<td class="FieldCellSub">For Unit</td>
													<td class="FieldCellSub">
														<select size="1" name="selUnit" class="FormElem">
															<option value="select">Select</option>
															<%	'Calling the Function which populates Organization Unit list
																populateUnit
															%>
														</select>
													</td>
												</tr>-->

												<!--<tr>
													<td class="FieldCellSub">Item Type</td>
													<td class="FieldCellSub" valign="top">
                                                        <select size="1" name="selItmType" class="FormElem">
															<option value="select">Select</option>
															<%	'Calling the Function which populates the Item Type list
																'populateItemType
															%>
														</select>
													</td>
												</tr>-->

												<tr>
													<td class="FieldCellSub">Party</td>
													<td class="FieldCellSub">
														<input type="text" name="txtRefName" value size="60" class="formelemread" readonly>&nbsp;
														<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click here to Select Party" width="11" height="11" onClick="popSuppAgent()"></a>
													</td>
												</tr>
												<tr>
													<td class="FieldCellSub">Item Description</td>
													<td class="FieldCellSub">
														<input type="text" name="txtItemDesc" value size="60" class="formelemread" readonly>&nbsp;
														<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="center" alt="Click here to Select Item" width="11" height="11" onClick="SelectItem()"></a>
													</td>
												</tr>

												<tr>
													<td class="FieldCellSub">Other Description</td>
													<td class="FieldCellSub">
														<input type="text" name="txtDesc" maxlength=50 size="60" class="formelem">
													</td>
												</tr>

												<tr>
													<td class="FieldCellSub">Reason</td>
													<td class="FieldCellSub">
														<input type="text" name="txtReason" maxlength=35 size="60" class="formelem" value="SENT FOR REPAIRS - TO BE RETURNED">
													</td>
												</tr>

												<tr>
													<td class="FieldCellSub">Quantity & Value</td>
													<td class="FieldCellSub">
														<input type="text" name="txtQty" value size="12" class="formelem" onkeypress="DoKeyPress('Y',7,3)">&nbsp;
														<select size="1" name="selUOM" class="FormElem">
															<option value="select">Select</option>
															<%
															populateUoM()
															%>
														</select>
														&nbsp;
														<input type="text" name="txtValue" size="12" class="formelem" onkeypress="DoKeyPress('Y',7,2)" >&nbsp;
													</td>
												</tr>
												<tr>
													<td class="FieldCellSub">Form JJ Applicable</td>
													<td class="FieldCellSub">
														<input type="checkbox" name="ChkFormJJ" value="Y" class="formelem" >
														<input type="button" value="Add" name="B3" class="AddButtonX" onclick = "AddDetails()">
													</td>
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
										<div class="frmBody" id="frm1" style="width: 585; height:230;">
											<table border="0" cellspacing="1" class="ExcelTable" width="585" id=tblDetails>
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10">
														<p align="center">S.No.
													</td>
													<td class="ExcelHeaderCell" align="center" width=10>
														<img border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15" height="15" onclick="DelItem()">
													</td>
													<td class="ExcelHeaderCell" align="center" width=500>
														Item Description
													</td>
													<td class="ExcelHeaderCell" align="center">Quantity</td>
													<td class="ExcelHeaderCell" align="center">UoM</td>
													<td class="ExcelHeaderCell" align="center">Value</td>
													<td class="ExcelHeaderCell" align="center">Form JJ Applicable</td>
													<td class="ExcelHeaderCell" align="center">Reason</td>
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
									<td colspan="4">
										<table border="0" cellspacing="0" cellpadding="0" width="572">
											<tr>
												<td class="FieldCell" Rowspan="3">Remarks</td>
												<td class="FieldCellSub" Rowspan="3" >
													<textarea rows="3" name="txtRemarks" cols="50" class="Formelem"></textarea>
												</td>
												<td class="FieldCell">Transport</td>
												<td class="FieldCellSub">
													<input type="text" name="txtTransport" maxlength=50 size="50" class="formelem">
												</td>
											</tr>
											<tr>
												<td class="FieldCell">Taken By</td>
												<td class="FieldCellSub">
													<input type="text" name="txtTakenBy" maxlength=50 size="50" class="formelem">
												</td>
											</tr>
											<tr>
												<td class="FieldCell">Delivery By</td>
												<td class="FieldCellSub">
													<input type="text" name="txtDeliveryBy" maxlength=50 size="50" class="formelem">
												</td>
											</tr>
										</table>
									</td>
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
													<p align="center">
													<input type="button" value="Save" name="BtnSubmit" class="ActionButton" onClick="CheckSubmit('<%=FormatDate(date)%>')">
 													<input type="reset" value="Reset" name="B1" class="ActionButton">
 													&nbsp;
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

