<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	stkMgmtSMEntry.asp
	'Module Name				:	Inventory (Stock Management Status Management)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	June 02, 2003
	'Modified By				:	Ragavendran R
	'Modified On				:	Dec 21,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	stkMgmtSMInsert.asp
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Stock Management - Status Management</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData">
<Output/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="OutData1">
<Output/>
</script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
dim j,OutDataValue
dim sorgID,iClass,sStore,iInvRec,sLot,sBin
j = 0

Function FnInit(sItemCode)
	'document.formname.selItem.value = sItemCode
	GetXML()
	DisplayDetails()
End Function

Function GetXML()
	clearXML

	'set obj = document.formname.selItem
	Dim	sTemp
	
	sorgID = trim(document.formname.hOrgID.value)
	iClass = trim(document.formname.hClass.value)
	iItem = Trim(document.formname.hItemCode.value)
	
	sTemp = sorgID & ":" & iClass & ":" & iItem
	
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","itmStatusXMLSelect.asp", false
	objhttp.send
	
	'alert(objhttp.responseText)
	if objhttp.responseXML.xml <> "" then
		OutData.loadXML objhttp.responseXML.xml
	else
		clearXML
	end if
	
End Function

Function CheckLot(obj,iValue)
	sTemp = document.formname.selItem(document.formname.selItem.selectedIndex).text & " -- " & trim(idClass.innerText) & "`" & iValue
	set OutDataValue = showModalDialog("stkMgmtSMPoP.asp?sTemp="&obj.value&"&sValue="&sTemp,OutData,"dialogHeight:310px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No")

	arrTemp = split(obj.value,"AAAA")
	sOrgID = arrTemp(1)
	iItem = arrTemp(2)
	iClass = arrTemp(3)
	sLot = arrTemp(4)

	if sLot = "0" then sLot = "NULL"

	sStore = arrTemp(6)
	sBin = arrTemp(7)
	iInvRec = arrTemp(8)

	Set RootO = OutData.documentElement
	For Each HeaderNode In RootO.childNodes
		if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then
			if HeaderNode.Attributes.Item(0).nodeValue = sStore and HeaderNode.Attributes.Item(1).nodeValue = sBin then
				For Each HNode In HeaderNode.childNodes
					if HNode.Attributes.Item(0).nodeValue = sLot and HNode.Attributes.Item(1).nodeValue = iInvRec then
						set Q = eval("document.formname."&Replace(obj.name,"h","txtRej"))
						Q.value = HNode.Attributes.Item(4).nodeValue

						set Q = eval("document.formname."&Replace(obj.name,"h","txtOnH"))
						Q.value = HNode.Attributes.Item(5).nodeValue

						set Q = eval("document.formname."&Replace(obj.name,"h","txtRes"))
						Q.value = HNode.Attributes.Item(6).nodeValue

						exit function
					end if
				next
			end if
		end if
	Next
End Function

Function CheckLotUn(obj,iValue)
	sTemp = document.formname.selItem(document.formname.selItem.selectedIndex).text & " -- " & trim(idClass.innerText) & "`" & iValue
	set OutDataValue = showModalDialog("stkMgmtSMUnPoP.asp?sTemp="&obj.value&"&sValue="&sTemp,OutData,"dialogHeight:310px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")

	arrTemp = split(obj.value,"AAAA")
	sOrgID = arrTemp(1)
	iItem = arrTemp(2)
	iClass = arrTemp(3)
	sLot = arrTemp(4)

	if sLot = "0" then sLot = "NULL"

	sStore = arrTemp(6)
	sBin = arrTemp(7)
	iInvRec = arrTemp(8)
	
	'alert(OutData.xml)
	iTot = 0 
	'exit function
	Set RootO = OutData.documentElement
	For Each HeaderNode In RootO.childNodes
		if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then
			if HeaderNode.Attributes.Item(0).nodeValue = sStore and HeaderNode.Attributes.Item(1).nodeValue = sBin then
				For Each HNode In HeaderNode.childNodes
					if HNode.Attributes.Item(0).nodeValue = sLot and HNode.Attributes.Item(1).nodeValue = iInvRec then
						For Each INode In HNode.childNodes
							if StrComp(Trim(INode.NodeName),"RESERVEDETAILS") = 0 then
								iTot = cdbl(iTot) + (cdbl(INode.Attributes.Item(2).nodeValue) - cdbl(INode.Attributes.Item(3).nodeValue))
							end if
						next
						set Q = eval("document.formname."&Replace(obj.name,"h","txtRes"))
						Q.value = iTot
					end if
				next
			end if
		end if
	Next
End Function

Function Check(obj)

	if trim(obj.value) = "" then obj.value = obj.defaultValue
	obj.defaultValue = obj.value
	sStr = obj.name
	iValue = obj.value

	arrTemp = split(sStr,"A")
	iCount = arrTemp(1)

	set objH = eval("document.formname."&Replace(obj.name,"txtStk","h"))
	if cdbl(iValue) <> 0 then CheckLot objH,iValue

End Function

Function SelPop()
	Dim iItem
	iItem = trim(document.formname.hItem.value)  
	sorgID = trim(document.formname.hOrgID.value)
	'alert sorgID
	iClass = trim(document.formname.hClass.value)
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	objhttp.Open "GET","itmMgmtXMLSelect.asp?orgID="&sorgID&"&iClass="&iClass&"&iItem="&iItem, false
	objhttp.send
	'alert objhttp.responseTEXT
	if objhttp.responseXML.xml <> "" then
		OutData1.loadXML objhttp.responseXML.xml
	end if
	
End Function

Function SelPopUp(obj)
	
	dim sItem,sClass,aobj,iSerialNo, sOrg, sLoc, sBin
	
	sLot = eval("document.formname.TxtLot"&obj).value
	'alert sLot
	sItem = document.formname.hItem.value 
	sClass = document.formname.hClass.value
	sOrg = document.formname.hOrgID.value   
	sLoc = document.formname.hLoc.value
	sBin = document.formname.hBin.value    
	
	sTempValues = sLot&"`"&sLot&"`"&sClass&"`"&sItem&"`"&sOrg&"`"&sLoc&"`"&sBin
	Set RootO = OutData1.documentElement
'	alert RootO.XML
'	Exit Function
	sExp ="//ITEM [ @CLACODE = "&sClass&" and @ITMCODE = "&sItem&"]/PickDet"
	Set ItemNode = RootO.Selectnodes(sExp)
'	alert ItemNode.Length
'	Exit Function	
	if ItemNode.Length > 0 then
		Set OutData.documentElement = ItemNode.Item(0)
	end if			
'	alert sTempValues
	
	set OutDataValue = showModalDialog("SerialWiseStockPoP.asp?sTemp="&sTempValues,OutData1,"dialogHeight:370px;dialogWidth:380px;center:Yes;help:No;resizable:No;status:No")

	Set Root = OutData1.documentElement
	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.Item(5).nodeValue = iCtr then
			if HeaderNode.HaschildNodes() then
			
				set Q = eval("document.formname.txtIss"&obj)
				Q.value = HeaderNode.Attributes.Item(3).nodeValue
			end if
			exit function
		end if
	Next
	'alert Root.XML
	sExp = "//PickDet [ @CLAS= "&sClass&" and @ITM = "&sItem&"]/PICK [ @LOC ="&sLoc&" and @BIN = "&sBin&" and @LOTNO = '"&cStr(sLot)&"']"   
	'alert sExp
	Set ItemNode = Root.Selectnodes(sExp)	
	If ItemNode.Length > 0 Then
		eval("document.formname.txtRejA"&obj).value  = ItemNode.Item(0).Attributes.getNamedItem("ISSQTY").Value	
	End If
	'alert OutData.XML
	'alert OutData1.XML
	
End Function

Function CheckUn(obj)

	if trim(obj.value) = "" then obj.value = obj.defaultValue
	obj.defaultValue = obj.value
	sStr = obj.name
	iValue = obj.value

	arrTemp = split(sStr,"A")
	iCount = arrTemp(1)

	set objH = eval("document.formname."&Replace(obj.name,"txtRes","h"))
	if cdbl(iValue) <> 0 then CheckLotUn objH,iValue

End Function

Function DisplayDetails(iSTNo)
Dim sRcptNumbering
	ClearTable
	j = 0

	iItem = trim(document.formname.selItem.value)

	arrTemp1 = split(iSTNo,":")
	if ubound(arrTemp1) > 0 then
		sLoc = arrTemp1(0)
		sBin = arrTemp1(1)
	else
		sLoc = arrTemp1(0)
		sBin = "0"
	end if

	Set Root = OutData.documentElement
	Set RootO = OutData.documentElement
	'alert RootO.XML
	For Each HeaderONode In RootO.childNodes
		if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
			For Each PickNode In HeaderONode.childNodes
				PickNode.setAttribute "QTYCLE", "0"
				PickNode.setAttribute "QTYREJ", "0"
				PickNode.setAttribute "QTYOHO", "0"
				PickNode.setAttribute "QTYRES", "0"
				PickNode.setAttribute "QTY", "0"
			next
		end if
		if StrComp(Trim(HeaderONode.NodeName),"UOM") = 0 then
			sCheck = HeaderONode.Attributes.Item(2).nodeValue
		end if
	next
	i = 1
	For Each HeaderNode In Root.childNodes
		if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then
			if HeaderNode.Attributes.Item(0).nodeValue = sLoc and HeaderNode.Attributes.Item(1).nodeValue = sBin then
				document.formname.hLoc.value = HeaderNode.Attributes.Item(0).nodeValue
				document.formname.hBin.value = HeaderNode.Attributes.Item(1).nodeValue
				sRcptNumbering = trim(HeaderNode.Attributes.Item(4).nodeValue)
				document.formname.hItem.value = trim(HeaderNode.Attributes.Item(5).nodeValue)
				'alert document.formname.hItem.value  
				For Each ItemNode In HeaderNode.childNodes
					j = j + 1
					set oRow = document.all.tblData.insertRow(j+1)

					set headerCell=oRow.insertCell()
					headerCell.innerHTML=j
					headerCell.className="ExcelSerial"
					headerCell.align="center"
					set headerCell=oRow.insertCell()
					if trim(ItemNode.Attributes.Item(0).nodeValue) = "NULL" then
						set oText = document.createElement("<input type=""text"" name=""txtLot"&CStr(j)&""" size=""30"" value=""-"" READONLY class=""FormelemRead"">" )
						headerCell.appendChild(oText)
					else
						set oText = document.createElement("<input type=""text"" name=""txtLot"&CStr(j)&""" size=""30"" value="""&ItemNode.Attributes.Item(0).nodeValue&""" READONLY class=""FormelemRead"">" )
						headerCell.appendChild(oText)
					end if
					headerCell.className="ExcelDisplayCell"
					headerCell.align="left"
					headerCell.width = "150"

					if not trim(ItemNode.Attributes.Item(0).nodeValue) = "NULL" then

						For Each StNode In ItemNode.childNodes
							iStockNo = trim(StNode.Attributes.Item(0).nodeValue)
							exit for
						next

						set headerCell=oRow.insertCell()
						'set oText = document.createElement("<input type=""text"" name=""txtStkA"&CStr(j)&""" size=""12"" maxlength=""10"" value="""&ItemNode.Attributes.Item(2).nodeValue&""" READONLY style=""text-align:right;cursor:hand;FONT-WEIGHT:bold"" class=""FormelemRead"" onClick=""Check(this)"">" )
						set oText = document.createElement("<input type=""text"" name=""txtStkA"&CStr(j)&""" size=""12"" maxlength=""10"" value="""&ItemNode.Attributes.Item(2).nodeValue&""" READONLY style=""text-align:right"" class=""FormelemRead"">" )
						headerCell.appendChild(oText)
						headerCell.className="ExcelDisplayCell"
						headerCell.width = "10"
						
						bHiddenFlag = false
						
						for iii = 0 to document.formname.elements.length - 1
							if document.formname.elements(iii).type = "hidden" then
								if (instr(document.formname.elements(iii).name,"hA"&j)) > 0 then
									bHiddenFlag = true
								end If
							end if	
						next

						if bHiddenFlag then
							set objHidden = eval("document.formname.hA"&j)
							objHidden.value = "btnAAAA"&sorgID&"AAAA"&iItem&"AAAA"&iClass&"AAAA"&ItemNode.Attributes.Item(0).nodeValue&"AAAA"&ItemNode.Attributes.Item(2).nodeValue&"AAAA"&sLoc&"AAAA"&sBin&"AAAA"&ItemNode.Attributes.Item(1).nodeValue&"AAAA"&iStockNo&"AAAA"&SCheck
						else
							set oText1 = document.createElement("<input type=""hidden"" name=""hA"&j&""" value=""btnAAAA"&sorgID&"AAAA"&iItem&"AAAA"&iClass&"AAAA"&ItemNode.Attributes.Item(0).nodeValue&"AAAA"&ItemNode.Attributes.Item(2).nodeValue&"AAAA"&sLoc&"AAAA"&sBin&"AAAA"&ItemNode.Attributes.Item(1).nodeValue&"AAAA"&iStockNo&"AAAA"&SCheck&""">" )
							document.formname.appendChild(oText1)
						end if
						
	
						For Each StNode In ItemNode.childNodes
							if StNode.Attributes.Item(1).nodeValue = ItemNode.Attributes.Item(0).nodeValue and StNode.Attributes.Item(2).nodeValue = ItemNode.Attributes.Item(1).nodeValue then
								set headerCell=oRow.insertCell()
								set oText = document.createElement("<a href=""#"">" )
								
								If document.formname.hRcptNumbering.value = "S" or document.formname.hRcptNumbering.value = "LS" Then 
									'set oText1 = document.createElement("<img name=""btnA"" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" onClick=""SelPop()"" width=""15"" height=""15"">")
									set oText1 = document.createElement("<img name=""btnA"&j&""" value="""&j&""" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" onClick=""selPopUp(this.value)"" width=""15"" height=""15"">")
									headerCell.appendChild(oText1)
								End If
								headerCell.appendChild(oText)								
								'set oText = document.createElement("<input type=""text"" name=""txtRejA"&j&""" size=""12"" READONLY onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&StNode.Attributes.Item(6).nodeValue&""" class=""Formelem"" style=""text-align=right"">" )
								set oText = document.createElement("<input type=""text"" name=""txtRejA"&j&""" size=""12"" READONLY onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value=""0"" class=""Formelem"" style=""text-align=right"">" )
								headerCell.appendChild(oText)
'								headerCell.className="ExcelInputCell"
'								headerCell.width = "10"

								'set headerCell=oRow.insertCell()									

								headerCell.className="ExcelDisplayCell"
								headerCell.align="center"
								headerCell.width = "15"


								set headerCell=oRow.insertCell()
								set oText = document.createElement("<input type=""text"" name=""txtOnHA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&StNode.Attributes.Item(5).nodeValue&""" class=""Formelem"" style=""text-align=right"">" )
								headerCell.appendChild(oText)
								headerCell.className="ExcelInputCell"
								headerCell.width = "10"

								set headerCell=oRow.insertCell()
								if trim(StNode.Attributes.Item(4).nodeValue) = "0" then
									set oText = document.createElement("<input type=""text"" name=""txtResA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&StNode.Attributes.Item(4).nodeValue&""" READONLY class=""Formelem"" style=""text-align=right"">" )
								else
									set oText = document.createElement("<input type=""text"" name=""txtResA"&j&""" size=""12"" value="""&StNode.Attributes.Item(4).nodeValue&""" READONLY style=""text-align:right;cursor:hand;FONT-WEIGHT:bold"" class=""FormelemRead"" onClick=""CheckUn(this)"">" )
									'set oText = document.createElement("<input type=""text"" name=""txtStkA"&CStr(j)&""" size=""12"" maxlength=""10"" value="""&ItemNode.Attributes.Item(2).nodeValue&""" READONLY style=""text-align:right;cursor:hand;FONT-WEIGHT:bold"" class=""FormelemRead"" onClick=""Check(this)"">" )
								end if
								headerCell.appendChild(oText)
								headerCell.className="ExcelDisplayCell"
								headerCell.width = "10"
								
							end if
							exit for
						next
					else
						set headerCell=oRow.insertCell()
						set oText = document.createElement("<input type=""text"" name=""txtStkA"&CStr(j)&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&ItemNode.Attributes.Item(2).nodeValue&""" style=""text-align:right"" class=""FormelemRead"">" )
						headerCell.appendChild(oText)
						headerCell.className="ExcelDisplayCell"
						headerCell.width = "10"
						headerCell.align = "right"

						For Each StNode In ItemNode.childNodes
							if StNode.Attributes.Item(1).nodeValue = ItemNode.Attributes.Item(0).nodeValue and StNode.Attributes.Item(2).nodeValue = ItemNode.Attributes.Item(1).nodeValue then

								set headerCell=oRow.insertCell()
								set oText = document.createElement("<a href=""#"">" )
				
								If document.formname.hRcptNumbering.value = "S" or document.formname.hRcptNumbering.value = "LS" Then    		
									'set oText1 = document.createElement("<img name=""btnA"" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" onClick=""SelPop()"" width=""15"" height=""15"">")
									set oText1 = document.createElement("<img name=""btnA"&j&""" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" value="""&j&""" onClick=""SelPopup(this.value)"" width=""15"" height=""15"">")
									headerCell.appendChild(oText1)
								End If
								headerCell.appendChild(oText)								
								'set oText = document.createElement("<input type=""text"" name=""txtRejA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&StNode.Attributes.Item(6).nodeValue&""" class=""Formelem"" style=""text-align=right"">" )
								set oText = document.createElement("<input type=""text"" name=""txtRejA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value=""0"" class=""Formelem"" style=""text-align=right"">" )
								headerCell.appendChild(oText)
'								headerCell.className="ExcelInputCell"
'								headerCell.width = "10"

								'set headerCell=oRow.insertCell()									

								headerCell.className="ExcelDisplayCell"
								headerCell.align="center"
								headerCell.width = "15"


								set headerCell=oRow.insertCell()
								set oText = document.createElement("<input type=""text"" name=""txtOnHA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&StNode.Attributes.Item(5).nodeValue&""" class=""Formelem"" style=""text-align=right"">" )
								headerCell.appendChild(oText)
								headerCell.className="ExcelInputCell"
								headerCell.width = "10"

								set headerCell=oRow.insertCell()
								if trim(StNode.Attributes.Item(4).nodeValue) = "0" then
									set oText = document.createElement("<input type=""text"" name=""txtResA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&StNode.Attributes.Item(4).nodeValue&""" READONLY class=""Formelem"" style=""text-align=right"">" )
								else
									set oText = document.createElement("<input type=""text"" name=""txtResA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&StNode.Attributes.Item(4).nodeValue&""" class=""Formelem"" style=""text-align=right"">" )
								end if
								headerCell.appendChild(oText)
								headerCell.className="ExcelInputCell"
								headerCell.width = "10"
							end if
							exit for
						next
					end if
				next
			end if
		end if
	next
	
	SelPop() 
	
end Function

'***************
Function DisplayDetails()
	Dim sRcptNumbering
	ClearTable
	j = 0
	
	blnPrintUOM = "N"
	
	Set Root = OutData.documentElement
	Set RootO = OutData.documentElement
	'alert RootO.XML
	alert Root.XML
	For Each ndItem In RootO.childNodes
		For each HeaderONode in ndItem.childNodes
			if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
				sLoc = HeaderONode.getAttribute("LOC")
				sBin = HeaderONode.getAttribute("BIN")
				For Each PickNode In HeaderONode.childNodes
					PickNode.setAttribute "QTYCLE", "0"
					PickNode.setAttribute "QTYREJ", "0"
					PickNode.setAttribute "QTYOHO", "0"
					PickNode.setAttribute "QTYRES", "0"
					PickNode.setAttribute "QTY", "0"
				next
			end if
			if StrComp(Trim(HeaderONode.NodeName),"UOM") = 0 then
				sCheck = HeaderONode.Attributes.Item(2).nodeValue
			end if
		Next'For each HeaderONode in ndItem.childNodes
	next
	
	i = 1
	
	For Each ndItem In Root.childNodes
		sItemName = nditem.getAttribute("IName")
		For each HeaderNode in ndItem.childNodes
			if StrComp(Trim(HeaderNode.NodeName),"LOCDET") = 0 then
				
					document.formname.hLoc.value = HeaderNode.Attributes.Item(0).nodeValue
					document.formname.hBin.value = HeaderNode.Attributes.Item(1).nodeValue
					sLocName = HeaderNode.getAttribute("LOCNAME")
					sRcptNumbering = trim(HeaderNode.Attributes.Item(4).nodeValue)
					document.formname.hItem.value = trim(HeaderNode.Attributes.Item(5).nodeValue)
					'alert document.formname.hItem.value  
					For Each ItemNode In HeaderNode.childNodes
						j = j + 1
						set oRow = document.all.tblData.insertRow(j+1)

						set headerCell=oRow.insertCell()
						headerCell.innerHTML=j
						headerCell.className="ExcelSerial"
						headerCell.align="center"
						'alert(sItemName)
						set headerCell=oRow.insertCell()
							headerCell.innerText = sItemName
							headerCell.className="ExcelDisplayCell"
							
						set headerCell=oRow.insertCell()
							headerCell.innerText = sLocName
							headerCell.className="ExcelDisplayCell"
						
						set headerCell=oRow.insertCell()
						if trim(ItemNode.Attributes.Item(0).nodeValue) = "NULL" then
							set oText = document.createElement("<input type=""text"" name=""txtLot"&CStr(j)&""" size=""30"" value=""-"" READONLY class=""FormelemRead"">" )
							headerCell.appendChild(oText)
						else
							set oText = document.createElement("<input type=""text"" name=""txtLot"&CStr(j)&""" size=""30"" value="""&ItemNode.Attributes.Item(0).nodeValue&""" READONLY class=""FormelemRead"">" )
							headerCell.appendChild(oText)
						end if
						headerCell.className="ExcelDisplayCell"
						headerCell.align="left"
						headerCell.width = "150"

						if not trim(ItemNode.Attributes.Item(0).nodeValue) = "NULL" then
						
							blnPrintUOM = "Y"						
	
							For Each StNode In ItemNode.childNodes
								iStockNo = trim(StNode.Attributes.Item(0).nodeValue)
								exit for
							next

							set headerCell=oRow.insertCell()
							'set oText = document.createElement("<input type=""text"" name=""txtStkA"&CStr(j)&""" size=""12"" maxlength=""10"" value="""&ItemNode.Attributes.Item(2).nodeValue&""" READONLY style=""text-align:right;cursor:hand;FONT-WEIGHT:bold"" class=""FormelemRead"" onClick=""Check(this)"">" )
							set oText = document.createElement("<input type=""text"" name=""txtStkA"&CStr(j)&""" size=""12"" maxlength=""10"" value="""&ItemNode.Attributes.Item(2).nodeValue&""" READONLY style=""text-align:right"" class=""FormelemRead"">" )
							headerCell.appendChild(oText)
							headerCell.className="ExcelDisplayCell"
							headerCell.width = "10"
							
							bHiddenFlag = false
							
							for iii = 0 to document.formname.elements.length - 1
								if document.formname.elements(iii).type = "hidden" then
									if (instr(document.formname.elements(iii).name,"hA"&j)) > 0 then
										bHiddenFlag = true
									end If
								end if	
							next

							if bHiddenFlag then
								set objHidden = eval("document.formname.hA"&j)
								objHidden.value = "btnAAAA"&sorgID&"AAAA"&iItem&"AAAA"&iClass&"AAAA"&ItemNode.Attributes.Item(0).nodeValue&"AAAA"&ItemNode.Attributes.Item(2).nodeValue&"AAAA"&sLoc&"AAAA"&sBin&"AAAA"&ItemNode.Attributes.Item(1).nodeValue&"AAAA"&iStockNo&"AAAA"&SCheck
							else
								set oText1 = document.createElement("<input type=""hidden"" name=""hA"&j&""" value=""btnAAAA"&sorgID&"AAAA"&iItem&"AAAA"&iClass&"AAAA"&ItemNode.Attributes.Item(0).nodeValue&"AAAA"&ItemNode.Attributes.Item(2).nodeValue&"AAAA"&sLoc&"AAAA"&sBin&"AAAA"&ItemNode.Attributes.Item(1).nodeValue&"AAAA"&iStockNo&"AAAA"&SCheck&""">" )
								document.formname.appendChild(oText1)
							end if
							
		
							For Each StNode In ItemNode.childNodes
								if StNode.Attributes.Item(1).nodeValue = ItemNode.Attributes.Item(0).nodeValue and StNode.Attributes.Item(2).nodeValue = ItemNode.Attributes.Item(1).nodeValue then
									set headerCell=oRow.insertCell()
									set oText = document.createElement("<a href=""#"">" )
									
									If document.formname.hRcptNumbering.value = "S" or document.formname.hRcptNumbering.value = "LS" Then 
										'set oText1 = document.createElement("<img name=""btnA"" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" onClick=""SelPop()"" width=""15"" height=""15"">")
										set oText1 = document.createElement("<img name=""btnA"&j&""" value="""&j&""" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" onClick=""selPopUp(this.value)"" width=""15"" height=""15"">")
										headerCell.appendChild(oText1)
									End If
									headerCell.appendChild(oText)								
									'set oText = document.createElement("<input type=""text"" name=""txtRejA"&j&""" size=""12"" READONLY onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&StNode.Attributes.Item(6).nodeValue&""" class=""Formelem"" style=""text-align=right"">" )
									set oText = document.createElement("<input type=""text"" name=""txtRejA"&j&""" size=""12"" READONLY onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value=""0"" class=""Formelem"" style=""text-align=right"">" )
									headerCell.appendChild(oText)
	'								headerCell.className="ExcelInputCell"
	'								headerCell.width = "10"

									'set headerCell=oRow.insertCell()									

									headerCell.className="ExcelDisplayCell"
									headerCell.align="center"
									headerCell.width = "15"


									set headerCell=oRow.insertCell()
									set oText = document.createElement("<input type=""text"" name=""txtOnHA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&StNode.Attributes.Item(5).nodeValue&""" class=""Formelem"" style=""text-align=right"">" )
									headerCell.appendChild(oText)
									headerCell.className="ExcelInputCell"
									headerCell.width = "10"

									set headerCell=oRow.insertCell()
									if trim(StNode.Attributes.Item(4).nodeValue) = "0" then
										set oText = document.createElement("<input type=""text"" name=""txtResA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&StNode.Attributes.Item(4).nodeValue&""" READONLY class=""Formelem"" style=""text-align=right"">" )
									else
										set oText = document.createElement("<input type=""text"" name=""txtResA"&j&""" size=""12"" value="""&StNode.Attributes.Item(4).nodeValue&""" READONLY style=""text-align:right;cursor:hand;FONT-WEIGHT:bold"" class=""FormelemRead"" onClick=""CheckUn(this)"">" )
										'set oText = document.createElement("<input type=""text"" name=""txtStkA"&CStr(j)&""" size=""12"" maxlength=""10"" value="""&ItemNode.Attributes.Item(2).nodeValue&""" READONLY style=""text-align:right;cursor:hand;FONT-WEIGHT:bold"" class=""FormelemRead"" onClick=""Check(this)"">" )
									end if
									headerCell.appendChild(oText)
									headerCell.className="ExcelDisplayCell"
									headerCell.width = "10"
									
								end if
								exit for
							next
						else
							blnPrintUOM = "Y"						
							set headerCell=oRow.insertCell()
							set oText = document.createElement("<input type=""text"" name=""txtStkA"&CStr(j)&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&ItemNode.Attributes.Item(2).nodeValue&""" style=""text-align:right"" class=""FormelemRead"">" )
							headerCell.appendChild(oText)
							headerCell.className="ExcelDisplayCell"
							headerCell.width = "10"
							headerCell.align = "right"

							For Each StNode In ItemNode.childNodes
								if StNode.Attributes.Item(1).nodeValue = ItemNode.Attributes.Item(0).nodeValue and StNode.Attributes.Item(2).nodeValue = ItemNode.Attributes.Item(1).nodeValue then

									set headerCell=oRow.insertCell()
									set oText = document.createElement("<a href=""#"">" )
					
									If document.formname.hRcptNumbering.value = "S" or document.formname.hRcptNumbering.value = "LS" Then    		
										'set oText1 = document.createElement("<img name=""btnA"" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" onClick=""SelPop()"" width=""15"" height=""15"">")
										set oText1 = document.createElement("<img name=""btnA"&j&""" border=""0"" src=""../../assets/images/iTMS%20Icons/Entry.gif"" value="""&j&""" onClick=""SelPopup(this.value)"" width=""15"" height=""15"">")
										headerCell.appendChild(oText1)
									End If
									headerCell.appendChild(oText)								
									'set oText = document.createElement("<input type=""text"" name=""txtRejA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&StNode.Attributes.Item(6).nodeValue&""" class=""Formelem"" style=""text-align=right"">" )
									set oText = document.createElement("<input type=""text"" name=""txtRejA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value=""0"" class=""Formelem"" style=""text-align=right"">" )
									headerCell.appendChild(oText)
	'								headerCell.className="ExcelInputCell"
	'								headerCell.width = "10"

									'set headerCell=oRow.insertCell()									

									headerCell.className="ExcelDisplayCell"
									headerCell.align="center"
									headerCell.width = "15"


									set headerCell=oRow.insertCell()
									set oText = document.createElement("<input type=""text"" name=""txtOnHA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&StNode.Attributes.Item(5).nodeValue&""" class=""Formelem"" style=""text-align=right"">" )
									headerCell.appendChild(oText)
									headerCell.className="ExcelInputCell"
									headerCell.width = "10"

									set headerCell=oRow.insertCell()
									if trim(StNode.Attributes.Item(4).nodeValue) = "0" then
										set oText = document.createElement("<input type=""text"" name=""txtResA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&StNode.Attributes.Item(4).nodeValue&""" READONLY class=""Formelem"" style=""text-align=right"">" )
									else
										set oText = document.createElement("<input type=""text"" name=""txtResA"&j&""" size=""12"" onkeypress=""DoKeyPress('"&sCheck&"',7,3)"" value="""&StNode.Attributes.Item(4).nodeValue&""" class=""Formelem"" style=""text-align=right"">" )
									end if
									headerCell.appendChild(oText)
									headerCell.className="ExcelInputCell"
									headerCell.width = "10"
								end if
								exit for
							next
						end if
					next
			elseif StrComp(Trim(HeaderNode.NodeName),"UOM") = 0 and blnPrintUOM = "Y" then
				set headerCell=oRow.insertCell()
				headerCell.innerText = HeaderNode.getAttribute("UoMName")
				headerCell.ClassName="ExcelDisplayCell"
			end if
		next
	next
	
	SelPop() 
	
end Function
'****************


Function popClaDisplay()
	'document.formname.selStore.options.length = 1
	Set Root = OutData.documentElement
	For Each ndChild In Root.childNodes
		For each HeaderNode in ndChild.childNodes
			'alert(HeaderNode.NodeName)
			if Trim(HeaderNode.NodeName) = "LOCDET" then
				'if HeaderNode.hasChildNodes() then
				'alert HeaderNode.Attributes.Item(4).nodeValue
				document.formname.hRcptNumbering.value = HeaderNode.Attributes.Item(4).nodeValue  
					if not HeaderNode.Attributes.Item(1).nodeValue = "0" then
						document.formname.selStore.length = document.formname.selStore.length+1
						document.formname.selStore.options(document.formname.selStore.length-1).text = trim(HeaderNode.Attributes.Item(2).nodeValue)&" -- "&trim(HeaderNode.Attributes.Item(3).nodeValue)
						document.formname.selStore.options(document.formname.selStore.length-1).Value = trim(HeaderNode.Attributes.Item(0).nodeValue)&":"&trim(HeaderNode.Attributes.Item(1).nodeValue)
					else
						document.formname.selStore.length = document.formname.selStore.length+1
						document.formname.selStore.options(document.formname.selStore.length-1).text = trim(HeaderNode.Attributes.Item(2).nodeValue)
						document.formname.selStore.options(document.formname.selStore.length-1).Value = trim(HeaderNode.Attributes.Item(0).nodeValue)
					end if
				'end if
			end if
			if StrComp(Trim(HeaderNode.NodeName),"UOM") = 0 then
				idUoM.innerHTML = trim(HeaderNode.Attributes.Item(1).nodeValue) & "&nbsp;"
			end if
		next
	next
end Function

Function clearXML()
	Set Root = OutData.documentElement
	if Root.hasChildNodes() then
		For Each HeaderNode In Root.childNodes
			set a=Root.removeChild(HeaderNode)
		next
	end if
'	document.formname.selStore.options.length = 1
end Function

Function ClearTable()
	dim i
	for i=2 to document.all.tblData.rows.length - 1
		document.all.tblData.deleteRow(2)
	next
end Function

Function CheckSubmit()
	dim objQ,iQtyTot,iTempQty
	Set Root = OutData.documentElement
	if j = 0 then exit function
	

'	iSTNo = trim(document.formname.selStore.value)
'
'	arrTemp1 = split(iSTNo,":")
'	if ubound(arrTemp1) > 0 then
'		sLoc = arrTemp1(0)
'		sBin = arrTemp1(1)
'	else
'		sLoc = arrTemp1(0)
'		sBin = "0"
'	end if

	i = 0
	iTempQty = 0
	iTotQty = 0
	
	For Each ndItem In Root.childNodes
		For Each HeaderONode in ndItem.childNodes
		if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
		'	if (HeaderONode.Attributes.Item(0).nodeValue = sLoc and HeaderONode.Attributes.Item(1).nodeValue = sBin) then
				For Each PickNode In HeaderONode.childNodes
					i = i + 1
					'if trim(PickNode.Attributes.Item(0).nodeValue) = "NULL" then
						
						iTotQty = cdbl(iTotQty) + cdbl(PickNode.Attributes.Item(7).nodeValue)
						
						set obj = eval("document.formname.txtStkA"&cstr(i))

						set objR = eval("document.formname.txtRejA"&i)
						set objH = eval("document.formname.txtOnHA"&i)
						set objE = eval("document.formname.txtResA"&i)

						if trim(objR.value) = "" then
							alert("Enter Rejected Quantity")
							objR.select()
							exit function
						elseif trim(objH.value) = "" then
							alert("Enter On Hold Quantity")
							objH.select()
							exit function
						elseif trim(objE.value) = "" then
							alert("Enter Reserved Quantity")
							objE.select()
							exit function
						else

							iCQty = 0
							iCQty = cdbl(objR.value) + cdbl(objH.value) + cdbl(objE.value)
							iPQty = 0

							For Each StNode In PickNode.childNodes
								For Each INode In StNode.childNodes
									if StrComp(Trim(INode.NodeName),"RESERVEDETAILS") = 0 then
										iPQty = cdbl(iPQty) + (cdbl(INode.Attributes.Item(2).nodeValue) - cdbl(INode.Attributes.Item(3).nodeValue))
									end if
								next
								exit for
							next

							iTempQty = cdbl(iCQty) - cdbl(iPQty)

							if (cdbl(iCQty)) > 0 then
								if (cdbl(iCQty) > cdbl(obj.value)) then
									alert("Total Quantity should be equal to or less than Stock Quantity (" & obj.value & ")")
									exit function
								end if
							end if

						end if
					'end if
				next
			'end if
		end if
		Next
	next
	
'	if cdbl(iTempQty) = 0 and cdbl(iTotQty) = 0 then
'		alert("No Status Quantity has been changed / entered")
'		exit function
'	end if

	i = 0

	Set RootO = OutData.documentElement
	For Each ndItem In RootO.childNodes
		For Each HeaderONode in ndItem.childNodes
			if StrComp(Trim(HeaderONode.NodeName),"LOCDET") = 0 then
				'if (HeaderONode.Attributes.Item(0).nodeValue = sLoc and HeaderONode.Attributes.Item(1).nodeValue = sBin) then
					For Each PickNode In HeaderONode.childNodes
						i = i + 1
						'if trim(PickNode.Attributes.Item(0).nodeValue) = "NULL" then
							set objR = eval("document.formname.txtRejA"&i)
							set objH = eval("document.formname.txtOnHA"&i)
							set objE = eval("document.formname.txtResA"&i)

							PickNode.setAttribute "QTYCLE", cdbl(iTempQty)
							PickNode.setAttribute "QTYREJ", cdbl(objR.value)
							PickNode.setAttribute "QTYOHO", cdbl(objH.value)
							PickNode.setAttribute "QTYRES", cdbl(objE.value)

						'end if
					next
			'	else
			'		set a = RootO.removeChild(HeaderONode)
			'	end if
			end if
		Next
	next
'	ALERT Root.XML
'	exit Function

'	sorgID = trim(document.formname.hOrgID.value)
'	iClass = trim(document.formname.hClass.value)

'	RootO.setAttribute "ITEM", trim(document.formname.selItem.value)
'	RootO.setAttribute "CLASS", iClass
'	RootO.setAttribute "ORG", sorgID

	Set objhttp = CreateObject("Microsoft.XMLHTTP")

'	alert OutData1.xml
'	alert(OutData.xml)
	'exit function

	Set objhttp1 = CreateObject("Microsoft.XMLHTTP")
	objhttp1.Open "POST","SerialWiseStockXML.asp", false
	objhttp1.send OutData1.XMLDocument

	objhttp.Open "POST","stkMgmtSMInsert.asp", false
	objhttp.send OutData.XMLDocument

	'alert(objhttp.responseText)
	'exit function

	if objhttp.responseText = "" then
		Msgbox ("Status Management has done")
		'if document.formname.hCallFrom.value = "ItemList" then
			window.location.href = "../master/ITEMLISTENTRY.ASP"
		'Else
		'	window.location.href = "stkMgmtSMEntry.asp"
		'end if 	
	else
		alert(objhttp.responseText)
		window.location.href = "stkMgmtSMEntry.asp"
	end if

end Function

</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
</head>

<%
	dim iCtr,arrTemp,sTemp,arrValue,sOrgID,iClass,arrTempName,sTempName
	dim sOrgName,sClassName,rsTemp
	
	set rsTemp = server.CreateObject("ADODB.Recordset")
	
	'sOrgName = trim(Request.Form("hOrgName"))
	sClassName = trim(Request.Form("hClassName"))
	'sOrgID = trim(Request.Form("selUnit"))
	sOrgID = Session("organizationcode")
	iClass = trim(Request.Form("selClass"))
	sTemp  = trim(Request.Form("hSelectedValue"))
	sTempName = trim(Request.Form("hItemNames"))
	
	'Response.Write "<p>sTemp="&sTemp
	
	if sTempName  = "" then
		with rsTemp
			.ActiveConnection=con
			.CursorLocation=3
			.CursorType=3
			.Source= "Select ItemDescription from Inv_M_ItemMaster where ItemCode = " & mid(sTemp,1,len(sTemp)-1) 
			.Open
		end with
		
		if not rsTemp.EOF then
			sTempName = trim(rsTemp(0)) & "|"
		end if 
		rsTemp.Close 	
	end if 'if sTempName  = "" then
	
	if trim(sClassName) = "" then
		with rsTemp
			.ActiveConnection=con
			.CursorLocation=3
			.CursorType=3
			.Source= "Select GroupName from Inv_M_Classification where GroupCode = " & iClass
			.Open
		end with
		
		if not rsTemp.EOF then
			sClassName = trim(rsTemp(0))
		end if 
		rsTemp.Close
	end if 'if trim(sClassName) = "" then
	
	arrTempName = split(mid(sTempName,1,len(sTempName)-1),"|")
	arrTemp = split(mid(sTemp,1,len(sTemp)-1),"|")
	
%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="FnInit('<%= mid(sTemp,1,len(sTemp)-1)%>')">
<form method="POST" name="formname" action="">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hClass" value="<%=iClass%>">
<input type=hidden name="hItemCode" value="<%= mid(sTemp,1,len(sTemp)-1)%>">
<input type=hidden name="hItem" value="">
<input type=hidden name="hRcptNumbering" value="">
<input type=hidden name="hLoc" value="">
<input type=hidden name="hBin" value="">

<input type="hidden" name="hCallFrom" value="<%=Request.Form("hCallFrom")%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Status Management
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack"></td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<td height="20" valign="bottom">
						<!--<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td class="TabCell" valign="bottom" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable"  onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td width="100%" align="center">Header</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="70">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td width="100%" align="center">Control</td>
										</tr>
									</table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    <font face="Verdana" size="1" color="#FFFFFF">&nbsp;</font>
								</td>
							</tr>
						</table>-->
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
								<td align="center"></td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <!--<td class="FieldCell">Organization</td>
                                            <td class="FieldCellSub">
	                                            <span class="DataOnly"><%=sOrgName%>&nbsp;</span>
                                            </td>
                                            <td class="FieldCellSub"></td>-->
                                            <!--<td class="FieldCell">Classification</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly" id="idClass"><%=sClassName%>&nbsp;</span>
                                            </td>-->
                                        </tr>
                                        <!--<tr>
                                            <td class="FieldCell">Item Name</td>
                                            <td class="FieldCellSub" colspan=4>
												<select size="1" name="selItem" class="FormElem" onChange="GetXML()">
													<option value="select">Select</option>
												<%
													for iCtr = 0 to UBound(arrTempName)
												%>
													<option value="<%=arrTemp(iCtr)%>"><%=arrTempName(iCtr)%></option>
												<%
													next
												%>
												</select>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">UoM</td>
                                            <td class="FieldCellSub">
	                                            <span class="DataOnly" id="idUoM"></span>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Store</td>
                                            <td class="FieldCellSub">
												<select size="1" name="selStore" class="FormElem" onChange="DisplayDetails(this.value)">
													<option value="select">Select</option>
												</select>
                                            </td>
                                            <td class="FieldCellSub"></td>
                                            <td class="FieldCell"></td>
                                            <td class="FieldCellSub"></td>
                                        </tr>-->
                                    </table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
							<tr>
								<td align="center"></td>
								<td valign="top" width="100%">
                                    <div class="frmBody" id="frm2" style="width: 575; height:300;">
                                        <table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center" width="10" rowspan="2">S.No.</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan="2">Item Name</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan="2">Store</td>
                                                <td class="ExcelHeaderCell" align="center" colspan="2">Existing Stock Information</td>
                                                <td class="ExcelHeaderCell" align="center" colspan="3">Status Management</td>
                                                <td class="ExcelHeaderCell" align="center" rowspan="2">UoM</td>
                                            </tr>
                                            <tr>
                                                <td class="ExcelHeaderCell" align="center">Lot Number</td>
                                                <td class="ExcelHeaderCell" align="center">Quantity</td>
                                                <td class="ExcelHeaderCell" align="center">Rejected</td>
                                                <td class="ExcelHeaderCell" align="center">On Hold</td>
                                                <td class="ExcelHeaderCell" align="center">Reserved</td>
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
												<p align="center">
                                                <input type="button" value="Save" name="B3" class="ActionButton" onClick="CheckSubmit()">
                                                <input type="reset" value="Reset" name="B1" class="ActionButton">
                                                <input type="button" value="Cancel" name="B2" class="ActionButton" onClick="Cancel('stkMgmtEntry.asp')">
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
