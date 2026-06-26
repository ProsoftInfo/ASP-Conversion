<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	BankInsDetails.asp
	'Module Name				:	Fixed Deposit(Transaction)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Jun 15,2005
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
<!--#include File="../../include/DatabaseConnection.asp" -->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->
<!--#include File="../../include/IncludeDatePicker.asp" -->
<%
	Dim oDOM,Root,node,sTemp,sTemparr,sDHname,sJDname,sNDname,sGDname,iAmt,sVouTy,sVouDate,sBookNo,sOrgId,Elem
	Dim sQry,rs,sPrintCheq,dtIssue,sDrawOn,sPayAt,iEntNo,sFlag,sVouType,iTransNo,iUsInsNo,sVouName
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")


		sTemp = Trim(Request("sTemp"))
		'Response.Write "sTemp="&sTemp
		
		If 1 = 2 Then
		sVouTy = Right(sTemp,1)
		sTemparr = Split(sTemp,":")
		sVouDate = sTemparr(2)
		sBookNo  = sTemparr(3)
		sOrgId	 = sTemparr(4)
		sVouType = sTemparr(5)
		iTransNo = sTemparr(6)
		'sTemparr = Split(sTemp,":")
		'sDHname = sTemparr(0)
		'iAmt = sTemparr(1)

		sVouName = sTemparr(7)
		End IF
		
		
		sTemparr = Split(sTemp,":")
		sVouTy   = Right(sTemparr(0),1)
		sVouDate = sTemparr(1)
		sBookNo  = sTemparr(2)
		sOrgId	 = sTemparr(3)
		sVouType = sTemparr(0)
		iTransNo = sTemparr(4)
		sVouName = sTemparr(5)
		
		iUsInsNo = ""

		sFlag = False

	If trim(iTransNo) <> "0" then
		'oDOM.Load server.MapPath("../temp/transaction/Voucher AMD_"&sVouName&"_"&Session.SessionID&".xml")
		oDOM.Load server.MapPath("../temp/transaction/"&iTransNo&".xml")

		set Root=oDOM.documentElement
		IF Root.haschildnodes then
			For each node in Root.childnodes
				IF trim(node.NodeName) = "BankInstrumentDet" then
					Root.Removechild node
				End If
			Next
		End IF
	End IF ' If trim(iTransNo) <> "0" then
'	Response.Write "iTransNo="&sVouType

	IF trim(sVouType) = "P" then
		sQry = "Select PrintCheques from Acc_M_BankDetails where OUDefinitionID = '"&sOrgId&"' and BookNumber ="&sBookNo
		'Response.Write sQry
		rs.Open sQry,con
		IF not rs.EOF then
			sPrintCheq = rs(0)
		End If
		rs.Close
		' Response.Write "sPrintCheq="&sPrintCheq
		IF trim(sPrintCheq) = "1" then
			sQry = "Select convert(varchar,dateOfIssue,103),DrawnOn,PayableAt,EntryNo from Acc_R_BankInstrumentDetails where OUDefinitionID = '"&sOrgId&"' and BookNumber = "&sBookNo
			'Response.Write sQry
			rs.Open sQry,con
			IF not rs.EOF then
				sFlag = True
				dtIssue = rs(0)
				sDrawOn = rs(1)
				sPayAt  = rs(2)
				iEntNo  = rs(3)
			End If
			rs.Close
			'sQry = "Select EntryNo,InstrumentEntryNo,InstrumentNo from Acc_R_BankInstrumentUsage where CreatedTransNo = "&iTransNo&" and Status = 'U'"

			sQry = "Select I.InstrumentEntryNo,I.BankInstrumentEntryNo,U.EntryNo,U.InstrumentEntryNo,U.InstrumentNo,I.BankInstrumentType,convert(VarChar,I.BankInstrumentDate,103),"&_
				   "I.PayableAt,I.DrawnOnBank,I.InstrumentAmount from Acc_R_BankInstrumentUsage as U,Acc_T_CreatedVoucherInstrumentDet as I where U.CreatedTransNo = "&iTransNo&" and "&_
				   "I.CreatedTransNo = U.CreatedTransNo and I.BankInstrumentEntryNo = U.EntryNo and Isnull(I.InstrumentEntryNo1,0) = U.InstrumentEntryNo and U.Status = 'U'"

			 'Response.Write sQry

			 'Response.End
			rs.Open sQry,con
			IF not rs.EOF then

				Do while not rs.EOF
					Set Elem = oDOM.CreateElement("BankInstrumentDet")
					Elem.SetAttribute "SlNo",rs(0)

					IF trim(rs(1)) = "0" then
						Elem.SetAttribute "InsNo",rs(4)
					Else
						iUsInsNo = rs(2)&"-"&rs(3)&"-"&rs(4)
						Elem.SetAttribute "InsNo",iUsInsNo
					End IF
					Elem.SetAttribute "InsType",rs(5)
					Elem.SetAttribute "InsDate",rs(6)
					Elem.SetAttribute "PayAt",rs(7)
					Elem.SetAttribute "DrawnOn",rs(8)
					Elem.SetAttribute "InsAmt",rs(9)
					Elem.setAttribute "Option","Y"
					Elem.setAttribute "Action","0"
					Root.appendchild Elem
					rs.MoveNext
				loop
			End If
			rs.Close
			' Response.Write iUsInsNo
		End If

	End IF ' IF trim(sVouType) = "P" then
oDOM.Save Server.MapPath ("../temp/transaction/Voucher Amd_"&sVouName&"_"&Session.SessionID&".xml")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Bank Voucher - Instrument Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<XML ID="NewData">
<Root>
</Root>
</XML>
<!--XML ID="OutData" src="<%="../temp/transaction/Voucher Amd_"&sVouName&"_"&Session.SessionID&".xml"%>"-->
<XML ID="OutData">

<!--BankInstrumentDet>
</BankInstrumentDet-->

</XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<Script Language=vbscript>
Dim objTemp,Root,newElem,Hnode
set objTemp = window.dialogArguments
'*************************************************************************************************
Function OptFun(obj)
	document.formname.hInsType.value  = Obj.value

	if trim(Obj.value) = "C" then
		IF trim(document.formname.hFlag.value) = "True" then
			document.formname.SelInsNo.disabled = False
			document.formname.txtInsNo.disabled = True
			document.formname.txtInsNo.value = ""
		Else
			document.formname.txtInsNo.disabled = False
		End IF
	else
		IF trim(document.formname.hFlag.value) = "True" then
			document.formname.SelInsNo.disabled = True
			document.formname.SelInsNo.value = "0"
			document.formname.txtInsNo.disabled = False
		End IF
		if trim(obj.value)="T" then
			document.formname.txtInsNo.value = "0"
		else
			document.formname.txtInsNo.value = ""
		end if
	end if

End Function
'*************************************************************************************************
Function CheckSubmit()

set Root = OutData.documentElement
	 'alert(Root.xml)
	'alert(document.formname.hInsType.value)

	' alert("ChkSub="&objTemp.xml)

	' Set Root1 = OutData.documentElement
	 'IF Root1.haschildnodes then
	'	For each BkNode in Root1.childnodes
	'		Set AddNode = BkNode
	'		Root.Appendchild AddNode
	'	Next
	' End IF
	'alert(Root.xml)
	'exit function
Chkflag = False 'For Failure Condition
IF trim(Chkflag) = True then
	Set newElem = OutData.createElement("BankInstrumentDet")
	newElem.setAttribute "InsType", sInsTy
	IF trim(document.formname.hFlag.value) = "True" then

		IF trim(document.formname.SelInsNo.Value) <> "0" then
			newElem.setAttribute "InsNo",document.formname.SelInsNo.Value
			newElem.setAttribute "InsDate",document.formname.ctlDate.GetDate()
			newElem.setAttribute "PayAt",document.formname.txtPayableAt.value
			newElem.setAttribute "DrawnOn",document.formname.txtDrawnOn.value
			newElem.setAttribute "Option","Y" 'Option Selected

		Else
			newElem.setAttribute "InsNo",""
			newElem.setAttribute "InsDate",""
			newElem.setAttribute "PayAt",""
			newElem.setAttribute "DrawnOn",""
			newElem.setAttribute "Option","N" 'Option Not Selected

		End If
	Else
		newElem.setAttribute "InsNo",document.formname.txtInsNo.Value
		newElem.setAttribute "InsDate",document.formname.ctlDate.GetDate()
		newElem.setAttribute "PayAt",document.formname.txtPayableAt.value
		newElem.setAttribute "DrawnOn",document.formname.txtDrawnOn.value
		newElem.setAttribute "Option","T" ''Option is Text
	End IF
	IF trim(document.formname.hTransNo.value) <> "0" then
		newElem.setAttribute "Action",document.formname.SelAct.value
	Else
		newElem.setAttribute "Action","0"
	End IF

	Root.appendChild newElem
End IF 'IF 	Chkflag = True then
'	alert(Root.xml)

	'set objhttp = CreateObject("Microsoft.XMLHTTP")
	'objhttp.Open "POST","XMLSave.asp?Name=Voucher Amd&Mod=BA", false
	'objhttp.send OutData.XMLDocument
window_onunload()
	'window.close()
End Function
'*************************************************************************************************
Function window_onunload()
	'alert ObjTemp.xml
	set window.returnValue = ObjTemp.documentElement
	window.close()
end Function
'*************************************************************************************************
Function InsType()
	IF trim(document.formname.hFlag.value) = "True" then
		IF document.formname.optInsType(0).checked then
			document.formname.SelInsNo.disabled = False
			document.formname.txtInsNo.disabled = True
		Else
			document.formname.SelInsNo.disabled = True
			document.formname.txtInsNo.disabled = False
		End IF
	End IF
End Function
'*************************************************************************************************
Function SelAction()
	IF trim(document.formname.SelAct.value) <> "0" then
		document.formname.SelInsNo.disabled = False
	Else
		sTemp = split(document.formname.hUsInsNo.value,"-")
		sText = sTemp(2)
		document.formname.SelInsNo.options(document.formname.SelInsNo.selectedIndex).value = document.formname.hUsInsNo.value
		document.formname.SelInsNo.options(document.formname.SelInsNo.selectedIndex).text = sTemp(2)
		document.formname.SelInsNo.disabled = True
	End IF
End Function
'*************************************************************************************************
Function Init()

	IF (document.formname.hTransNo.value) <> "0" then
		set Root = OutData.documentElement
		'alert Root.xml
	Else
		set objTemp = window.dialogArguments
	End IF

	' Alert(objTemp.xml)


	IF document.formname.hDate.value <> "" then
		document.formname.ctlDate.setDate = document.formname.hDate.value
	Else
		document.formname.ctlDate.setDate =date()
	End IF

	'DisplayDet()
	sExp = "//BankInstrumentDet"
	Set TempNode = objTemp.selectNodes(sExp)
	'Msgbox TempNode.length
	IF TempNode.length <> 0 Then
		DisplayTable()
	End IF
End Function
'*************************************************************************************************

Function DisplayDet()
	Dim sExp,TempNode,iCtr,iInsNo,sInsDate,sPayat,sDrawn,sInsTy

	Set Root = OutData.documentElement
	'alert(Root.xml)
	sExp = "//BankInstrumentDet"
	Set TempNode = Root.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		sInsTy = TempNode.Item(0).Attributes.getNamedItem("InsType").Value
		iInsNo = TempNode.Item(0).Attributes.getNamedItem("InsNo").Value
		sInsDate = TempNode.Item(0).Attributes.getNamedItem("InsDate").Value
		sPayat = TempNode.Item(0).Attributes.getNamedItem("PayAt").Value
		sDrawn = TempNode.Item(0).Attributes.getNamedItem("DrawnOn").Value

		IF CStr(sInsTy) = "C" Then
			document.formname.optInsType(0).checked = True
		Elseif CStr(sInsTy) = "D" Then
			document.formname.optInsType(1).checked = True
		Elseif CStr(sInsTy) = "B" Then
			document.formname.optInsType(2).checked = True
		Elseif CStr(sInsTy) = "W" Then
			document.formname.optInsType(4).checked = True
		Else
			document.formname.optInsType(3).checked = True
		End IF
		'Msgbox sInsDate
		'alert("Inside="&iInsNo)
		'alert document.formname.hFlag.value
		IF trim(document.formname.hFlag.value) = "True"  then
			IF trim(document.formname.hTransNo.value) = "0" then
				'document.formname.SelInsNo.options(document.formname.SelInsNo.selectedIndex).value = iInsNo
				IF trim(iInsNo) <> "" then document.formname.SelInsNo.options(document.formname.SelInsNo.selectedIndex).text =  iInsNo
			End IF
		Else
			document.formname.txtInsNo.value = iInsNo
		End IF
		document.formname.txtPayableAt.value = sPayat
		document.formname.txtDrawnOn.value = sDrawn
		document.formname.ctlDate.setDate = Trim(sInsDate)
'		TempNode.removeAll
	Else
		IF trim(document.formname.hTransNo.value) <> "0" then
			IF  trim(document.formname.hUsInsNo.value) <> "" then
				IF trim(document.formname.hFlag.value) = "True" then
				' alert document.formname.hUsInsNo.value
					sTemp = split(document.formname.hUsInsNo.value,"-")
					sText = sTemp(2)
					document.formname.SelInsNo.options(document.formname.SelInsNo.selectedIndex).value = document.formname.hUsInsNo.value
					document.formname.SelInsNo.options(document.formname.SelInsNo.selectedIndex).text =  sText
				Else
					document.formname.txtInsNo.value = iInsNo
				End IF
			End IF
		End IF ' IF trim(document.formname.hTransNo.value) <> "0" then
	End IF
End Function
'*************************************************************************************************
Function AddFun()

	Set Root = ObjTemp.documentElement

	sExists = document.formname.hExists.value
	 'alert Root.xml
	If Root.haschildnodes then
		For each InsNode in Root.childnodes
			If trim(InsNode.NodeName) = "BankInstrumentDet" then
				SlNo = InsNode.getAttribute("SlNo")
				If InsNode.getAttribute("Option") = "Y" then
					sTemp = split(InsNode.getAttribute("InsNo"),"-")
					sInsNo = sTemp(2)
				End If
					sAct = InsNode.getAttribute("Action")
					iAmt = InsNode.getAttribute("InsAmt")
				If trim(sExists) = "Y" then
					If trim(SlNo) = trim(document.formname.hEditNo.value) then
						IF document.formname.optInsType(0).checked = True then InsNode.SetAttribute "InsType","Cheque"
						IF document.formname.optInsType(1).checked = True then InsNode.SetAttribute "InsType","Demand Draft"
						IF document.formname.optInsType(2).checked = True then InsNode.SetAttribute "InsType","Bankers Cheque"
						IF document.formname.optInsType(3).checked = True then InsNode.SetAttribute "InsType","RTGS"
						IF document.formname.optInsType(4).checked = True then InsNode.SetAttribute "InsType","Cash Deposited"

						InsNode.SetAttribute "InsNo",document.formname.txtInsNo.value
						InsNode.SetAttribute "InsDate",document.formname.ctlDate.getDate
						InsNode.SetAttribute "PayAt",document.formname.txtPayableAt.value
						InsNode.SetAttribute "DrawnOn",document.formname.txtDrawnOn.value
						InsNode.SetAttribute "InsAmt",document.formname.txtAmount.value
						IF document.formname.txtInsNo.disabled = "True"  then
							InsNode.setAttribute "Option","Y"
						Else
							InsNode.setAttribute "Option",""
						End IF

						IF trim(document.formname.hTransNo.value) <> "0" then
							InsNode.setAttribute "Action",document.formname.SelAct.value
						Else
							InsNode.setAttribute "Action","0"
						End IF
					End IF
				End If	'If trim(sExists) = "Y" then

			End If
		Next

	End If
'	alert("BankInstrumentDet="&Root.xml)
	'alert("Exist="&	sExists)

	'alert sExists
	If trim(sExists) <> "Y" then
		Set node = ObjTemp.createElement("BankInstrumentDet")
		If document.formname.hCtr.value = "" then document.formname.hCtr.value = 1
		node.SetAttribute "SlNo",document.formname.hCtr.value

		IF document.formname.txtInsNo.disabled = "True"  then
			 'node.SetAttribute "InsNo",document.formname.SelInsNo.options(document.formname.SelInsNo.selectedIndex).text
			IF trim(document.formname.SelInsNo.value) = "S" or trim(document.formname.SelInsNo.value) = "0"	then
				Alert("Select Instrument No")
				Exit Function
			End If

			IF trim(document.formname.txtAmount.value) = "" then
				Alert("Enter Instrument Amount")
				Exit Function
			elseIF trim(document.formname.txtAmount.value) <= "0" then
				Alert("Instrument Amount Should be Greater than 0")
				Exit Function
			End If
			node.SetAttribute "InsNo",document.formname.SelInsNo.value
		Else
			IF trim(document.formname.txtInsNo.value) = "" then
				Alert("Enter Instrument No")
				Exit Function
			End If
			IF trim(document.formname.txtAmount.value) = "" then
				Alert("Enter Instrument Amount")
				Exit Function
			elseIF trim(document.formname.txtAmount.value) <= "0" then
				Alert("Instrument Amount Should be Greater than 0")
				Exit Function
			End IF
			IF Not isNumeric(document.formname.txtAmount.value)  then
				document.formname.txtAmount.value = ""
				Alert("Enter Numeric Values")
				Exit Function
			End If
			node.SetAttribute "InsNo",document.formname.txtInsNo.value
		End IF
		IF document.formname.optInsType(0).checked = True then node.SetAttribute "InsType","Cheque"
		IF document.formname.optInsType(1).checked = True then node.SetAttribute "InsType","Demand Draft"
		IF document.formname.optInsType(2).checked = True then node.SetAttribute "InsType","Bankers Cheque"
		IF document.formname.optInsType(3).checked = True then node.SetAttribute "InsType","RTGS"
		IF document.formname.optInsType(4).checked = True then node.SetAttribute "InsType","Cash Deposited"

		node.SetAttribute "InsDate",document.formname.ctlDate.getDate
		node.SetAttribute "PayAt",document.formname.txtPayableAt.value
		node.SetAttribute "DrawnOn",document.formname.txtDrawnOn.value
		node.SetAttribute "InsAmt",document.formname.txtAmount.value
		IF document.formname.txtInsNo.disabled = "True"  then
			node.setAttribute "Option","Y"
		Else
			node.setAttribute "Option",""
		End IF

		IF trim(document.formname.hTransNo.value) <> "0" then
			node.setAttribute "Action",document.formname.SelAct.value
		Else
			node.setAttribute "Action","0"
		End IF

		Root.AppendChild node

	End IF 'If trim(sExists) <> "Y" then
'alert(Root.xml)
document.formname.B2.disabled  = False
DisplayTable()
End Function
'*************************************************************************************************
Function ClearTable()
	Dim i
	for	i = 1 to document.all.InsTab.rows.length - 1
		document.all.InsTab.deleteRow(1)
	next
End function
'*************************************************************************************************
Function DeleteItems()
'Set Root = OutData.documentElement
Set Root = objTemp.documentElement
'alert(Root.xml)
K = 0
IF Root.haschildnodes then
	For each InsNode in Root.childnodes
		IF trim(InsNode.NodeName) = "BankInstrumentDet" then

			SlNo	= InsNode.getAttribute("SlNo")
			InsNo	= InsNode.getAttribute("InsNo")
			sOption = InsNode.getAttribute("Option")
			IF Trim(sOption) = "Y" then
				sTemp = split(InsNo,"-")
				ChkInsNo = sTemp(2)
			Else
				ChkInsNo = InsNo
			End IF

			For J = 1 to document.formname.hCtr.value-1
				Set Obj = eval("document.formname.ChkInsNo"&J)
				IF obj.checked = True then
				 	'alert "obj:"& obj.Value
					sVal = split(obj.Value,":")
					'alert(slNo &"="& sVal(0) &"and"& ChkInsNo &"="&sVal(1))
					If trim(slNo) = trim(sVal(0)) and trim(ChkInsNo) = trim(sVal(1)) then

						Set InsDetNode = InsNode
						 sFlag = True
					End IF
				End If
			Next
		End If
	Next
	IF sFlag = True then Root.Removechild InsDetNode

End IF
IF Root.haschildnodes then
iCtr = 1
	For each InsNode in Root.childnodes
		IF trim(InsNode.NodeName) = "BankInstrumentDet" then
			InsNode.setAttribute "SlNo",iCtr

			iCtr = iCtr + 1
		End IF
	Next
End IF
document.formname.hCtr.value = iCtr
DisplayTable()

End Function
'*************************************************************************************************
Function DispVal(SlNo,InsNo,OptType,InsDate,PayAt,DrawnOn,sOption,InsAmt)

If trim(sOption) = "Y" then

	sTemp   = split(InsNo,"-")
	document.formname.SelInsNo.options(document.formname.SelInsNo.selectedIndex).Value = InsNo
	document.formname.SelInsNo.options(document.formname.SelInsNo.selectedIndex).text =sTemp(2)
Else
	document.formname.txtInsNo.value =InsNo
End IF
document.formname.ctlDate.setDate		= InsDate
document.formname.txtPayableAt.value	= PayAt
document.formname.txtDrawnOn.value		= DrawnOn
document.formname.txtAmount.value		= InsAmt
'document.formname.B2.disabled  = True
document.formname.hExists.value = "Y"
document.formname.hEditNo.value = SlNo
End Function
'*************************************************************************************************
Function DisplayTable()
	ClearTable()
	'alert document.formname.hTransNo.value
	Set Root = objTemp.documentElement
	'alert("Disp="&Root.xml)
	IF Root.haschildnodes then
		For each Node in Root.childnodes
			IF trim(Node.NodeName) = "BankInstrumentDet" then
				SlNo	= Node.getAttribute("SlNo")
				InsNo	= Node.getAttribute("InsNo")
				InsDate = Node.getAttribute("InsDate")
				PayAt	= Node.getAttribute("PayAt")
				DrawnOn = Node.getAttribute("DrawnOn")
				OptType = trim(Node.getAttribute("InsType"))
				sOption = Node.getAttribute("Option")
				iInsAmt = Node.getAttribute("InsAmt")
				'alert(OptType)

				if OptType="RTGS" then
					document.formname.hInsType.value = "T"
					document.formname.optInsType(3).checked = true
				elseif OptType="Cheque" then
					document.formname.hInsType.value = "C"
					document.formname.optInsType(0).checked = true
				elseif OptType="Demand Draft" then
					document.formname.hInsType.value = "D"
					document.formname.optInsType(1).checked = true
				elseif OptType="Bankers Cheque" then
					document.formname.hInsType.value = "B"
					document.formname.optInsType(2).checked = true
				elseif OptType="Cash Withdrawn" or OptType="Cash Deposited" then
					document.formname.hInsType.value = "W"
					document.formname.optInsType(4).checked = true
				end if

				'alert InsNo
				'alert sOption
				IF trim(InsNo) <> "" then
					IF trim(sOption) = "Y"  then
						sTemp = split(InsNo,"-")
						iEntNo = sTemp(0)
						iInsEntNo = sTemp(1)
						InsNo = sTemp(2)
						sVal = iEntNo&"-"&iInsEntNo&"-"&InsNo
					End IF
				Else
					sTemp = split(document.formname.SelInsNo.value ,"-")
					InsNo = sTemp(2)
				End IF
				'If document.formname.hCtr.value = 1 and  SlNo = 1 then
				If document.all.InsTab.rows.length = 0 then
					set trow=document.all.InsTab.Insertrow(document.all.InsTab.rows.length)
					'Sl.No
					set Cell = trow.InsertCell()
					Cell.innerHTML= "Sl.No"
					Cell.className="ExcelHeaderCell"
					cell.width="3"
					Cell.align="center"
					'CheckBox
					set Cell = trow.InsertCell()
					Cell.innerHTML="<a href=""#""><img border=""0"" src=""../../assets/images/iTMS%20Icons/DeleteIcon.gif"" width=""15"" height=""15"" onClick=""DeleteItems()""></a>"
					Cell.className="ExcelHeaderCell"
					Cell.align="center"

					'InsNo
					set Cell = trow.InsertCell()
					Cell.innerHTML= "Instrument No"
					Cell.className="ExcelHeaderCell"
					Cell.align="center"
					Cell.width="7"
					'Ins Date
					set Cell = trow.InsertCell()
					Cell.innerHTML= "Instrument Date"
					Cell.className="ExcelHeaderCell"
					Cell.align="center"
					Cell.width="12"
					'Pay At
					set Cell = trow.InsertCell()
					Cell.innerHTML= "Payble At"
					Cell.className="ExcelHeaderCell"
					Cell.align="center"
					Cell.width="15"
					'Drawn On
					set Cell = trow.InsertCell()
					Cell.innerHTML= "Drawn On"
					Cell.className="ExcelHeaderCell"
					Cell.align="center"
					Cell.width="15"
					'InsType
					set Cell = trow.InsertCell()
					Cell.innerHTML= "Instrument Type"
					Cell.className="ExcelHeaderCell"
					Cell.align="center"
					Cell.width="22"
					'iInsAmt
					set Cell = trow.InsertCell()
					Cell.innerHTML= "Instrument Amount"
					Cell.className="ExcelHeaderCell"
					Cell.align="center"
					Cell.width="12"
				End IF 'IF document.formname.hCtr.value  = 1 then

				set trow=document.all.InsTab.Insertrow(document.all.InsTab.rows.length)
				set Cell = trow.InsertCell()
				Cell.innerHTML= SlNo
				Cell.width = 3
				Cell.className="ExcelSerial"
				Cell.align="center"

				set Cell=trow.insertCell()
				IF trim(sOption) = "Y"  then
					set oText = document.createElement("<input type=""CheckBox"" name=""ChkInsNo"& SlNo &""" value="""&SlNo&":"&InsNo&""" class=""Formelem"" onclick=""DispVal('"&SlNo&"','"&sVal&"','"&OptType&"','"&InsDate&"')"">")
				Else
					set oText = document.createElement("<input type=""CheckBox"" name=""ChkInsNo"& SlNo &""" value="""&SlNo&":"&InsNo&""" class=""Formelem"" onclick=""DispVal('"&SlNo&"','"&InsNo&"','"&OptType&"','"&InsDate&"','"&PayAt&"','"&DrawnOn&"','"&sOption&"','"&iInsAmt&"')"">")
				End IF
				Cell.appendChild(oText)
				Cell.width = 3
				Cell.className="ExcelDisplayCell"
				Cell.align="center"


				set Cell=trow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""InsNo"" value="""&InsNo&"""  style=""text-align: Left"" size=""7"" maxlength=""6"" class=""FormelemRead"">")
				Cell.appendChild(oText)
				Cell.width = 7
				Cell.className="ExcelDisplayCell"
				Cell.align="Left"

				set Cell=trow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""InsDate"" value="""&InsDate&"""  style=""text-align: center"" size=""12""  class=""FormelemRead"">")
				Cell.appendChild(oText)
				Cell.width = 12
				Cell.className="ExcelDisplayCell"
				Cell.align="center"

				set Cell=trow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""PayAt"" value="""&PayAt&"""  style=""text-align: Left"" size=""15""  class=""FormelemRead"">")
				Cell.appendChild(oText)
				Cell.width = 15
				Cell.className="ExcelDisplayCell"
				Cell.align="Left"

				set Cell=trow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""DrawnOn"" value="""&DrawnOn&"""  style=""text-align: Left"" size=""15""  class=""FormelemRead"">")
				Cell.appendChild(oText)
				Cell.width = 15
				Cell.className="ExcelDisplayCell"
				Cell.align="Left"


				set Cell=trow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""OptType"" value="""&OptType&"""  style=""text-align: Left"" size=""22""  class=""FormelemRead"">")
				Cell.appendChild(oText)
				Cell.width = 22
				Cell.className="ExcelDisplayCell"
				Cell.align="Left"

				'iInsAmt
				set Cell=trow.insertCell()
				set oText = document.createElement("<input type=""text"" name=""InsAmt"" value="""&iInsAmt&"""  style=""text-align: Right"" size=""12""  class=""FormelemRead"">")
				Cell.appendChild(oText)
				Cell.width = 12
				Cell.className="ExcelDisplayCell"
				Cell.align="Right"

				document.formname.hCtr.value = SlNo +  1
			End IF
		Next
		'document.formname.hCtr.value = document.formname.hCtr.value + 1
	End IF
	document.formname.txtInsNo.value  = ""
	document.formname.ctlDate.setDate  = date()
	document.formname.txtPayableAt.value  = ""
	document.formname.txtDrawnOn.value  = ""
	document.formname.txtAmount.value  = ""
	document.formname.hExists.value = ""
	'alert "InsideTable="&document.formname.hCtr.value

End Function
'*************************************************************************************************
</Script>
<script language="javascript" src="../scripts/ModalReturnCompat.js"></script>
<script language="javascript">
window.ITMSModalReturnCompat.install(function () {
	return window.ITMSModalReturnCompat.dialogArgumentsRoot();
});
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init()">

<form method="POST" name="formname" action="">
<Input Type="hidden" name="hExists" Value="">
<Input Type="hidden" name="hEditNo" Value="">

<Input Type="hidden" name="hTransNo" value="<%=iTransNo%>">
<Input Type="hidden" name="hUsInsNo" value="<%=iUsInsNo%>">
<Input Type="hidden" name="hVouDate" Value="<%=sVouDate%>">
<input type="hidden" name="hDate" value="<%=dtIssue%>">
<input type="hidden" name="hFlag" value = "<%=sFlag%>">
<input type="hidden" name="hCtr" value="1">
<input type="hidden" name="hInsType" value="C">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Instrument Details
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
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
													<table cellpadding="0" cellspacing="0">
                                                        <!--tr>
															<td class=FieldCell> Name of First Deposit Holder</td>
															<td class="FieldCellSub">
                                                            <span class="DataOnly">

                                                            </span>
                                                            </td>
                                                        </tr>

                                                        <tr>
															<td class=FieldCell> Name of Second Deposit Holder</td>
															<td class="FieldCellSub">
                                                            <span class="DataOnly">

                                                            </span>
                                                            </td>
                                                        </tr>

                                                        <tr>
															<td class=FieldCell> Name of Guardian</td>
															<td class="FieldCellSub">
                                                            <span class="DataOnly">

                                                            </span>
                                                            </td>
                                                        </tr>


                                                        <tr>
															<td class=FieldCell> Deposit amount in Rupees</td>
															<td class="FieldCellSub">
                                                            <span class="DataOnly">Rs &nbsp;
                                                            <%=iAmt%>
                                                            </span>
                                                            </td>
                                                        </tr>
                                                        -->
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
                                                <table  cellpadding="0" cellspacing="0">
                                            <tr>
                                        <td>
                                        <table cellpadding="0" cellspacing="0" width="100%">
                                    <tr>
                                <td class="GroupTitleLeft" width="10"><p align="left">&nbsp;</p>
                                </td>
                                <td class="GroupTitle" width="126" align="center"><p align="center">Instrument Details</td>
                                <td class="GroupTitleRight"><p align="left">&nbsp;</td>
                                    </tr>
                                        </table>
                                        </td>
                                            </tr>
                                            <tr>
                                        <td class="GroupTable">
                                        <table cellpadding="0" cellspacing="0">
                                    <tr>
                                <td class="MiddlePack" colspan="5"><p align="left"></td>
                                    </tr>
                                    <tr>
                                <td class="FieldCellSub"><p align="left">Instrument Type</p>
                                </td>
                                <td class="FieldCellSub" colspan="6"><p align="left">
                                <Input type="radio" name="optInsType" value="C" checked class="FormElem" onclick="OptFun(this)"> Cheque
                                  &nbsp;<input type="radio" name="optInsType" value="D" class="FormElem" onclick="OptFun(this)"> Demand Draft
                                  &nbsp;<input type="radio" name="optInsType" value="B" class="FormElem" onclick="OptFun(this)">
                                Bankers
                                Cheque&nbsp;<input type="radio" name="optInsType" value="T" class="FormElem" onclick="OptFun(this)">
                                RTGS&nbsp;
                                <input type="radio" name="optInsType" value="W" class="FormElem" onclick="OptFun(this)">
                                <%IF CStr(sVouTy) = "C" Then
										Response.Write "Cash Withdrawn"
								  Else
										Response.Write "Cash Deposited"
								  End IF
                                %>

                                </td>

                                    </tr>
                                    <tr>
										<td class="FieldCellSub"><p align="left">Instrument Number
										</td>
										<td class="FieldCellSub"><p align="left">
										<% IF trim(sPrintCheq) = "1" then %>
											<input type="text" name="txtInsNo" size="10" class="Formelem" maxlength="6" disabled>&nbsp;
											<%If trim(iTransNo) <> "0" then %>
												<Select size=1 name="SelInsNo" class=FormElem disabled>
											<%Else%>
												<Select size=1 name="SelInsNo" class=FormElem >
											<%End IF%>

											<option value=0>Select</option>
											<%
											 	sQry = "Select EntryNo,InstrumentEntryNo,InstrumentNo from Acc_R_BankInstrumentUsage where EntryNo = "&iEntNo&" and Status = 'N'"
											 	Response.Write sQry
											 	rs.Open sQry,con
											 	Do while not rs.Eof
											 	%>
											 	<option value="<%=rs(0)%>-<%=rs(1)%>-<%=rs(2)%>"><%=rs(2)%></option>
											 	<%
											 	rs.MoveNext
											 	loop
											 	rs.Close
											%>
											</Select>
										<% Else  %>
											<input type="text" name="txtInsNo" size="10" class="Formelem" maxlength="6">
										<% End IF 'IF trim(sPrintCheq) = "1" then%>

										<%If trim(iTransNo) <> "0" then %>

											&nbsp;<Select size="1" name="SelAct" class="Formelem" onchange="SelAction()">
											<option value="0">Select</option>
											<option value="C">Cancel</option>
											<option value="R">Reuse</option>
											</Select>

										<%End If %>
										</td></p>

										<td class="FieldCellSub"></td>
										<td class="FieldCellSub"><p align="left">Payable at
										</td>
										<td class="FieldCellSub"><p align="left">
										<!--<input type="text" name="txtIntDate" size="11" class="Formelem" value="<%=Formatdate(date)%>">-->
										<input type="text" name="txtPayableAt" size="20" value="<%=sPayAt%>"  class="Formelem">
										</td>
									</tr>

									<tr>
										<td class="FieldCellSub"><p align="left">Instrument Date
										</td>
										<td class="FieldCellSub"><p align="left">
										<% ' Function Call to Insert Date Picker
											'	Response.Write InsertDatePicker("ctlDate")
										%><object id="ctlDate" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"   codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="formelem" viewastext>
											<param name="_ExtentX" value="2355">
											<param name="_ExtentY" value="529">
										</object></p>
										</td>
										<td class="FieldCellSub"></td>
										<td class="FieldCellSub"><p align="left">Drawn On
										</td>
										<td class="FieldCellSub"><p align="left">
										<!--<input type="text" name="txtIntDate" size="11" class="Formelem" value="<%=Formatdate(date)%>">-->
											<input type="text" name="txtDrawnOn" value="<%=sDrawOn%>" size="20" class="Formelem">
										</td>

									</tr>
									<tr>
									<td class="FieldCellSub"><p align="left">Instrument Amount</td>
										<td class="FieldCellSub"><p align="left">
											<input type="text" name="txtAmount" value="" size="15" class="Formelem">
										</td>
										<td class="FieldCellSub"><p align="left">
										<td class="FieldCell"><p align="left">
											<input type="Button" name="ButAddList" Value="Add To List" class="ActionButtonX" Onclick="AddFun()">
										</td>

									</tr>
									<tr>
												<td class="FieldCell" colspan="8"><p align="left">
												<div class="frmBody" id="frm1" style="width: 655; height:80;">
													<table ID="InsTab" border="0" cellspacing="1" class="ExcelTable" width ="100%" ></table>
												</div>
												</td>
											</tr>

                                <td class="MiddlePack" colspan="5"><p align="left"></td>
                                    </tr>

                                        </table>

                                        </td>

                                            </tr>
                                            <tr>

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
                                                                <input type="button" value="Done" name="B2" class="ActionButton" onclick="CheckSubmit()" >
                                                                <input type="reset" value="Reset" name="B1" class="ActionButton" tabindex="4" >
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
