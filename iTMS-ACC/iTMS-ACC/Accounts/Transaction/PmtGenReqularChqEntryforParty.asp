<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PmtGenReqularChqEntryforParty.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	April  24, 2003
	'Modified By				:	UmaMaheswari S
	'Modified On				:	April 06, 2011
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<%
dim objRs,sQuery,objRs1
dim sOrgName,iPymtNo,sPymtFor,sOrgId,iAmount
dim sDate,sRequestBy,dAmount,sPayTo,sReason,nTransactionNo
dim sTemp,sApprovedBy,iSno,sReqType,sReqFrom,sParTemp,iPartyVal
dim iParCode,iParSubType,sParType,sParName,sParSubTypeName
set objRs  = server.CreateObject("adodb.recordset")
set objRs1  = server.CreateObject("adodb.recordset")

'sTemp=split(trim(Request("selRequestNo")),"?")

'iPymtNo=sTemp(0)
'sRequestBy=sTemp(1)
'sApprovedBy=sTemp(2)
iPymtNo = Trim(Request("RequestNo"))

sOrgName=Request("hUnitName")
sOrgId=Request("hUnitNo")
sReqType=Request("hReqTypeS")
'sReqFrom= Request("selReqFrom")

'Response.Write iPymtNo & sReqType
sQuery = "Select Distinct RequestedBy,isNull(ApprovedBy,0) from Acc_T_PaymentRequestHdr where PaymentRequestNo = "& iPymtNo &" "
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing

IF Not objRs.EOF then

	with objRs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "select EmployeeName from Ms_EmployeeMaster where EmployeeNumber="&objRs(0)
		.ActiveConnection = con
		.Open
	end with
	set objRs1.ActiveConnection = nothing	
	IF Not objRs1.EOF Then		
		sRequestBy=objRs1(0)
	End IF
	objRs1.Close
				
	with objRs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "select EmployeeName from Ms_EmployeeMaster where EmployeeNumber="&objRs(1)
		.ActiveConnection = con
		.Open
	end with
	set objRs1.ActiveConnection = nothing	
	IF Not objRs1.EOF Then		
		sApprovedBy=objRs1(0)
	End IF
	objRs1.Close
	
End IF
objrs.Close 
'-------------Newly Added on Dec 28 th 2007 by Maheswari to fetch Party Name -------------------
sQuery = "Select AccUnitPartyType,AccUnitPartySubType,AccUnitPartyCode,isNull(TransactionNumber,0) from Acc_T_PaymentRequestDet where PaymentRequestNo = "& iPymtNo &" "
with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing
'Response.Write "<p>sQuery="&sQuery
IF Not objRs.EOF then
	sParType	= objRs(0)
	iParSubType = objRs(1)
	iParCode	= objRs(2)
	nTransactionNo = objrs(3)
	
End IF 
objrs.Close 
If iParSubType <> "" Then
	sQuery = "Select isNull(PartyName,''),isNull(SubTypeName,'') from VwOrgParty where PartyCode = "&iParCode &" and PartyType = '"&sParType&"' and PartySubType = "& iParSubType &" "  	
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		'Response.Write "<p>sQuery="&sQuery
		.Open
	end with

	set objRs.ActiveConnection = nothing
	IF Not objRs.EOF then
		sParName		= objRs(0)														
		sParSubTypeName = objRs(1)			
		sParTemp = sParType &"-"& sParSubTypeName											
	End IF 
	objrs.Close 
End IF
iPartyVal = sParType&"?"&iParSubType&"?"&sParTemp&"?"&iParCode
'Response.Write iPartyVal
'------------------------------------------------------------------------------------------------------------------                                                    
sQuery="select ReasonForPayment,AmountToPay,ToBePaidTo,PayablesNumber from Acc_T_PaymentRequestDet where PaymentRequestNo="&iPymtNo
	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set objRs.ActiveConnection = nothing
		sPayTo=objRs(2)		
		iAmount=FormatNumber(objRs(1),2,,,0)
		
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<XML id="VoucherData">
	<voucher UnitNo="<%=sOrgId%>" UnitName="<%=sOrgName%>" BookNo="0" BookName="0" CRDR="C" VouDate="" BookAcchead="" Approver=""/>
</XML>
<XML id="EntryData">
	<Entry No="0" CRDR="D" Payto="" Amount="" AccUnit="" AccName=""/></XML>
<XML id="RequestData">
	<RequestDetails/></XML>
</XML>
<XML id="AccHeadData"></XML>
	<account/>
</XML>
<XML id="PartyData"><Root></Root></XML>
<XML id="TempXMLData"><Root></Root></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<SCRIPT language="javascript" SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<script language="javascript" src="../../scripts/checkdate.js"></script>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<SCRIPT language="vbscript">
Dim iEntryNo,VouRoot,EntryRoot,bVouFlag,bSavFlag
dim RequestRoot
iEntryNo=1
bVouFlag=false
bSavFlag=false

set VouRoot=VoucherData.documentElement
set EntryRoot=EntryData.documentElement
set RequestRoot=RequestData.documentElement

function selAccountHead(objAcc)
if objAcc.selectedIndex >0 then
	if objAcc.selectedIndex >1 then
		'If selected account Head is Party type
		sTemp=objAcc.value& "?" & objAcc.options(objAcc.selectedIndex).text
		showPartyHead document.formname.hUnitId.value ,sTemp
	else
		IF document.formname.selBookId.selectedIndex = 0 Then
			MsgBox "Select Book "
			document.formname.selBookId.focus()
			Exit Function
		Else
			showGLHead(document.formname.hUnitId.value)
		End IF
	End if 'End of select Account Head Type check GL or PARTY
End if 'End of If any Account Head Selected Check
End function
'---------------------End Of Function selAccountHead----------------------
function showPartyHead(sOrgId,sPartyType)
dim sPartyCode,bRecivable,bPayable
dim sDocNo,sInvNo,sInvDate,sAmtRec,sAmtRecd
dim nodAccHead,nodPayRec,nodCC,iSno,objhttp
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth


set objhttp = CreateObject("MSXML2.XMLHTTP")

		sTempValWindowSize = GetWindowSizeForPopup("2")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)
		
	    Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgid="&sOrgId&"&Party="&sPartyType,PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))
	    if ucase(trim(sAct)) <> "CLOSE" then
		    do while sAct <> "DONE"
			    set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
			    sAct = UCase(trim(OutValue.getAttribute("Action")))
			    if ucase(Trim(sAct)) = "CLOSE" then exit do
			    sQuery = trim(OutValue.getAttribute("PassQuery"))
		    loop
	    end if
	    

		'	OutValue = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		'	arrTemp = split(OutValue,":")
'
'			while UBound(arrTemp) = 0 
'				OutValue = showModalDialog("PartySelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'				arrTemp = split(OutValue,":")
'			wend
'
'			sRetValue = OutValue
'			if UBound(arrTemp) <= 1 then exit function
'			sTemp = Split(sRetValue,":")
'			sParTy = sTemp(4)
'			sParSubType = sTemp(3)
'			sParCode = sTemp(1)
'			sPartyName = sTemp(0)

	
		if OutValue.hasChildNodes() then
            for each ndEntry in OutValue.childNodes
                if ndEntry.nodeName="Entry" then
                    sParTy = ndEntry.getAttribute("RetField3")
		            sParSubType = ndEntry.getAttribute("RetField4")
		            sParCode = ndEntry.getAttribute("RetField1")
		            sPartyName = ndEntry.getAttribute("RetField0")
		        exit for
                end if
            next
        end if

objhttp.Open "GET","XMLGetPayRecCount.asp?orgID="&sOrgId&"&ParSubType="&sParSubType&"&ParType=" & sParTy&"&PartyCode="&sParCode , false
objhttp.send

IF objhttp.responseText <> "" Then
	sRetVal2 = objhttp.responseText
	GetPartyHeadXml sParCode,sPartyName,sRetVal2
End IF
Set nodAccHead = AccHeadData.documentElement


'Set nodAccHead = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
if nodAccHead.hasChildNodes then
	'User Has Selected a Party
	clearXML()
	For Each HeaderNode In nodAccHead.childNodes
		bVouFlag=true
		sPartyCode=sPartyType&"?"& HeaderNode.Attributes.Item(0).nodeValue
		HeaderNode.Attributes.Item(0).nodeValue=sPartyCode
		bPayable=HeaderNode.Attributes.Item(1).nodeValue
		bRecivable=HeaderNode.Attributes.Item(2).nodeValue

		'document.formname.txtPayto.value=HeaderNode.Attributes.Item(3).nodeValue
		EntryRoot.appendChild HeaderNode
	next

	if cint(bPayable)=1 then
	'If Selected Party Has Payable or Receiavable
		Set nodPayRec = showModalDialog("PayRecSelection.asp?orgId="+sOrgId+"&ParCode="+sPartyCode&"&Type=C","","")
		if nodPayRec.Attributes.Item(0).nodeValue=1 then
			'Set the Additional Display Layer Visible
			For Each HeaderNode In nodPayRec.childNodes
					EntryRoot.appendChild HeaderNode
					if HeaderNode.hasChildNodes then
						'If user has Selected Documnets
						iSno=1
						setPayableDisplay 1
						ClearTable "tblPayable",2,1
						for each  nodCC in HeaderNode.childNodes
							sDocNo=nodCC.Attributes.Item(0).nodeValue
							sInvNo=nodCC.Attributes.Item(1).nodeValue
							sInvDate=nodCC.Attributes.Item(2).nodeValue
							sAmtRec=nodCC.Attributes.Item(3).nodeValue
							sAmtRecd=nodCC.Attributes.Item(4).nodeValue

							set oRow = document.all.tblPayable.insertRow(iSno+1)
							InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
							InsertCell oRow,1,"",sDocNo,"ExcelDisplayCell","","",0,0,0,0,""
							InsertCell oRow,1,"",sInvNo,"ExcelDisplayCell","","",0,0,0,0,""
							InsertCell oRow,1,"",sInvDate,"ExcelDisplayCell","","",0,0,0,0,""
							InsertCell oRow,1,"",sAmtRec,"ExcelDisplayCell","","",0,0,0,0,""
							InsertCell oRow,1,"",sAmtRecd,"ExcelDisplayCell","","",0,0,0,0,""
							InsertCell oRow,2,"txtDocAmount"&CStr(sDocNo),"0","ExcelInputCell","","",12,10,0,0,""
							iSno=iSno+1
						next
					end if 'End of Check Documnet Node
			next	'End of Processing PayRec Node
		else
			'User Has canceled Documnet Selection
			'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
     		setPayableDisplay 0
		end if	'End of Documnet has Childs Check
    else
		'Selected Head has no Documnets
		'Set the Additional Layer Display Layer Hidden
		setPayableDisplay 0
	end if	'End of Party Has Payable Or Recivables
else
	'User canceled Party Head Selection
	'document.formname.txtPayto.value=""
	'Set the Additional Layer Display Layer Hidden
	setPayableDisplay 0
end if 'End of Party Head Processing

set nodAccHead=nothing
set nodPayRec=nothing
set nodCC=nothing
End function
'---------------------End Of Function showGLHead--------------------------
function showGLHead(sOrgId)
dim iAccCode,bAnal,bCostCenter
dim nodAccHead,nodCCAnly,nodCC,nodANL,iSno
dim sCode,sDesc,dRatio,iBookNo,sTemp
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth

sTemp=Split(document.formname.selBookId.value,"?")

		sTempValWindowSize = GetWindowSizeForPopup("5")
        sArrTempValWindowSize = split(sTempValWindowSize,":")
        sProgramName = sArrTempValWindowSize(0)
        sPopupHeight = sArrTempValWindowSize(1)
        sPopupWidth = sArrTempValWindowSize(2)
		
		Set	OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="&sOrgId&"&BookId="&sBookid&"&BookNo="&sBookNo,TempXMLData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
		sAct = UCase(trim(OutValue.getAttribute("Action")))
		sQuery = trim(OutValue.getAttribute("PassQuery"))
		if ucase(trim(sAct)) <> "CLOSE" then
			do while sAct <> "DONE"
				set OutValue = showModalDialog("../../Common/"&sProgramName&"?"&sQuery,TempXMLData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
				sAct = UCase(trim(OutValue.getAttribute("Action")))
				if ucase(Trim(sAct)) = "CLOSE" then exit do
				sQuery = trim(OutValue.getAttribute("PassQuery"))
			loop
		end if



	'	OutValue = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=02&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	'	arrTemp = split(OutValue,":")
	'	while UBound(arrTemp) = 0 
	'		OutValue = showModalDialog("GLHeadSelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	'		arrTemp = split(OutValue,":")
	'	wend
	'		
	'	sRetVal = OutValue
	'	if UBound(arrTemp) <= 1 then exit function
	
	if OutValue.hasChildNodes() then
        for each ndEntry in OutValue.childNodes
            if ndEntry.nodeName="Entry" then
                sRetVal = ndEntry.getAttribute("RetField0")&":"&ndEntry.getAttribute("RetField1")&":"&ndEntry.getAttribute("RetField2")&":"&ndEntry.getAttribute("RetField3")&":"&ndEntry.getAttribute("RetField4")&":"&ndEntry.getAttribute("RetField5")&":"&ndEntry.getAttribute("RetField6")
            end if
        next
    end if
	
	
GetGlHeadXml(sRetVal)

Set nodAccHead = AccHeadData.documentElement


'Set nodAccHead = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=02&BookNo="+sTemp(0)+"&AccHead="+sTemp(1),"","dialogHeight:400px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
if nodAccHead.hasChildNodes then
	'User Has Selected a GL Account Head
	clearXML()
	For Each HeaderNode In nodAccHead.childNodes
		bVouFlag=true
		iAccCode=HeaderNode.Attributes.Item(0).nodeValue
		bAnal=HeaderNode.Attributes.Item(1).nodeValue
		bCostCenter=HeaderNode.Attributes.Item(2).nodeValue

		'document.formname.txtPayto.value=HeaderNode.Attributes.Item(3).nodeValue
		EntryRoot.appendChild HeaderNode
	next
	if cint(bCostCenter)=1 or cint(bAnal)=1 then
	'If Selected GL Account Head has Cost Center
		Set nodCCAnly = showModalDialog("CCAnalysisSelection.asp?orgId="+sOrgId+"&AccCode="+iAccCode,"","")
		if nodCCAnly.Attributes.Item(0).nodeValue=1 then
			'Set the Additional and CCANAL Display Layer Visible
			setADDDisplay 1
			For Each HeaderNode In nodCCAnly.childNodes
				if 	HeaderNode.nodeName="CostCenter" then
					EntryRoot.appendChild HeaderNode
					if HeaderNode.hasChildNodes then
						'If user has Selected Cost centers
						iSno=1
						setAnalDisplay "C",1
						ClearTable "tblCost",1,1
						for each  nodCC in HeaderNode.childNodes
							sCode=nodCC.Attributes.Item(0).nodeValue
							sDesc=nodCC.Attributes.Item(2).nodeValue
							dRatio=nodCC.Attributes.Item(3).nodeValue

							set oRow = document.all.tblCost.insertRow(iSno)
							InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
							InsertCell oRow,1,"",sDesc,"ExcelDisplayCell","","",0,0,0,0,""
							InsertCell oRow,2,"txtCCRatio"&CStr(sCode),dRatio,"ExcelInputCell","","",4,3,0,0,""
							InsertCell oRow,2,"txtCCAmount"&CStr(sCode),"0","ExcelInputCell","","",12,10,0,0,""

							iSno=iSno+1
						next
					else
						'No Cost Center Selected
						setAnalDisplay "C",0
					end if 'End of Check for Selected Cost centers
				end if 'End of Check for Cost Center Node
				if 	HeaderNode.nodeName="Analytical" then
					EntryRoot.appendChild HeaderNode
					if HeaderNode.hasChildNodes then
						iSno=1
						setAnalDisplay "A",1
						ClearTable "tblAnal",1,1
						for each  nodANL in HeaderNode.childNodes
							sCode=nodANL.Attributes.Item(0).nodeValue
							sDesc=nodANL.Attributes.Item(2).nodeValue
							dRatio=nodANL.Attributes.Item(3).nodeValue

							set oRow = document.all.tblAnal.insertRow(iSno)

							InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
							InsertCell oRow,1,"",sDesc,"ExcelDisplayCell","","",0,0,0,0,""
							InsertCell oRow,2,"txtANALRatio"&CStr(sCode),dRatio,"ExcelInputCell","","",4,3,0,0,""
							InsertCell oRow,2,"txtANALAmount"&CStr(sCode),"0","ExcelInputCell","","",12,10,0,0,""

							iSno=iSno+1
						next
					else
						'No Analytical Selected
						setAnalDisplay "A",0
					end if 'End of Check for Selected Analytical
				end if 'End of Check for Analytical Node
			next	'End of Processing CCANAL Node
		else
			'User Has canceled CC,ANAL Selection
			'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
     		setADDDisplay 0
		end if	'End of CC,ANAL has Childs Check
    else
		'Selected Head has no CC or ANAL
		'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
		setADDDisplay 0
	end if	'End of GL has Cost Center or not
else
	'User canceled Account Head Selection
	'document.formname.txtPayto.value=""
	'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
	setADDDisplay 0
end if 'End of GL Head Processing

set nodAccHead=nothing
set nodCCAnly=nothing
set nodCC=nothing
End function
'---------------------End Of Function showGLHead--------------------------
Function AddEntry()
dim iCode,dRatio,dAmount
	
	if Cdbl(document.formname.hAmount.value) < cdbl(document.formname.txtAmount.value) then
		Alert("Amount Should be Less then or equal to PayAmount")
		Exit function
	end if  
	
	
	IF document.formname.selBookId.selectedIndex = "0" then
		alert("Select Book")	
		exit function
	End IF
	if not checkFileds then exit function
	sTemp=Split(document.formname.selBookId.value,"?")

	VouRoot.Attributes.Item(2).nodeValue=sTemp(0)
	VouRoot.Attributes.Item(3).nodeValue=document.formname.selBookId.options(document.formname.selBookId.selectedIndex).text
	'VouRoot.Attributes.Item(4).nodeValue=document.formname.hParSubName.value  
	'VouRoot.Attributes.Item(6).nodeValue=sTemp(1)

	VouRoot.Attributes.Item(5).nodeValue=document.formname.ctlDate.GetDate


	EntryRoot.Attributes.Item(0).nodeValue=iEntryNo
	EntryRoot.Attributes.Item(1).nodeValue="D"

	EntryRoot.Attributes.Item(2).nodeValue=document.formname.hParName.value
	EntryRoot.Attributes.Item(3).nodeValue=document.formname.txtAmount.value

	EntryRoot.Attributes.Item(4).nodeValue=document.formname.hUnitId.value
	EntryRoot.Attributes.Item(5).nodeValue=document.formname.hOrgName.value
	'sVal = document.formname.hParValue.value
	Set newElem = EntryData.createElement("AccHead")
	newElem.setAttribute "No",document.formname.hParValue.value
	newElem.setAttribute "Pay","0"
	newElem.setAttribute "Rec","0"
	newElem.setAttribute "Name",document.formname.hParName.value
	newElem.setAttribute "Type","P"
	newElem.setAttribute "Adv","0"
	EntryRoot.appendChild newElem
	
	Set newElem1 = EntryData.createElement("Narration")
	newElem1.text= document.formname.txtNarration.value
	EntryRoot.appendChild newElem1

	for each HeaderNode in EntryRoot.childNodes
		if 	HeaderNode.nodeName="CostCenter" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.Item(0).nodeValue
				dRatio=eval("document.formname.txtCCRatio"&iCode).value
				dAmount=eval("document.formname.txtCCAmount"&iCode).value
				nodANL.Attributes.Item(3).nodeValue=dRatio
				nodANL.Attributes.Item(4).nodeValue=dAmount
			next
		end if 'End of Check for Cost Center Node
		if 	HeaderNode.nodeName="Analytical" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.Item(0).nodeValue
				dRatio=eval("document.formname.txtANALRatio"&iCode).value
				dAmount=eval("document.formname.txtANALAmount"&iCode).value
				nodANL.Attributes.Item(3).nodeValue=dRatio
				nodANL.Attributes.Item(4).nodeValue=dAmount
			next
		end if 'End of Check for Analytical Node
		if 	HeaderNode.nodeName="PayRec" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.Item(0).nodeValue
				dAmount=eval("document.formname.txtDocAmount"&iCode).value
				nodANL.Attributes.Item(5).nodeValue=dAmount
			next
		end if 'End of Check for Analytical Node

	next
	VouRoot.appendChild EntryRoot
	VouRoot.appendChild RequestRoot
	'alert(VouRoot.xml)
	'exit function
	SaveXML
End Function
'---------------------End Of Function AddEntry----------------------------
Function SaveXML()

	set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","XMLSave.asp?Name=Payment Request&Mod="&document.formname.hVouName.value, false
	objhttp.send VoucherData.XMLDocument
	if objhttp.responseText <> "" then
		Msgbox(objhttp.responseText)
	else
		
		document.formname.submit()
	end if
End Function
'---------------------End Of Function SaveXML-----------------------------
Function  checkFileds()
	if trim(document.formname.txtNarration.value)="" then
		Msgbox("Enter Narration")
		document.formname.txtNarration.select
		checkFileds=false
		exit Function
	end if
	if  trim(document.formname.txtAmount.value)="" then
		Msgbox("Select atleast one term in Request Details Section")
		checkFileds=false
		exit Function
	elseif cdbl(document.formname.txtAmount.value)=0 then
		Msgbox("Select atleast one term in Request Details Section")
		checkFileds=false
		exit Function
	end if
	checkFileds=true
end Function
'---------------------End Of Function checkFileds--------------------------
Function clearXML()
	Set EntryRoot = EntryData.createElement("Entry")
		EntryRoot.setAttribute "No",iEntryNo
		EntryRoot.setAttribute "CRDR","D"
		EntryRoot.setAttribute "Payto","0"
		EntryRoot.setAttribute "Amount","0"
		EntryRoot.setAttribute "AccUnit","0"
		EntryRoot.setAttribute "AccName",""
end Function
'---------------------End Of Function clearXML----------------------------

Function ClearTable(objTable,startlen,Count)
	dim i

	for i=startlen to eval("document.all."&objTable).rows.length - Count
		eval("document.all."&objTable).deleteRow(startlen)
	next
end Function
'---------------------End Of Function ClearTable--------------------------
Function setPayableDisplay(iFlag)

if iFlag=0 then
	window.Disaddtional.style.height="1px"
	window.Disaddtional.style.visibility="hidden"
	window.DisPayable.style.height="1px"
	window.DisPayable.style.visibility="hidden"
else
	window.Disaddtional.style.height="115px"
	window.Disaddtional.style.visibility="visible"
	window.DisPayable.style.height="110px"
	window.DisPayable.style.visibility="visible"
end if

end Function
'---------------------End Of Function setPayableDisplay----------------------------

Function setAnalDisplay(sDisplay,iFlag)
if sDisplay="A" then
	if iFlag=0 then
		window.DisAnal.style.height="1px"
		window.DisAnal.style.width ="1px"
		window.DisAnal.style.visibility="hidden"
	else
		window.DisAnal.style.height="100px"
		window.DisAnal.style.width ="280px"
		window.DisAnal.style.visibility="visible"
	end if
else
	if iFlag=0 then
		window.DisCost.style.height="1px"
		window.DisCost.style.width ="1px"
		window.DisCost.style.visibility="hidden"
	else
		window.DisCost.style.height="100px"
		window.DisCost.style.width ="280px"
		window.DisCost.style.visibility="visible"
	end if
end if
end Function
'---------------------End Of Function setAnalDisplay----------------------------
Function setADDDisplay(iFlag)
	if iFlag=0 then
		window.Disaddtional.style.height="1px"
		window.Disaddtional.style.visibility="hidden"
		window.DisCCANL.style.height="1px"
		window.DisCCANL.style.visibility="hidden"
	else
		window.Disaddtional.style.height="115px"
		window.Disaddtional.style.visibility="visible"
		window.DisCCANL.style.height="114px"
		window.DisCCANL.style.visibility="visible"
	end if

end Function
'---------------------End Of Function setAnalDisplay----------------------------
Function CancelAction(sPage)
	document.formname.action=sPage
	document.formname.submit
end Function
'---------------------End Of Function ActionCancel----------------------------

Function PopInsDet()
	Dim sTempValues,sExp,InsNode,sChkDet,sCurrVouTy
	sCurrVouTy = document.formname.hVouType.Value
	sVouDate = ""
	sVouCode = document.formname.hVouCode.value 
	sUnitID= document.formname.hUnitId.value
	nTransNo = document.formname.hTransNo.value
	sVouName = document.formname.hVouName.value
	sTempValues = sCurrVouTy & ":"& sVouDate & ":"&sVouCode &":"&sUnitID&":"&nTransNo&":"&sVouName
	'sTempValues = sTempValues&":"&sCurrVouTy
	'alert(sTempValues)
	Set OutDataValue = showModalDialog("BankInsDetails.asp?sTemp="&sTempValues,VoucherData,"dialogHeight:250px;dialogWidth:710px;center:Yes;help:No;resizable:No;status:No")
	'alert(OutDataValue.xml)
	sExp = "//BankInstrumentDet"
	
	Set InsNode = OutDataValue.selectNodes(sExp)
	
	IF InsNode.Length <> 0 Then
		IF CStr(InsNode.Item(0).Attributes.Item(0).nodeValue) = "C" Then
			sChkDet = "C: " 
		Elseif CStr(InsNode.Item(0).Attributes.Item(0).nodeValue) = "D" Then
			sChkDet = "D: "
		Elseif CStr(InsNode.Item(0).Attributes.Item(0).nodeValue) = "B" Then
			sChkDet = "B: "
		Elseif CStr(InsNode.Item(0).Attributes.Item(0).nodeValue) = "T" Then
			sChkDet = "T: "
		Else
			sChkDet = "Cash: "
		End IF
		'sPayAt =  InsNode.Item(0).Attributes.Item(3).nodeValue
		'sDrwOn = InsNode.Item(0).Attributes.Item(4).nodeValue
		sPayAt =  InsNode.Item(0).Attributes.Item(4).nodeValue
		sDrwOn = InsNode.Item(0).Attributes.Item(5).nodeValue
		sChkDet = sChkDet & InsNode.Item(0).Attributes.Item(1).nodeValue
		document.all.spInsNo.innerHTML = sChkDet
		sChkDet = sChkDet  
		sChkDet = sChkDet &":"& InsNode.Item(0).Attributes.Item(2).nodeValue
		document.all.spInsDate.innerHTML =  InsNode.Item(0).Attributes.Item(3).nodeValue
		
		document.formname.hInsDet.Value = sChkDet&":"&sPayAt&":"&sDrwOn & ":" & InsNode.Item(0).Attributes.Item(3).nodeValue
		'alert(document.formname.hInsDet.Value)
		
	End IF
	'alert InsNode.Length
End Function

'---------------------End Of Function PopInsDet----------------------------
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="PmtGenerate.asp">
<input type="hidden" name="hUnitId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hPaymentNo" value="<%=iPymtNo%>">
<input type="hidden" name="hParSubName" value="<%=sParSubTypeName%>">
<input type="hidden" name="hParName" value="<%=sParName%>">
<input type="hidden" name="hVouType" value="<%=Left(sParType,1)%>">
<input type="hidden" name="hAmount" value="<%=iAmount%>">
<input type="hidden" name="hInsDet" value="">
<input type="hidden" name="hParCode" value="<%=iParCode%>">
<input type="hidden" name="hParValue" Value="<%=iPartyVal%>">
<input type="hidden" name="hTransNo" Value="<%=nTransactionNo%>">
<%if sReqType="A" then%>
<input type="hidden" name="hVouCode" value="02">
<input type="hidden" name="hVouName" value="BA">
<%else%>
<input type="hidden" name="hVouCode" value="01">
<input type="hidden" name="hVouName" value="CA">
<%end if%>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">&nbsp;Regular Payment -
		<%if sReqType="A" then
			Response.Write "CHEQUE"
		else
			Response.Write "CASH"
		end if
		%>
		Generation for <%=sOrgname%>
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%" >
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
												<center>
                                                    <div align="left">
													<table cellpadding="0" cellspacing="0" width="90%">
														<tr>
															<td>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class='GroupTitleLeft' width="10">&nbsp;
                                                            </td>
															<td class='GroupTitle' width="110"><p align="center">Request Details
                                                            </td>
												</center>
															<td class='GroupTitleRight'><p align="left">&nbsp;
                                                            </td>
														</tr>
													</table>
                                                        </td>
														</tr>
														<tr>
															<td class=GroupTable>
												<center>
													<table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=MiddlePack colspan="3"> </td>
														</tr>
                                                        <tr>
								<td align="center">&nbsp;	</td>
								<td valign="top">
                                                   <table cellpadding="0" cellspacing="0" width="100%">
														<tr>
															<td class=FieldCell width="78"> Raised
                                                              By</td>
															<td class="FieldCell">
                                                            <span class="DataOnly"><%=sRequestBy%>&nbsp;</span></td>
															<td width="20" class="FieldCell">
                                                            </td>
															<td width="80" class="FieldCell">
                                                            Approved By</td>
															<td class="FieldCell">
                                                            <span class="DataOnly"><%=sApprovedBy%>&nbsp;</span></td>
														</tr>
														<tr>
															<td class=FieldCell width="78"> Pay To</td>
															<td class="FieldCell"><span class="DataOnly"><%=Trim(sPayTo)%>&nbsp;</span></td>
														</tr>

														<tr>
															<td class=MiddlePack colspan="5"> </td>
														</tr>

														<tr>
															<td class=FieldCell width="178" colspan="5">
															<DIV class=frmBody id=frm1 style="width: 385; height:50;">

                                    <table border="0" id="tblPayable0" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" width="30">S.No.</td>
                                        <td class="ExcelHeaderCell" align="center" >Reason</td>
                                        <td class="ExcelHeaderCell" align="center" >Amount To Pay</td>
                                            </tr>
                                            <%
												iSno=1
												objRs.MoveFirst
                                        	do while not objRs.EOF
                                        %>
                                        	    <tr>
													<td class="ExcelSerial" align="center" width="30"><%=iSno%></td>
													<td class="ExcelDisplayCell" align="Left"><%=objRs(0)%></td>
													<td class="ExcelDisplayCell" align="right"><%=FormatNumber(objRs(1),2,,,0)%></td>
                                            </tr>
                                        <%
												iSno=cint(iSno)+1
												objRs.MoveNext
											loop
											objRs.Close
										%>

                                                </table>
                                                												</div>

 </td>
														</tr>

													</table>
								</td>
								<td align="center">&nbsp;</td>
                                                        </tr>
												</center>
														<tr>
															<td class=MiddlePack colspan="3">
                                                   </td>
														</tr>
													</table>
                                                            </td>
														</tr>
													</table>
                                                        </div>
								</td>
								<td align="center">
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
                                     <table border="0" cellspacing="0" cellpadding="0">
                                       <tr>
											<td class="FieldCellSub"> <Input type="button" name="btnInsDet" Class="ActionButton2" Value="Instrument Details" onClick="PopInsDet()">&nbsp;</td >
                                       </tr>
                                       <tr>
											<td class="FieldCellSub" width="139">Inst No</td>
											<td>                 
												 <span id="spInsNo" class="DataOnly">-</span>
											</td >&nbsp;
                                            <td class="FieldCellSub" width="139">Inst Date</td>
											<td width="296">
												<span id="spInsDate" class="DataOnly">-</span>
                                            </td>
                                        </tr>
                                        <tr>
                                           <td class="FieldCellSub" width="139">Select Book</td>
                                           <td width="296">
                                               <select size="1" name="selBookId" class="FormElem">
                                                    <option value="Select">Select</option>
														<%
															if sReqType="A" then
																sQuery="select BookNumber,BookName,BookAccountHead from vwOrgBookNames where OUDefinitionID = '" & sorgID & "' and BookCode=02"
															else
																sQuery="select BookNumber,BookName,BookAccountHead from vwOrgBookNames where OUDefinitionID = '" & sorgID & "' and BookCode=01"
															end if
															with objRs1
																.CursorLocation = 3
																.CursorType = 3
																.Source = sQuery
																.ActiveConnection = con
																.Open
															end with
															set objRs1.ActiveConnection = nothing

														do while not objRs1.EOF
															Response.Write    "<option value="""&trim(objRs1(0))&"?"&trim(objRs1(2))&""">"&trim(objRs1(1))&"</option>"
															objRs1.MoveNext
														loop
														objRs1.Close
														%>
												</select>
											</td>
                                            <!--<td class="FieldCell" colspan="2" valign="bottom"><p align="center">Payment
                                             </td>-->
                                         </tr>
                                                         
                                          <tr>
                                                <td class="FieldCellSub" width="139">Party Name</td>
                                                 <!--select size="1" name="selAccType" class="FormElem" onChange="selAccountHead(this)"-->
                                                 <td class="FieldCell" width="296">
													<span id="ParNamID" class="DataOnly"><%If sParSubTypeName <> "" Then Response.Write sParSubTypeName Else Response.Write "-" End IF%></span>                                                    
										         </td>
                                                 <td class="FieldCell">Date</td>
                                                    <td class="FieldCell"> <p align="center">
													<% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
													%>
											</tr>
											<tr>
                                                 <td class="FieldCellSub" width="139">Pay to </td>
                                                 <td class="FieldCell" colspan="3"> <span id="ParNamID" class="DataOnly"><%=sParName%></span>  </td>
                                            </tr>
                                            <tr>
                                                 <td class="FieldCellSub" width="139" valign="top">Narration</td>
                                                 <td class="FieldCell" colspan="3"> <textarea rows="3" name="txtNarration" cols="70" class="FormElem"></textarea> </td>
                                            </tr>
                                            <tr>
                                                 <td class="FieldCellSub" width="139">Amount</td>
                                                 <td class="FieldCell" colspan="3"> <input type="text" name="txtAmount" value="<%=iAmount%>" size="15" style="text-align:right" class="Formelem"> </td>
                                            </tr>
                                     </table>
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
<DIV class=frmBody id="Disaddtional" style="height:1; visibility: hidden;">
<div id="DisCCANL" class=frmBody style="height:1; visibility: hidden;">
	<table cellpadding="0" cellspacing="0" >
		<tr>
			<td class=MiddlePack colspan="3"> </td>
		</tr>
		<tr>
			<td class=FieldCell>
				<DIV class=frmBody id="DisCost" style="width:280;height:100;">
					<table border="0" id="tblCost" cellspacing="1" class="ExcelTable">
						<tr>
							<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
								<td class="ExcelHeaderCell" align="center" width="150">Cost Center Head</td>
								<td class="ExcelHeaderCell" align="center">Ratio</td>
								<td class="ExcelHeaderCell" align="center">Amount</td>
						 </tr>
					</table>
				</div><!--End of CostCenter Display Division -->
			</td>
			<td class=ClearPixel width="5">	<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">                   </td>
			<td class=FieldCell>
				<DIV class=frmBody id="DisAnal" style="width:280; height:100;">

					<table border="0" id="tblAnal" cellspacing="1" class="ExcelTable">
						<tr>
								<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
								<td class="ExcelHeaderCell" align="center" width="150">Analytical Head</td>
								<td class="ExcelHeaderCell" align="center">Ratio</td>
								<td class="ExcelHeaderCell" align="center">Amount</td>
					    </tr>
					</table>
				</div>	<!--End of Analytical Display Division -->
			</td>
		</tr>
		<tr>
			<td class=MiddlePack  colspan="3"></td>
		</tr>
	</table>
</div> <!--End of CCANAL Display Division -->
	<DIV class=frmBody id="DisPayable" style="width: 555; visibility: hidden; height:1;">
		<table border="0" id="tblPayable" cellspacing="1" class="ExcelTable" width="555">
			<tr>
				<td class="ExcelHeaderCell" align="center" rowspan="2" width="10">S.No.</td>
				<td class="ExcelHeaderCell" align="center" width="300" colspan="3">Document</td>
				<td class="ExcelHeaderCell" align="center" width="250" colspan="3">Amount</td>
		    </tr>
		   <tr>
				<td class="ExcelHeaderCell" align="center">Number</td>
				<td class="ExcelHeaderCell" align="center">Date</td>
				<td class="ExcelHeaderCell" align="center">Type</td>
				<td class="ExcelHeaderCell" align="center">Amount</td>
				<td class="ExcelHeaderCell" align="center">Adjusted</td>
				<td class="ExcelHeaderCell" align="center">To adjust</td>
		   </tr>
		   <tr>
				<td class="ExcelSerial" align="center">1</td>
				<td class="ExcelDisplayCell">xNumber</td>
				<td class="ExcelDisplayCell" align="right"><p align="left">xDate</td>
				<td class="ExcelDisplayCell" align="right"><p align="left">xType</td>
				<td class="ExcelDisplayCell" align="right">xAmount
				</td>
				<td class="ExcelDisplayCell" align="right"><p align="right">xAdjusted
				</td>
				<td class="ExcelDisplayCell" align="right">xToAdjust</td>
			</tr>

		</table>
	</div>
</div><!--End of Addtional Details Display  -->
						</td>
								<td align="center" class="ClearPixel" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
												<input type="button" value="Create" name="B4" class="ActionButton" onClick="AddEntry()">
												 <input type="reset" value="Reset" name="B1" class="ActionButton" >
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
