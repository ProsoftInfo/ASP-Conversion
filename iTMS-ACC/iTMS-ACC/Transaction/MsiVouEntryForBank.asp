<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MsiVouEntryForBank.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Manohar Prabhu.R
	'Created On					:	Aug 20, 2004
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<%
dim sOrgId,sBookCode,sVouType,sOrgName,sBookName,oDOM
dim sQuery,objRs,bOtherUnits,iBookAccHead,dTransLimit
dim dOpeningBal,iTransNo,sPurVouNarr
Dim sPayTo,sVouNo,sVouDate,sVouAmount,sTransTy

dim sCreatedMiscPymtNo,sPaymentAgainst,sPayRefNo,iInvNo,sInvCode
dim iRcptCode,sGRNAgainstStr,sReceiptRouteStr,sReceiptCode,sItemType
Dim sReceiptRouting,iGRNNo,iRcptNo,iInspNo,sGrnCode,sGrnAgainst
Dim sRcptStr,sInspCode,saTemp,sUserID,sCrDrIndi
Dim sFinPeriod,sFromYr,sToYr,sTempYr,sChequeNo,sChequeDate
Dim sPartyType,sParSubType,sPartyCode,sFinFrom,sFinTo

sFinPeriod = Session("FinPeriod")
IF CStr(sFinPeriod) <> "" Then
	sTempYr = Split(sFinPeriod,":")
	sFromYr = sTempYr(0)
	sToYr = sTempYr(1)
	sFinFrom = "01/04/"&sFromYr
	sFinTo = "31/03/"&sToYr
End IF

set objRs = Server.CreateObject("ADODB.Recordset")

sOrgId=Request("OrgID")
sOrgName=Request("OrgName")
iTransNo = Request("TransNo")
sVouType="C"
sUserID = getUserID()

sQuery = "Select PayToRecdFrom,CreatedMiscPymtNo,Convert(Char,VoucherDate,103), "&_
		 "VoucherAmount,isNull(PaymentAgainst,''),isNull(ReferenceNo,0),CrDrIndication,TransactionType,PartyType,PartySubType,PartyCode,IsNull(ChequeNo,''),IsNull(ChequeDate,'') From Acc_T_MiscPymtRequestHeader  "&_
		 "Where MiscTransNo = "&iTransNo&" "



With objRs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = Con
	.Source = sQuery
	.Open
End With

Set objRs.ActiveConnection = Nothing

sCreatedMiscPymtNo = ""

IF Not objRs.EOF Then
	sPayTo = objRs(0)
	sVouNo = objRs(1)
	sVouDate = objRs(2)
	sVouAmount = objRs(3)

	
	sCreatedMiscPymtNo = objRs(1)
	sPaymentAgainst = UCase(objRs(4))
	sPayRefNo = objRs(5)
	sCrDrIndi = objRs(6)
	sTransTy = objRs(7)
	sPartyType = objRs(8)
	sParSubType = objRs(9)
	sPartyCode = objRs(10)
	sChequeNo = objRs(11)
	sChequeDate = objRs(12)
End IF
objRs.Close

IF CStr(sTransTy) = "CAP" or CStr(sTransTy)="BAP" Then
	sVouType = "C"
Else 'Cash Receipt Only in Sales Invoices.
	sVouType = "D"
End IF


IF CStr(sCrDrIndi) = "D" Then
	'****
	if trim(sPaymentAgainst) = "I" then
		sQuery = "Select distinct isNull(InvoiceNumber,0),(isnull(SuppInvoiceNo,InvoiceCode) + '--' + convert(varchar,InvoiceDate,103)) from RCV_T_InvoiceHeader where isNull(InvoiceNumber,0) = " & sPayRefNo & " "
		with objRs
			.ActiveConnection = con
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.Open
		end with

		set objRs.ActiveConnection = nothing

		iInvNo = 0
		sInvCode = ""
		if not objRs.EOF then
			iInvNo		= objRs(0)
			sInvCode	= objRs(1)
		end if 	'if not objRs.EOF then
		objRs.Close
	end if 'if trim(sPaymentAgainst) = "I" then

	if trim(sPaymentAgainst) = "R" then
		sQuery = " Select Distinct GRNNumber,isnull(ReceiptNumber,0) RcptNum ,isnull(InvoiceNumber,0),isnull(InspectionNumber,0)," &_
				" GRNCode = (Select (GRNCode + '--' + convert(varchar,GRNDate,103)) AS GRN from Rcv_T_GateReceiptHeader where GRNNumber = R.GRNNumber)," &_
				" GRNAgainst = (Select ReceiptAgainst from Rcv_T_GateReceiptHeader where GRNNumber = R.GRNNumber), " &_
				" RCPTCode = isnull((Select (ReceiptCode + '--' + convert(varchar,ReceiptDate,103) + '|' + isnull(ReceiptRouting,0)) AS Rcpt from RCV_T_ActualReceiptHeader Where ReceiptNumber = R.ReceiptNumber),0)," &_
				" InvCode = isnull((Select (isnull(SuppInvoiceNo,InvoiceCode) + '--' + convert(varchar,SuppInvoiceDate,103)) AS INV from RCV_T_InvoiceHeader where InvoiceNumber = R.InvoiceNumber),0)," &_
				" InspCode = isnull((Select Distinct InspectionNumber from RCV_T_PurchInspectionHeader where InspectionNumber = R.InspectionNumber),0), " &_
				" ItemType = isNull((Select Distinct IT.ItemtypeID from RCV_T_GRNItemDetails IT where  IT.GrnNumber = R.GRNNumber),'')" & _
				" from PUR_T_RefferenceNumberDet R where isnull(ReceiptNumber,0) = " & sPayRefNo & ""

		with objRs
			.ActiveConnection = con
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.Open
		end with

		set objRs.ActiveConnection = nothing


		if not objRs.EOF then
			iGRNNo = objRs(0)
			iRcptNo = objRs(1)
			iInspNo = objRs(3)
			sGrnCode = objRs(4)
			sGrnAgainst = objRs(5)
			sRcptStr = objRs(6)

			sInspCode = objRs(8)

			sItemType = objRs(9)

			sGRNAgainstStr = getReceiptType(sGrnAgainst)
			sReceiptCode = ""
			sReceiptRouting = ""

			if sRcptStr <> "" and sRcptStr <> "0" then
				saTemp = Split(sRcptStr,"|")
				sReceiptCode  = saTemp(0)
				sReceiptRouting = saTemp(1)
			End if

			If trim(sReceiptRouting) = "" or trim(sReceiptRouting) = "0" Then
				sReceiptRouteStr = "--"
			Else
				sReceiptRouteStr = getReceiptRoute(sReceiptRouting)
			End if
		end if 	'if not objRs.EOF then
		objRs.Close
	end if 'if trim(sPaymentAgainst) = "I" then
End IF



'****

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

oDOM.Load server.MapPath("../xmldata/CreditLimit.xml")
dTransLimit=CDbl(oDOM.documentElement.childNodes.item(0).text)

sQuery = "Select D.VoucherNarration,H.PayToRecdFrom From Acc_T_MiscPaymentReqDetails D,  "&_
		 "Acc_T_MiscPymtRequestHeader H Where D.MiscTransNo = "&iTransNo&" and H.MiscTransNo = "&iTransNo&" "
With objRs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = Con
	.Source = sQuery
	.Open
End With

Set objRs.ActiveConnection = Nothing
IF Not objRs.EOF Then
	sPurVouNarr = objRs(0)
	'sPayTo = objRs(1)
End IF
objRs.Close
sPurVouNarr = Trim(sPurVouNarr)


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/cancel.js"></SCRIPT>
<!--SCRIPT FOR COMMON VOUCHER FUNCTIONS -->
<script language="javascript" src="../scripts/VouTransactions.js"></script>
<!--SCRIPT FOR ADD ENTRY TABLE FUNCTIONS -->
<script language="javascript" src="../../scripts/ExcelFunctions.js"></script>
<!--XML ISLAND FOR VOUCHER DATA -->
<XML id="VoucherData"><voucher UnitNo="<%=sOrgId%>" UnitName="<%=sOrgName%>" BookNo="" BookName="" CRDR="<%=sVouType%>" VouDate="" BookAcchead="" Approver="" InstNo="" InstDate="" PayAt="" DrawnOn=""/></XML>
<!--XML ISLAND FOR ENTRY DATA -->
<XML id="EntryData"><Entry No="0" CRDR="0" Payto="" Amount="" AccUnit="" AccName="" TdsAmount="" TDSElgi="0" TdsPercentage="0" /></XML>
<!--XML ISLAND FOR TEMP DATA'S (PARTY TYPE /GLHEAD) -->
<XML id="OutData"><Root/></xml>
<XML id="AccHeadData">
<account/>
</XML>
<XML id="PartyData"><Root></Root></XML>
<XML id="TempXMLData"><Root></Root></XML>
<SCRIPT LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></SCRIPT>
<script language="vbscript">
DIM iEntryNo,VouRoot,EntryRoot,bVouFlag,bSavFlag
DIM iBookAcchead,dTransLimit,sTransFlag

iEntryNo=1
bVouFlag=false
bSavFlag=false
set VouRoot=VoucherData.documentElement
set EntryRoot=EntryData.documentElement

'iBookAcchead=<%=iBookAccHead%>
dTransLimit=<%=dTransLimit%>
sTransFlag="A"

FUNCTION CheckAccHead(nodRoot,sAccHead)
dim sExp
sExp="//AccHead[@No='"&sAccHead&"']"
set tempNode=nodRoot.selectNodes(sExp)

if tempNode.length > 0 then
	CheckAccHead=true
else
	CheckAccHead=false
end if
END FUNCTION
'---------------------END OF FUNCTION CHECKACCHEAD----------------------
FUNCTION popMonBalance(sValue)
dim saTemp,sbTemp
	IF document.formname.selBook.selectedIndex = 0 Then
		Msgbox "Select Book "
		document.formname.selBook.focus()
		Exit Function
	Else
		saTemp=Split(sValue,"~")
		sbTemp = Split(document.formname.selBook.value,"?")
		showModalDialog "PopMonBalance.asp?orgid="+saTemp(0)+"&Acchead="+sbTemp(1)+"&TillDate="+saTemp(2),"","dialogHeight:390px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No"
	End IF
END FUNCTION
'---------------------END OF FUNCTION POPMONBALANCE----------------------
FUNCTION popDayBalance(sValue)
dim saTemp,sbTemp
	IF document.formname.selBook.selectedIndex = 0 Then
		Msgbox "Select Book "
		document.formname.selBook.focus()
		Exit Function
	Else
		saTemp=Split(sValue,"~")
		sbTemp = Split(document.formname.selBook.value,"?")
		showModalDialog "PopDayBalance.asp?orgid="+saTemp(0)+"&Acchead="+sbTemp(1)+"&TillDate="+saTemp(2),"","dialogHeight:390px;dialogWidth:620px;center:Yes;help:No;resizable:No;status:No"
	End IF
END FUNCTION
'---------------------END OF FUNCTION POPDAYBALANCE----------------------
FUNCTION popAccHead()
	dim iHeadCount
	iUnitNo=document.formname.hOrgId.value
	iHeadCount=cint(document.formname.hHeadCount.value)
	iBkNo=document.formname.hBookcode.value
	document.formname.selAccHead.selectedIndex=0

	for iCounter=1 to iHeadCount
		document.formname.selAccHead.remove(1)
	next

	set objhttp = CreateObject("MSXML2.XMLHTTP")

	objhttp.Open "GET","XMLGetOrgFreqHeads.asp?BkCode=01&BkNo="&iBkNo&"&orgID=" & iUnitNo , false
	objhttp.send

	if objhttp.responseXML.xml <> "" then
		OutData.loadXML objhttp.responseXML.xml
		Set Root = OutData.documentElement
		iCounter=1

		For Each HeaderNode In Root.childNodes

			set oText1 = document.createElement("<Option>" )
				oText1.Text = HeaderNode.text
				oText1.Value = HeaderNode.Attributes.getNamedItem("optValue").Value

			document.formname.selAccHead.add oText1,iCounter
			iCounter=CDbl(iCounter)+1
		next
			document.formname.hHeadCount.value=CDbl(iCounter)-1
			iHeadCount=CDbl(iCounter)+1
	else
		document.formname.hHeadCount.value=0
		iHeadCount=2
	end if

	for iCounter=iHeadCount+1 to document.formname.selAccHead.length
		document.formname.selAccHead.remove(iHeadCount)
	next

	objhttp.Open "GET","XMLGetOrgParType.asp?orgID=" & iUnitNo , false
	objhttp.send

	if objhttp.responseXML.xml <> "" then
		OutData.loadXML objhttp.responseXML.xml
		Set Root = OutData.documentElement
		iCounter=document.formname.selAccHead.length
		For Each HeaderNode In Root.childNodes
			set oText1 = document.createElement("<Option>" )
				oText1.Text = HeaderNode.text
				oText1.Value = HeaderNode.Attributes.getNamedItem("ParType").Value
			document.formname.selAccHead.add oText1,iCounter
			iCounter=CDbl(iCounter)+1
		next
	end if
END FUNCTION
'---------------------END OF FUNCTION POPACCHEAD----------------------
FUNCTION selAccountHead(objAcc)

DIM sVouType,sOrgId,sTemp,iHeadCount,sDesc


'iHeadCount=cint(document.formname.hHeadCount.value)
iHeadCount = cint(1)

	if objAcc.selectedIndex = 1 then
		if document.formname.hOtherUnitFlag.value=1 then


			sOrgId=document.formname.hOrgId.value
			if objAcc.selectedIndex = 1 then
				'MsgBox "OK "
				showGLHead sOrgId
			else
				sTemp=objAcc.value& "?" & objAcc.options(objAcc.selectedIndex).text
				showPartyHead  sOrgId,sTemp,document.formname.hVouCRDR.value
			End if 'END OF SELECTED ACCOUNT HEAD TYPE IS GL(1) OR PARTY(>1)
			document.formname.txtNarration.focus

		else
			'MsgBox "Calling Else "
			sOrgId=document.formname.hOrgId.value
			if objAcc.selectedIndex <= iHeadCount then
					sTemp=Split(objAcc.value,"?")
					document.formname.hTdsElgi.value = sTemp(4)
					IF CStr(sTemp(4)) = "1" Then
						document.formname.txtTdsAmount.disabled = False
						document.formname.txtTdsper.disabled = False
					Else
						document.formname.txtTdsAmount.disabled = True
						document.formname.txtTdsper.disabled = True
					End IF

					'if CheckAccHead(VouRoot,trim(sTemp(0))) then
					'	MsgBox"Account Head already Exisit in Voucher"
					'	document.formname.selAccHead.selectedIndex=0
					'	exit function
					'end if

					sDesc=objAcc.options(objAcc.selectedIndex).text
					bVouFlag=true
					Set newElem = EntryData.createElement("AccHead")
						newElem.setAttribute "No", trim(sTemp(0))
						newElem.setAttribute "CostCenter", trim(sTemp(1))
						newElem.setAttribute "Analytical", trim(sTemp(2))
						newElem.setAttribute "Name", sDesc
						newElem.setAttribute "Type", "G"
						newElem.setAttribute "TransFalg", trim(sTemp(3))
	   					EntryRoot.appendChild newElem

						sTransFlag=trim(sTemp(3))
						window.spAccHead.innerHTML=sDesc&"&nbsp;"
						if iEntryNo=1 then
							document.formname.txtPayto.value=sDesc
						end if

					showCCAnal sOrgId,trim(sTemp(0)),trim(sTemp(1)),trim(sTemp(2))
				elseif objAcc.selectedIndex =iHeadCount+1 then
						showGLHead sOrgId
				else
					sTemp=objAcc.value& "?" & objAcc.options(objAcc.selectedIndex).text
					showPartyHead  sOrgId,sTemp,document.formname.hVouCRDR.value
				End if 'END OF SELECTED ACCOUNT HEAD TYPE IS GL(1) OR PARTY(>1)
			document.formname.txtNarration.focus
		end if	'END OF BOOK HAS OTHER UNIT TRANSCATION OR NOT CHECK
	End if 'END OF IF ANY ACCOUNT HEAD SELECTED CHECK
END FUNCTION
'---------------------END OF FUNCTION SELACCOUNTHEAD----------------------

FUNCTION showPartyHead(sOrgId,sPartyType,sVouType)
dim sPartyCode,bRecivable,bPayable,sVouDate
dim sDocNo,sInvNo,sInvDate,sAmtRec,sAmtRecd
dim nodAccHead,nodPayRec,nodCC,iSno,sRetValue,sTemp
Dim sParSubType,Objhttp,sRetVal2,sPartyName,sParCode,sParTy,sNarration
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth

set objhttp = CreateObject("Microsoft.XMLHTTP")

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

     '   OutValue = showModalDialog("PartySelection.asp?orgId="+sOrgId&"&Party="&sPartyType,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
     '   arrTemp = split(OutValue,":")

     '   while UBound(arrTemp) = 0
	 '       OutValue = showModalDialog("PartySelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	 '       arrTemp = split(OutValue,":")
     '   wend

     '   sRetValue = OutValue
     '   if UBound(arrTemp) <= 1 then exit function

     '   sTemp = Split(sRetValue,":")
     '   sParTy = sTemp(4)
     '   sParSubType = sTemp(3)
     '   sParCode = sTemp(1)
     '   sPartyName = sTemp(0)
     
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

if nodAccHead.hasChildNodes then
	'User Has Selected a Party
	clearXML()
	For Each HeaderNode In nodAccHead.childNodes
		sPartyCode=sPartyType&"?"& HeaderNode.Attributes.getNamedItem("No").Value

		bVouFlag=true
		HeaderNode.Attributes.getNamedItem("No").Value=sPartyCode
		bPayable=HeaderNode.Attributes.getNamedItem("Pay").Value
		bRecivable=HeaderNode.Attributes.getNamedItem("Rec").Value

		window.spAccHead.innerHTML=HeaderNode.Attributes.getNamedItem("Name").Value&"&nbsp;"
		if iEntryNo=1 then
			document.formname.txtPayto.value=HeaderNode.Attributes.getNamedItem("Name").Value
		end if

		EntryRoot.appendChild HeaderNode
		sTransFlag="A"
	next


	if (cint(bRecivable)>=1 and sVouType="D") or (cint(bPayable)>=1 and sVouType="C") then
	sVouDate=document.formname.ctlDate.getdate
	'If Selected Party Has Payable or Receiavable
		Set nodPayRec = showModalDialog("PayRecSelection.asp?VouDate="+sVouDate+"&orgId="+sOrgId+"&ParCode="+sPartyCode&"&Type="&sVouType,"","")
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
							sDocNo=nodCC.Attributes.getNamedItem("No").Value
							sInvNo=nodCC.Attributes.getNamedItem("InvNo").Value
							sInvDate=nodCC.Attributes.getNamedItem("InvDate").Value
							sTransAmount=nodCC.Attributes.getNamedItem("TransAmount").Value
							sAmtAdjusted=nodCC.Attributes.getNamedItem("AmtAdjusted").Value
							sAmtToAccount=nodCC.Attributes.getNamedItem("AmtToAccount").Value
							sNarration = sNarration &" "&sInvNo

							sTempTransAmt = CDbl(sTempTransAmt + sTransAmount)
							sTempAccAmt = CDbl(sTempAccAmt + sAmtToAccount)
							sTempAdjAmt = CDbl(sTempAdjAmt + sAmtAdjusted)


							set oRow = document.all.tblPayable.insertRow(iSno+1)
							InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
							InsertCell oRow,1,"",sInvNo,"ExcelDisplayCell","","",0,0,0,0,""
							InsertCell oRow,1,"",sInvDate,"ExcelDisplayCell","","",0,0,0,0,""
							InsertCell oRow,1,"",sTransAmount,"ExcelDisplayCell","Right","",0,0,0,0,""
							InsertCell oRow,1,"",sAmtAdjusted,"ExcelDisplayCell","Right","",0,0,0,0,""
							InsertCell oRow,1,"",sAmtToAccount,"ExcelDisplayCell","Right","",0,0,0,0,""
							InsertCell oRow,2,"txtDocAmount"&CStr(sDocNo),"0","ExcelInputCell","","",12,10,0,0,"style=""text-align:right"""
							iSno=iSno+1
						next
						sNarration = Mid(sNarration,2)
						dTotal = CDbl(sTempTransAmt - sTempAccAmt - sTempAdjAmt)
						document.formname.txtAmount.value = dTotal
						document.formname.txtNarration.value = sNarration

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
	window.spAccHead.innerHTML=""
	document.formname.selAccHead.selectedIndex=0
	'if iEntryNo=1 then	document.formname.txtPayTo.value=""
	'Set the Additional Layer Display Layer Hidden
	setPayableDisplay 0
	bVouFlag=false
end if 'End of Party Head Processing

set nodAccHead=nothing
set nodPayRec=nothing
set nodCC=nothing
END FUNCTION
'---------------------END OF FUNCTION showPartyHead--------------------------

function showGLHead(sOrgId)
dim iAccCode,bAnal,bCostCenter,arrTemp,sRetVal
dim nodAccHead,nodCCAnly,nodCC,nodANL,iSno
dim sCode,sDesc,dRatio,iBookNo,sTemp2,sTdsElgi,sTempVal
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth
iBookNo=document.formname.hBookcode.value
sRetVal = ""

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


   ' OutValue = showModalDialog("GLHeadSelection.asp?orgId="+sOrgId+"&BookId=01&BookNo="+iBookNo+"&AccHead="+cstr(iBookAcchead),"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
   ' arrTemp = split(OutValue,":")'

   ' while UBound(arrTemp) = 0
'	    OutValue = showModalDialog("GLHeadSelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
'	    arrTemp = split(OutValue,":")
 '   wend
'
 '   sRetVal = OutValue
  '  sTempVal = OutValue

   ' if UBound(arrTemp) <= 1 then exit function
    'sTemp2 = Split(sTempVal,":")
    'sTdsElgi = sTemp2(6)
    
    if OutValue.hasChildNodes() then
        for each ndEntry in OutValue.childNodes
            if ndEntry.nodeName="Entry" then
                sRetVal = ndEntry.getAttribute("RetField0")&":"&ndEntry.getAttribute("RetField1")&":"&ndEntry.getAttribute("RetField2")&":"&ndEntry.getAttribute("RetField3")&":"&ndEntry.getAttribute("RetField4")&":"&ndEntry.getAttribute("RetField5")&":"&ndEntry.getAttribute("RetField6")
                sTdsElgi = ndEntry.getAttribute("RetField6")
            end if
        next
    end if
    
    
document.formname.hTdsElgi.value = sTdsElgi
IF CStr(sTdsElgi) = "1" Then
	document.formname.txtTdsAmount.disabled = False
	document.formname.txtTdsper.disabled = False
Else
	document.formname.txtTdsAmount.disabled = True
	document.formname.txtTdsper.disabled = True
End IF
GetGlHeadXml(sRetVal)

Set nodAccHead = AccHeadData.documentElement


if nodAccHead.hasChildNodes then
	'User Has Selected a GL Account Head
	clearXML()
	For Each HeaderNode In nodAccHead.childNodes
		iAccCode=HeaderNode.Attributes.getNamedItem("No").Value

		'if CheckAccHead(VouRoot,iAccCode) then
		'		MsgBox"Account Head already Exisit in Voucher"
		'		document.formname.selAccHead.selectedIndex=0
		'		exit function
		'end if

		bVouFlag=true
		bAnal=HeaderNode.Attributes.getNamedItem("Analytical").Value
		bCostCenter=HeaderNode.Attributes.getNamedItem("CostCenter").Value
		sTransFlag=HeaderNode.Attributes.getNamedItem("TransFlag").Value

		window.spAccHead.innerHTML=HeaderNode.Attributes.getNamedItem("Name").Value&"&nbsp;"

		if iEntryNo=1 then
			document.formname.txtPayto.value=HeaderNode.Attributes.getNamedItem("Name").Value
		end if

		EntryRoot.appendChild HeaderNode
	next
	showCCAnal sOrgId,iAccCode,bCostCenter,bAnal
else
	'User canceled Account Head Selection
	'Set the Additional,Costcenter and Analy Layer Display Layer Hidden
	window.spAccHead.innerHTML=""
	document.formname.selAccHead.selectedIndex=0
	'if iEntryNo=1 then	document.formname.txtPayTo.value=""
	bVouFlag=false
	setADDDisplay 0
end if 'End of GL Head Processing
set nodAccHead=nothing
End function
'---------------------End Of Function showGLHead--------------------------

Function AddEntry(bFlag)
dim iCode,dRatio,dAmount,sGroupCode,sExp,CheckNode,sTemp

sExp = "//voucher"
Set CheckNode = VouRoot.selectNodes(sExp)

' New Validation for check blank data - included on 02/04/2004
if bFlag = "S" then
	if iEntryNo > 1 and document.formname.txtAmount.value = "0.00" then
		sTemp = Split(document.formname.selBook.value,"?")
		IF CheckNode.length <> 0 Then
			CheckNode.Item(0).Attributes.getNamedItem("BookNo").value = sTemp(0)
			CheckNode.Item(0).Attributes.getNamedItem("BookName").value = document.formname.selbook.options(document.formname.selBook.selectedIndex).text
			CheckNode.Item(0).Attributes.getNamedItem("BookAcchead").value = sTemp(1)

			IF document.formname.optApprove(0).checked = True Then
				CheckNode.Item(0).Attributes.getNamedItem("Approver").value = "Y"
			Else
				CheckNode.Item(0).Attributes.getNamedItem("Approver").value = "N"
			End IF
			checkNode.Item(0).Attributes.getNamedItem("InstNo").value = document.formname.txtInstNo.value
			checkNode.Item(0).Attributes.getNamedItem("InstDate").value = document.formname.ctlInsDate.getdate()
			checkNode.Item(0).Attributes.getNamedItem("PayAt").value = document.formname.txtPayableAt.value
			checkNode.Item(0).Attributes.getNamedItem("DrawnOn").value = document.formname.txtDrawnOn.value
		End IF
		SaveXML
		Exit Function
	end if

	IF document.formname.selBook.selectedIndex = 0 Then
		MsgBox "Select Book"
		document.formname.selBook.focus()
		Exit Function
	End IF

	sTemp = Split(document.formname.selBook.value,"?")

	IF CheckNode.length <> 0 Then
		CheckNode.Item(0).Attributes.getNamedItem("BookNo").value = sTemp(0)
		CheckNode.Item(0).Attributes.getNamedItem("BookName").value = document.formname.selbook.options(document.formname.selBook.selectedIndex).text
		CheckNode.Item(0).Attributes.getNamedItem("BookAcchead").value = sTemp(1)

		IF document.formname.optApprove(0).checked = True Then
			CheckNode.Item(0).Attributes.getNamedItem("Approver").value = "Y"
		Else
			CheckNode.Item(0).Attributes.getNamedItem("Approver").value = "N"
		End IF
        checkNode.Item(0).Attributes.getNamedItem("InstNo").value = document.formname.txtInstNo.value
		checkNode.Item(0).Attributes.getNamedItem("InstDate").value = document.formname.ctlInsDate.getdate()
		checkNode.Item(0).Attributes.getNamedItem("PayAt").value = document.formname.txtPayableAt.value
		checkNode.Item(0).Attributes.getNamedItem("DrawnOn").value = document.formname.txtDrawnOn.value
	End IF

	'SaveXML()

end if
' End of Validation


if bVouFlag then
	if not checkFileds then exit function
	bSavFlag=true

	VouRoot.Attributes.getNamedItem("VouDate").Value=document.formname.ctlDate.getdate
	EntryRoot.Attributes.getNamedItem("No").Value=iEntryNo

	if document.formname.selCRDR(0).checked then
		EntryRoot.Attributes.getNamedItem("CRDR").Value=document.formname.selCRDR(0).value
	else
		EntryRoot.Attributes.getNamedItem("CRDR").Value=document.formname.selCRDR(1).value
	end if

	EntryRoot.Attributes.getNamedItem("Payto").Value=document.formname.txtPayTo.value
	EntryRoot.Attributes.getNamedItem("Amount").Value=document.formname.txtAmount.value
	EntryRoot.Attributes.getNamedItem("TdsAmount").Value=document.formname.txtTdsAmount.value
	EntryRoot.Attributes.getNamedItem("TDSElgi").Value=document.formname.hTdsElgi.value
	EntryRoot.Attributes.getNamedItem("TdsPercentage").Value=document.formname.txtTdsper.value
	'document.formname.hTdsElgi.value = "0"

	if document.formname.hOtherUnitFlag.value=1 then
		EntryRoot.Attributes.getNamedItem("AccUnit").Value=document.formname.hOrgId.value
		EntryRoot.Attributes.getNamedItem("AccName").Value=document.formname.hOrgName.value

	else
		EntryRoot.Attributes.getNamedItem("AccUnit").Value=document.formname.hOrgId.value
		EntryRoot.Attributes.getNamedItem("AccName").Value=document.formname.hOrgName.value
	end if

	Set newElem = EntryData.createElement("Narration")
	newElem.text= document.formname.txtNarration.value
	EntryRoot.appendChild newElem

	for each HeaderNode in EntryRoot.childNodes
		if 	HeaderNode.nodeName="CostCenter" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value
				dRatio=eval("document.formname.txtCCRatio"&iCode).value
				dAmount=eval("document.formname.txtCCAmount"&iCode).value
				nodANL.Attributes.getNamedItem("Ratio").Value=dRatio
				nodANL.Attributes.getNamedItem("Amount").Value=dAmount
			next
		end if 'End of Check for Cost Center Node
		if 	HeaderNode.nodeName="Analytical" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value
				sGroupCode=nodANL.Attributes.getNamedItem("GroupCode").Value

				dRatio=eval("document.formname.txtANALRatio"&iCode&"Z"&sGroupCode).value
				dAmount=eval("document.formname.txtANALAmount"&iCode&"Z"&sGroupCode).value
				nodANL.Attributes.getNamedItem("Ratio").Value=dRatio
				nodANL.Attributes.getNamedItem("Amount").Value=dAmount
			next
		end if 'End of Check for Analytical Node
		if 	HeaderNode.nodeName="PayRec" then
			for each  nodANL in HeaderNode.childNodes
				iCode=nodANL.Attributes.getNamedItem("No").Value
				dAmount=eval("document.formname.txtDocAmount"&iCode).value
				nodANL.Attributes.getNamedItem("AmtToAdjust").Value=dAmount
			next
		end if 'End of Check for Analytical Node
	next
	VouRoot.appendChild EntryRoot

	if bFlag="A" then
		DisplayVoucher
		iEntryNo=iEntryNo+1
		bVouFlag=false
		sTransFlag="A"
		clearXML()
		document.formname.txtTdsAmount.value = "0.00"
		document.formname.txtTdsper.value = "0.00"
		setADDDisplay 0
		setPayableDisplay 0

		window.spAccHead.innerHTML=""
		window.spEntryNo.innerHTML=iEntryNo

		if iEntryNo>1 then
			document.formname.txtPayTo.readOnly=true
		end if

		document.formname.selAccHead.selectedIndex=0
		document.formname.selCRDR(0).disabled=false
		document.formname.selCRDR(1).disabled=false

		document.formname.txtAmount.value="0.00"
		document.formname.txtNarration.value=""
	else
		IF document.formname.selBook.selectedIndex = 0 Then
			MsgBox "Select Book "
			document.formname.selBook.focus()
			Exit Function
		End IF
		'MsgBox "OK "
		SaveXML
		Exit Function
	end if
end if
END FUNCTION

'---------------------END OF FUNCTION ADDENTRY----------------------------
FUNCTION DisplayVoucher()
dim sNarration,sAccount,sAddtional,iSno,sAmount
dim dTotal,sAccUnit,sTotalCRDR,sTdsAmount,sTdsPer

window.DisVoucher.style.height="200px"
window.DisVoucher.style.visibility="visible"
ClearTable "tblVoucher",1,1
dTotal=0

'alert(VouRoot.xml)

For Each EntryNode in VouRoot.childNodes
	iSno=EntryNode.Attributes.Item(0).nodeValue
	sAmount=EntryNode.Attributes.Item(3).nodeValue
	sAmount=FormatNumber(CDbl(sAmount),2,,,0)
	sTdsAmount = EntryNode.Attributes.Item(6).nodeValue
	sTdsAmount = FormatNumber(CDbl(sTdsAmount),2,,,0)
	sTdsPer = EntryNode.Attributes.Item(8).nodeValue
	sTdsPer = FormatNumber(CDbl(sTdsPer),2,,,0)

	if EntryNode.Attributes.Item(1).nodeValue ="C" then
		dTotal=dTotal-CDbl(sAmount)
	else
		dTotal=dTotal+CDbl(sAmount)
	end if

	sAccUnit=EntryNode.Attributes.Item(5).nodeValue
	sAmount=sAmount&"&nbsp;"&EntryNode.Attributes.Item(1).nodeValue&"r"
	sAddtional=""
	For Each HeaderNode in EntryNode.childNodes
		if HeaderNode.nodeName="AccHead" then
				if HeaderNode.Attributes.Item(4).nodeValue="P" then
					sAccount=HeaderNode.Attributes.Item(3).nodeValue
				else
					sAccount= HeaderNode.Attributes.Item(3).nodeValue
				end if
		end if 'End of Check for Account head Node
		if 	HeaderNode.nodeName="Narration" then
				sNarration=HeaderNode.text
		end if 'End of Check for Narration Node
		if 	HeaderNode.nodeName="CostCenter" then
				for each  nodANL in HeaderNode.childNodes
					sAddtional=sAddtional&nodANL.Attributes.Item(2).nodeValue&"-"
					sAddtional=sAddtional&nodANL.Attributes.Item(3).nodeValue &"%&nbsp;"
					sAddtional=sAddtional&nodANL.Attributes.Item(4).nodeValue&"<br>"
				next
		end if 'End of Check for Cost Center Node
		if 	HeaderNode.nodeName="Analytical" and HeaderNode.hasChildnodes then
				sAddtional=sAddtional&"---------------------------  <br>"
				for each  nodANL in HeaderNode.childNodes
					sAddtional=sAddtional&nodANL.Attributes.Item(2).nodeValue&"-"
					sAddtional=sAddtional&nodANL.Attributes.Item(3).nodeValue &"%&nbsp;"
					sAddtional=sAddtional&nodANL.Attributes.Item(4).nodeValue&"<br>"
				next
		end if 'End of Check for Analytical Node
		if 	HeaderNode.nodeName="PayRec" then
				for each  nodANL in HeaderNode.childNodes
					sAddtional=sAddtional&nodANL.Attributes.Item(1).nodeValue&":"
					sAddtional=sAddtional&nodANL.Attributes.Item(2).nodeValue &"-&nbsp;"
					sAddtional=sAddtional&nodANL.Attributes.Item(5).nodeValue&"<br>"
				next
		end if 'End of Check for Analytical Node
	next 'End of Entry Node Loop

	set oRow = document.all.tblVoucher.insertRow(iSno)
	InsertCell oRow,1,"",iSno,"ExcelSerial","Center","top",0,0,0,0,""
	InsertCell oRow,1,"",sAccUnit,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sAccount,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sNarration,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sAmount,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",sAddtional,"ExcelDisplayCell","left","top",0,0,0,0,""
	InsertCell oRow,1,"",sTdsAmount,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"",sTdsPer,"ExcelDisplayCell","right","top",0,0,0,0,""

next'End of Voucher Node Loop
	if dTotal < 0 then
		sTotalCRDR="&nbsp;Cr"
		dTotal=CDbl(dTotal)*-1
	else
		sTotalCRDR="&nbsp;Dr"
	end if

	dTotal="Rs. &nbsp;"&FormatNumber(dTotal,2,,,0)

	set oRow = document.all.tblVoucher.insertRow(iSno+1)
	InsertCell oRow,1,"","<b>Total</b>","ExcelDisplayCell","right","top",0,0,4,0,""
	InsertCell oRow,1,"",CStr(dTotal)&sTotalCRDR ,"ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
	InsertCell oRow,1,"","","ExcelDisplayCell","right","top",0,0,0,0,""
END FUNCTION
'---------------------END OF FUNCTION DISPLAYVOUCHER----------------------
FUNCTION SaveXML()
	Dim sExp,TempNode,sVouName
	sExp = "//voucher"
	sVouName = document.formname.hVouName.value

	Set TempNode = VoucherData.selectNodes(sExp)
	IF TempNode.length <> 0 Then
		IF document.formname.optApprove(0).checked = True Then
			TempNode.Item(0).Attributes.getNamedItem("Approver").value = "Y"
		Else
			TempNode.Item(0).Attributes.getNamedItem("Approver").value = "N"
		End IF
	End IF
	
	

	if bSavFlag then

		set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.Open "POST","XMLSave.asp?Name=Voucher Entry&Mod="&sVouName, false
		objhttp.send VoucherData.XMLDocument
		if objhttp.responseText <> "" then
			Msgbox(objhttp.responseText)
		else
			document.formname.B12.disabled = True
			document.formname.submit()
		end if
	end if
END FUNCTION
'---------------------END OF FUNCTION SAVEXML-----------------------------
FUNCTION  checkFileds()

	if trim(document.formname.txtNarration.value)="" then
		Msgbox("Enter Narration")
		document.formname.txtNarration.select
		checkFileds=false
		exit Function
	end if
	if Trim(document.formname.txtInstNo.value)="" then
	    MsgBox("Enter Instrument No")
	    document.formname.txtInstNo.focus
	    checkFileds = false
	    exit function
	end if 
	if trim(document.formname.txtPayableAt.value)="" then
	    MsgBox("Enter PayableAt")
	    document.formname.txtPayableAt.focus
	    checkFileds = false
	    exit function 
	end if 
	if Trim(document.formname.txtDrawnOn.value)="" then
	    MsgBox("Enter DrawnOn")
	    document.formname.txtDrawnOn.focus
	    checkFileds = false
	    exit function
	end if
	if ValidateAmount(document.formname.txtAmount.value)=false then
		document.formname.txtAmount.select
		checkFileds=false
		exit Function
	end if
	if CDbl(document.formname.txtAmount.value) > CDbl(dTransLimit) then
		select case sTransFlag
			case "W"
					MsgBox "Amount is greater than the amount limit",,"Warning"
			case "R"
					MsgBox "Amount should be less than "&dTransLimit
					checkFileds=false
					exit Function
		end select
	end if
	for each HeaderNode in EntryRoot.childNodes
		if HeaderNode.nodeName="PayRec" then
			dAmount=CDbl(document.formname.txtAmount.value)
			dTotalAmtAdjust=0
			iCounter=1
				for each  nodANL in HeaderNode.childNodes
					iCode=nodANL.Attributes.getNamedItem("No").Value
					dTransAmount=nodANL.Attributes.getNamedItem("TransAmount").Value
					dAmtAdjusted=nodANL.Attributes.getNamedItem("AmtAdjusted").Value
					dAmtToAccount=nodANL.Attributes.getNamedItem("AmtToAccount").Value

					dAmtAdjust=CDbl(dTransAmount)-(CDbl(dAmtAdjusted)+CDbl(dAmtToAccount))
					dTotal=eval("document.formname.txtDocAmount"&iCode).value
					if  CDbl(dTotal)>CDbl(dAmtAdjust) then
						MsgBox """To Adjust Amount"" should be less than ""Document Amount-(Adjusted +To Account)"""
						eval("document.formname.txtDocAmount"&iCode).focus
						checkFileds=false
						exit Function
					else
						dTotalAmtAdjust=CDbl(dTotalAmtAdjust)+CDbl(dTotal)
					end if
				next
				if  CDbl(dTotalAmtAdjust)>CDbl(dAmount) then
					MsgBox "Total of ""To Adjust Amount"" should be less than ""Voucher Amount"""
					checkFileds=false
					exit Function
				end if
		end if 'End of Check for PayRec Node
	next
	checkFileds=true
END FUNCTION

'---------------------END OF FUNCTION CHECKFILEDS-------------------------
FUNCTION CancelAction(sPage)
	document.formname.action=sPage
	document.formname.submit
END FUNCTION
'---------------------END OF FUNCTION ACTIONCANCEL----------------------------

Function SelMisParty()
	Dim arrTemp,sRetValue,sParCode,sPartyName,sTemp

	'Only for First Entry The Pay To Can able to get Changed.
	IF (document.formname.txtPayTo.readOnly) Then
		Exit Function
	End IF

	OutValue = showModalDialog("MisPartySelection.asp","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	IF CStr(OutValue) = "AN" Then
		AddNewParty()
		Exit Function
	End IF
	arrTemp = split(OutValue,":")


	while UBound(arrTemp) = 0
		OutValue = showModalDialog("MisPartySelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
	wend

	sRetValue = OutValue
	'MsgBox sRetValue
	if UBound(arrTemp) <= 1 then exit function

	sTemp = Split(sRetValue,":")
	document.formname.txtPayTo.value = sTemp(0)
	'sParTy = sTemp(4)
	'sParSubType = sTemp(3)
	'sParCode = sTemp(1)
	'sPartyName = sTemp(0)
End Function

Function AddNewParty()
	OutValue = showModalDialog("MisParCreate.asp?"&OutValue,"","dialogHeight:495px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	'MsgBox OutValue
	document.formname.txtPayTo.value = OutValue
End Function
'*------------------------------------------------------------------------------------------
function showReceiptpopup(iGrnDt,iRcptCode,sGRNAgainstStr,sReceiptRouteStr,sInvCode,sReceiptCode,sItemType)
	showmodalDialog "../../Purchase/Transaction/RepActualReceiptDetailspopup.asp?ItemType="+cstr(sItemType)+"&iGrnDt="+cstr(iGrnDt)+"&iRcptCode="+cstr(iRcptCode)+"&sGRNAgainstStr="+cstr(sGRNAgainstStr)+"&sReceiptRouteStr="+cstr(sReceiptRouteStr)+"&sInvCode="+cstr(sInvCode)+"&sReceiptCode="+cstr(sReceiptCode)+"","","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
	'window.open "RepActualReceiptDetailspopup.asp?iRcptCode="+cstr(iRcptCode)+"&sGRNAgainstStr="+cstr(sGRNAgainstStr)+"&sReceiptRouteStr="+cstr(sReceiptRouteStr)+"&sInvCode="+cstr(sInvCode)+"&sReceiptCode="+cstr(sReceiptCode)+"",""
end function
'*------------------------------------------------------------------------------------------
function ViewInvoiceDetailspopup(iInvNo,sInvCode)
	'window.open "RepPurInvoiceDetailspopup.asp?iInvNo="+cstr(iInvNo)+"&sInvCode="+cstr(sInvCode)+"","A",""
	showmodalDialog "../../Purchase/Transaction/RepPurInvoiceDetailspopup.asp?iInvNo="+cstr(iInvNo)+"&sInvCode="+cstr(sInvCode)+"","A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
end function
'*------------------------------------------------------------------------------------------
Function ResetList(sObj)
	IF sObj.Value = "Y" Then
		document.formname.selUserId.disabled = False
	Else
		document.formname.selUserId.disabled = True
	End IF
End Function

'=======================================================================================================
Function Init()
    sFinFrom = document.formname.hFinFrom.value
    sFinTo = document.formname.hFinTo.value
    if DateDiff("d",sFinTo,date)>0 then
        document.formname.ctlDate.setMinDate= sFinFrom
        document.formname.ctlDate.setMaxDate= sFinTo
        document.formname.ctlDate.setDate = sFinTo
    else
        document.formname.ctlDate.setMinDate= sFinFrom
        document.formname.ctlDate.setMaxDate= Date
        document.formname.ctlDate.setDate = date
    end if 
    sCheqNo = document.formname.hChequeNo.value
    sCheqDate = document.formname.hChequeDate.value
    document.formname.txtInstNo.value = sCheqNo
    document.formname.ctlInsDate.setDate = sCheqDate
    
End Function
'******************************************
Function setParty()
Dim sParType,sParSubType,sParCode
sParTy = document.formname.hParType.value
sParSubType = document.formname.hParSubType.value
sParCode = document.formname.hParCode.value
sOrgId = document.formname.hOrgId.value

set objAccHead = eval("document.formname.selAccHead")
For iCnt = 0 to objAccHead.length-1
    if Trim(sParTy&"?"&sParSubType) = trim(objAccHead(iCnt).value) then
    sPartyType = objAccHead(iCnt).value&"?"&objAccHead(iCnt).text
        objAccHead.selectedIndex = iCnt
        exit for
    end if 
Next

set objhttp = CreateObject("Microsoft.XMLHTTP")
objhttp.open "GET","../../Include/GetPartyName.asp?ParCode="&sParCode,false
objhttp.send
if Trim(objhttp.responseText)<>"" then
    sPartyName= objhttp.responseText
end if 

objhttp.Open "GET","XMLGetPayRecCount.asp?orgID="&sOrgId&"&ParSubType="&sParSubType&"&ParType=" & sParTy&"&PartyCode="&sParCode , false
objhttp.send

IF objhttp.responseText <> "" Then
	sRetVal2 = objhttp.responseText
	GetPartyHeadXml sParCode,sPartyName,sRetVal2
End IF
Set nodAccHead = AccHeadData.documentElement

if nodAccHead.hasChildNodes then
	'User Has Selected a Party
	clearXML()
	For Each HeaderNode In nodAccHead.childNodes
	    sPartyCode=sPartyType&"?"& HeaderNode.Attributes.getNamedItem("No").Value
		
		bVouFlag=true
		HeaderNode.Attributes.getNamedItem("No").Value=sPartyCode
		bPayable=HeaderNode.Attributes.getNamedItem("Pay").Value
		bRecivable=HeaderNode.Attributes.getNamedItem("Rec").Value

		window.spAccHead.innerHTML=HeaderNode.Attributes.getNamedItem("Name").Value&"&nbsp;"
		if iEntryNo=1 then
			If document.formname.txtPayto.value = "" Then
				document.formname.txtPayto.value=HeaderNode.Attributes.getNamedItem("Name").Value
			End If
		end if

		EntryRoot.appendChild HeaderNode
		sTransFlag="A"
	next
	
	
	if (cint(bRecivable)>=1 and sVouType="D") or (cint(bPayable)>=1 and sVouType="C") then
	sVouDate=document.formname.ctlDate.getdate
	'If Selected Party Has Payable or Receiavable
		Set nodPayRec = showModalDialog("PayRecSelection.asp?VouDate="+sVouDate+"&orgId="+sOrgId+"&ParCode="+sPartyCode&"&Type="&sVouType,"","")
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
							sDocNo=nodCC.Attributes.getNamedItem("No").Value
							sInvNo=nodCC.Attributes.getNamedItem("InvNo").Value
							sInvDate=nodCC.Attributes.getNamedItem("InvDate").Value
							sTransAmount=nodCC.Attributes.getNamedItem("TransAmount").Value
							sAmtAdjusted=nodCC.Attributes.getNamedItem("AmtAdjusted").Value
							sAmtToAccount=nodCC.Attributes.getNamedItem("AmtToAccount").Value
							sNarration = sNarration &" "&sInvNo
							
							sTempTransAmt = CDbl(sTempTransAmt + sTransAmount)
							sTempAccAmt = CDbl(sTempAccAmt + sAmtToAccount)
							sTempAdjAmt = CDbl(sTempAdjAmt + sAmtAdjusted)
							

							set oRow = document.all.tblPayable.insertRow(iSno+1)
							InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""
							InsertCell oRow,1,"",sInvNo,"ExcelDisplayCell","","",0,0,0,0,""
							InsertCell oRow,1,"",sInvDate,"ExcelDisplayCell","","",0,0,0,0,""
							InsertCell oRow,1,"",sTransAmount,"ExcelDisplayCell","Right","",0,0,0,0,""
							InsertCell oRow,1,"",sAmtAdjusted,"ExcelDisplayCell","Right","",0,0,0,0,""
							InsertCell oRow,1,"",sAmtToAccount,"ExcelDisplayCell","Right","",0,0,0,0,""
							InsertCell oRow,2,"txtDocAmount"&CStr(sDocNo),"0","ExcelInputCell","","",12,10,0,0,"style=""text-align:right"""
							iSno=iSno+1
						next
						sNarration = Mid(sNarration,2)
						dTotal = CDbl(sTempTransAmt - sTempAccAmt - sTempAdjAmt)
						document.formname.txtAmount.value = dTotal
						document.formname.txtNarration.value = sNarration
						
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
	window.spAccHead.innerHTML=""
	document.formname.selAccHead.selectedIndex=0
	'if iEntryNo=1 then	document.formname.txtPayTo.value=""
	'Set the Additional Layer Display Layer Hidden
	setPayableDisplay 0
	bVouFlag=false
end if 'End of Party Head Processing

set nodAccHead=nothing
set nodPayRec=nothing
set nodCC=nothing



End Function
</script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init();setparty();">
<form method="POST" name="formname" action="VouMsiGenerate.asp?CallFrom=Bank">
<input type="hidden" name="hVouCode" value="02">
<input type="hidden" name="hVouCRDR" value="<%=sVouType%>">
<input type="hidden" name="hVouName" value="BA">
<input type="hidden" name="hOrgId" value="<%=sOrgId%>">
<input type="hidden" name="hOrgName" value="<%=sOrgName%>">
<input type="hidden" name="hBookcode" value="<%=sBookCode%>">
<input type="hidden" name="hOtherUnitFlag" value="1">
<input type="hidden" name="hTransNo" value="<%=iTransNo%>">
<input type="hidden" name="hEntryNo" value="0">
<input type="hidden" name="hTdsElgi" value="0">
<input type="hidden" name="hFromYr" value="<%=sFromYr%>">
<input type="hidden" name="hToYr" value="<%=sToYr%>">
<input type="hidden" name="hFinFrom" value="<%=sFinFrom%>" />
<input type="hidden" name="hFinTo" value="<%=sFinTo%>" />
<input type="hidden" name="hParType" value="<%=sPartyType%>" />
<input type="hidden" name="hParSubType" value="<%=sParSubType%>" />
<input type="hidden" name="hParCode" value="<%=sPartyCode%>" />
<input type="hidden" name="hChequeNo" value="<%=sChequeNo%>" />
<input type="hidden" name="hChequeDate" value="<%=sChequeDate%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		 Miscellaneous Voucher
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
								<td class="TabCell" valign="bottom" width="105">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										<tr>
											<td align="center">Book Selection
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCurrentCell" valign="bottom" align="center" width="110">
									<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable">
										<tr>
											<td align="center">Voucher Details
											</td>
										</tr>
									</table>
								</td>
								<td class="TabCell" valign="bottom" align="center" width="70">
								  <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
								  	<tr>
								  		<td align="center">Voucher</td>
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
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <!--tr>
                            <td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr-->
                            <!--tr>
                            <td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%" align="left">
                            <table border="0" cellspacing="0" cellpadding="0" class="ToolBarTable">
                        <tr>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                     <a href="javascript:popMonBalance('<%=sOrgId&"~"&iBookAccHead&"~"&Year(date)&Month(date)%>')"><span style="cursor: hand" Title="Month wise Balance" >
                    <p align="center"><font face="Webdings" size="5">?</font>
                    </span></a>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
                    <p align="center">
                    <a href="javascript:popDayBalance('<%=sOrgId&"~"&iBookAccHead&"~"&FormatDate(date)%>')">
                    <span style="cursor: hand" Title="Daywise Balance"><font face="Webdings" size="5">?</font>
                    </span></a>
                    </p>
                    </td>
                    <td class="ToolBarCell" width="40" onClick="toolClick(this)" onMouseOver="toolrollover(this)" onMouseOut="toolrollout(this)" >
						<p align="center">
							<span style="cursor: hand" Title="Voucher History">
								<font face="Webdings" size="5">?</font>
							</span>
						</p>
                    </td>
                        </tr>
                            </table>
                            </td>
                            <td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            </tr-->
							<tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
                                                            <table border="0" cellspacing="0" cellpadding="0">

                                                        <tr>
                                                    <td class="FieldCell" colspan="2">
                                                      <table border="0" width="100%" cellspacing="1" class="TableOutlineOnly">
                                                        <tr>
                                                          <td class="MiddlePack" colspan="6"></td>
                                                        </tr>
                                                        <tr>
                                                          <td class="FieldCellSub" width="90">Voucher
                                                            Date</td>
                                                          <td class="FieldCellSub" width="125">
                                                          <% ' Function Call to Insert Date Picker
														Response.Write InsertDatePicker("ctlDate")
													%>

                                                          </td>
                                                          <td class="FieldCellSub" width="180">Entry
                                                      Number</td>
                                                          <td class="FieldCellSub"><span class="DataOnly" id="spEntryNo">1&nbsp;</span></td>
                                                          <td class="FieldCellSub" width="100">
                                                    </td>
                                                          <td class="FieldCellSub">
   <span class="DataOnly">
                                                            <%
                                                             'dOpeningBal =GetDayOpening(sOrgId,iBookAccHead,FormatDate(date+1))
                                                             'dOpeningBal=FormatNumber(dOpeningBal,2,,,0)
                                                             'if dOpeningBal<0 then
															'	Response.Write dOpeningBal*-1 &"&nbsp;Cr"
															 'else
															'	Response.Write dOpeningBal &"&nbsp;Dr"
															 'end if
                                                            %></span>
   &nbsp;</td>
                                                        </tr>
                                                        <tr>
                                                          <td class="FieldCellSub" width="90">Entry
                                                            Type</td>
                                                          <td class="FieldCellSub" colspan="2">
											<%if sVouType="C" then%>
                                                            <input type=radio name="selCRDR" value="C" disabled>Receipts
                                                            <input type=radio name="selCRDR" value="D" checked>Payments
                                                            <%else%>
                                                            <input type=radio name="selCRDR" value="C" checked>Receipts
                                                            <input type=radio name="selCRDR" value="D" disabled >Payments&nbsp;
											<%end if%>
                                                            </td>
                                                          <td class="FieldCellSub" width="100">
                                                      &nbsp;</td>

                                                        </tr>
                                                        <tr>
                                                          <td class="FieldCellSub" width="90">Select
                                                            Book</td>
                                                          <td class="FieldCellSub" colspan="3">
                                                          <select size="1" name="selBook" class="FormElem">
																<option value="S">Select Book</option>
															<%
																sQuery = "select BookNumber,BookName,isnull(BookAccountHead,0),OtherUnitTransaction from "&_
																		 "vwOrgBookNames where OUDefinitionID = '" & sorgID & "' and BookCode='02' "&_
																		 "and BookAccountHead is not null "

																with objRs
																	.CursorLocation = 3
																	.CursorType = 3
																	.Source = sQuery
																	.ActiveConnection = con
																	.Open
																end with
																set objRs.ActiveConnection = nothing
																while not objRs.EOF
															%>
																	<Option value="<%=objRs(0)%>?<%=objRs(2)%>"><%=objRs(1)%></Option>
															<%
																objRs.MoveNext
																Wend
																objRs.Close
															%>
                                                          </td>
                                                        </tr>
                                                         <tr >
                                                          <td class="FieldCellSub" width="190">Instrument Type</td>

                                                          <td class="FieldCell" colspan="5">&nbsp;<Input type="radio" name="optInsType" value="Cheque" checked class="FormElem"> Cheque
                                                          &nbsp;<input type="radio" name="optInsType" value="Demand Draft" class="FormElem"> Demand Draft
                                                          &nbsp;<input type="radio" name="optInsType" value="Bankers Cheque" class="FormElem">
                                                        Bankers
                                                        Cheque&nbsp;&nbsp; <input type="radio" name="optInsType" value="Telegraphic Transfer" class="FormElem">
                                                        Telegraphic Transfer </td>
                                                        </tr>

                                                        <tr>
                                                          <td class="FieldCellSub" width="120">Instrument Number</td>
                                                          <td class="FieldCell" colspan="3"><input type="text" name="txtInstNo" size="20" class="Formelem">
                                                          </td>
                                                          <td class="FieldCellSub" width="120">Payable at</td>
                                                          <td class="FieldCell" colspan="3"><input type="text" name="txtPayableAt" size="20" class="Formelem">
                                                          </td>
                                                       </tr>
                                                        <tr>
                                                          <td class="FieldCellSub" width="120">Instrument Date</td>
                                                          <td class="FieldCell" colspan="3">

	<OBJECT ID="ctlInsDate"
		CLASSID="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"
		codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="formelem" VIEWASTEXT>
    <param name="_ExtentX" value="2355">
    <param name="_ExtentY" value="529">
	</OBJECT>

                                                          </td>
                                                          <td class="FieldCellSub" width="120">Drawn On</td>
                                                          <td class="FieldCell" colspan="3"><input type="text" name="txtDrawnOn" size="20" class="Formelem">
                                                          </td>
                                                       </tr>

                                                        <tr>
                                                          <td class="MiddlePack" colspan="6"></td>
                                                        </tr>
                                                      </table>
                                                    </td>
                                                        </tr>

                                                        <tr>
                                                    <td class="MiddlePack" colspan="2"></td>
                                                        </tr>
													<tr>
														<td class="FieldCellSub" width="133">Reference No</td>
														<td class="FieldCell" width="591">
														<% Response.Write sCreatedMiscPymtNo
															IF CStr(sCrDrIndi) = "D" Then
																if trim(sPaymentAgainst) = "I" then
														%>
																	<a class="ExcelDisplayLink" href="javascript:void(0)" onClick="vbscript:ViewInvoiceDetailspopup <%=iInvNo%>,'<%=sInvCode%>'" >
																		<img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Invoice">
																	</a>
														<%		End if
															End IF
														%>


															<%	IF CStr(sCrDrIndi) = "D" Then
																	if trim(sPaymentAgainst) = "R" then
															%>
																	<a class="ExcelDisplayLink" href="javascript:void(0)" onClick="vbscript:showReceiptpopup '<%=sGrnCode%>','<%=iRcptNo%>','<%=sGRNAgainstStr%>','<%=sReceiptRouteStr%>','','<%=sReceiptCode%>','<%=sItemType%>'	" >
																		<img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Receipt">
																	</a>
															<%
																	End if
																End IF
															%>
														</td>
				                                     </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Pay For</td>
                                                    <td class="FieldCell" width="591">
													<% Response.Write sPayTo %>
															</td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Accounting Head</td>
                                                    <td class="FieldCell" width="591">
                                                            <select size="1" name="selAccHead" class="FormElem" onChange="selAccountHead(this)">
															<option value="A">Select Account Head</option>
									<%
										dim iHeadCount
										'iHeadCount=popFrequentHead(sOrgId,"01",sBookCode)

								    %>
															<option value="G">General Ledger</option>
														<%populatePartyType(sOrgId)%>
                                                    </select>
                                                    </td>
                                                    <input type="hidden" name="hHeadCount" value="<%=iHeadCount%>">
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133"></td>
                                                    <td width="596">
                                                            <span class="DataOnly" id="spAccHead"></span>  </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Pay to / Received from</td>
                                                    <td class="FieldCell" width="591"> <input type="text" name="txtPayTo" size="40" class="Formelem" value="<%=sPayTo%>"  >
                                                    &nbsp; <a href="javascript:SelMisParty()"><img border="0" src="../../assets/images/iTMS Icons/EntryIcon.gif" alt="Miscellaneous Party"></a>
                                                    </td>
                                                        </tr>
                                                        <tr>
                                                    <td width="143" valign="top">
                                                      <table border="0" width="100%" cellspacing="1">
                                                        <tr>
                                                          <td width="50%" class="FieldCellSub">Narration</td>
                                                          <td width="50%" class="FieldCellSub"></td>
                                                        </tr>
                                                      </table>
                                                    </td>
                                                    <td class="FieldCell" valign="top" width="591">
                                                    <textarea rows="3" name="txtNarration" cols="50" class="FormElem"><%Response.Write(sPurVouNarr)%> </textarea> </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Amount</td>
                                                    <td class="FieldCell" width="591"> <input type="text" name="txtAmount" size="15" style="text-align:right" maxlength="13" class="Formelem" onblur="popAddAmount()" value="<%=sVouAmount%>" > </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Deduction @</td>
                                                    <td class="FieldCell" width="591"> <input type="text" name="txtTdsper" value="0.00" size="4" style="text-align:right" maxlength="13" class="Formelem" disabled>
                                                    % On Amount &nbsp; <input type="text" name="txtTdsAmount" value="0.00" size="15" style="text-align:right" maxlength="13" class="Formelem" disabled>
                                                    </td>
                                                        </tr>
                                                        <tr>
                                                    <td class="FieldCellSub" width="133">Approval</td>
                                                    <td class="FieldCell" width="591"> <input type="radio" value="Y" checked name="optApprove" class="FormElem" onClick="ResetList(this)">
                                                      Yes&nbsp;&nbsp;
                            <input type="radio" value="N" name="optApprove" class="FormElem" onClick="ResetList(this)"> No &nbsp;&nbsp; Approver &nbsp;
                            <select size="1" name="selUserId" class="FormElem">
											<option value="I">Immediate Approver</option>
											<%=populateEmployeeWithVal(sUserId)%>
											    </select></td>
                                                        </tr>
                                                            </table>
								</td>
								<td align="center" class="ClearPixel" width="5">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5"><img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack" height="8">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td >
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
				<td class="ExcelHeaderCell" align="center" colspan="2">Document</td>
				<td class="ExcelHeaderCell" align="center" width="275" colspan="4">Amount</td>
		    </tr>
		   <tr>
				<td class="ExcelHeaderCell" align="center">Detail</td>
				<td class="ExcelHeaderCell" align="center">Date</td>
				<td class="ExcelHeaderCell" align="center">Amount</td>
				<td class="ExcelHeaderCell" align="center">Adjusted</td>
				<td class="ExcelHeaderCell" align="center">To Account</td>
				<td class="ExcelHeaderCell" align="center">To adjust</td>
		   </tr>
		</table>
	</div>
</div><!--End of Addtional Details Display  -->
								</td>
								<td align="center" class="ClearPixel" width="5">
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
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="Button" value="Add Entry" name="B11" onClick="AddEntry('A')" class="ActionButton">
                                                                 <input type="button" value="Next" name="B12" onClick="AddEntry('S')" class="ActionButton" >
                                                                 <input type="button" value="Cancel" name="btnCancel" onClick="Cancel('MSIVOUBOOKSELECTION.ASP')" class="ActionButton" >
                                                                <input type="reset" value="Reset" name="B14" class="ActionButton" >
														</td>
													</tr>
												</table>
								</td>
								<td align="center" class="ClearPixel" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
                            <tr>
								<td align="center" class="BottomPack" colspan="3">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5" class="ClearPixel">
								</td>
								<td valign="top">
<DIV class=frmBody id="DisVoucher" style="width:585; visibility:hidden; height:1;">
	<table border="0" cellspacing="1" id="tblVoucher" class="ExcelTable" width="700">
	<tr>
		<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
		<td class="ExcelHeaderCell" align="center" width="75">AU</td>
		<td class="ExcelHeaderCell" align="center">Account Code - Name</td>
		<td class="ExcelHeaderCell" align="center" width="125">Narration</td>
		<td class="ExcelHeaderCell" align="center" width="125">Amount</td>
		<td class="ExcelHeaderCell" align="center">Additional Details</td>
		<td class="ExcelHeaderCell" align="center">Deduction Amount</td>
		<td class="ExcelHeaderCell" align="center">Deduction Percentage</td>
	</tr>
	</table>
</div>
								</td>
								<td align="center" class="ClearPixel" width="5">
								</td>
                            </tr>
<tr>
								<td align="center" class="BottomPack" colspan="3">
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

Function getReceiptType(sRcptAgainst)
Select Case sRcptAgainst
		Case "01" :	getReceiptType = "Purchase Order"
		Case "02" :	getReceiptType = "Supplier Replacements"
		Case "03" :	getReceiptType ="Supplier Samples For Approval"
		Case "04" :	getReceiptType = "Job Order"
		Case "05" :	getReceiptType = "Job Order Rework/Replacement"
		Case "06" : getReceiptType = "Customer Samples"
		Case "07" :	getReceiptType = "Sales Returns"
		Case "08" : getReceiptType = "Return Of Transferred Goods"
		Case "09" :	getReceiptType = "Inter-unit Transfer"
		Case "10" :	getReceiptType = "Without Reference"
	End Select
End Function

Function getReceiptRoute(sRcptRoute)
Select Case sRcptRoute
	Case "DU" : getReceiptRoute = "Direct User"
	Case "IN" : getReceiptRoute = "Inspection"
	Case "ST" : getReceiptRoute = "Stock"
	Case "ID" : getReceiptRoute = "Inspection-Direct"
	Case "IS" : getReceiptRoute = "Inspection-Stock"
	Case "SD" : getReceiptRoute = "Inspection-Stock-Direct"
End Select
End Function

%>