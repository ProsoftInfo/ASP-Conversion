<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouPURGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	March 03, 2003
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
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/NoSeries.asp"-->
<%

'XML DOM Variables
Dim oDOM,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTaxRoot,oNodAdvRoot,oNodTemp,sQuery
dim sParCode,sParSubType,sParType
dim EntryNode,HeaderNode,nodANL,newElem
dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
dim sOrgId,sBookNo,sVouType,iVouNo,sVouName
dim sVouCode,sApprove,sVoucDate,sAccUnit
dim dTotal,sTransType,dCRAmt,dDRAmt
dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo
dim sDocType,sVouStatus
dim iSeriesNo,iSeriesCode
dim sSalType,sRefernceNo,iCounter
dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sPurInvNo,sPurInvDt
dim iCrTransNo,sCrVouNo,iCrSeriesNo,iCrSeriesCode
Dim sAppTy,sAppBy,sExp,dRoundOff,sRoundoffHead,sGroupCode,bAddFlag,Objfs
Dim CheckNode,CheckVal,dInvWtRnd,dClassCode,dItemCode,TempNode

dim dAdjAmtTotal
sVouStatus="010105" 'Crearted For Accounting to be Approved
sVouCode="04"
sVouType="C"

set objRs  = server.CreateObject("adodb.recordset")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

Set objfs = CreateObject("Scripting.FileSystemObject")
if Objfs.FileExists(server.MapPath("../../Purchase/xmldata/General.xml")) then
    oDOM.Load server.MapPath("../../Purchase/xmldata/General.xml")
    Set oNodRoot = oDOM.documentElement
    sExp = "//ROUNDOFF"
    Set oNodEntry = oNodRoot.Selectnodes(sExp)
    If oNodEntry.Length > 0 then
	    sRoundoffHead = oNodEntry.Item(0).Attributes.Item(0).nodevalue
    else
	    sRoundoffHead = "0"
    end if
else
     sRoundoffHead = "0"
end if ' if Objfs.FileExists(server.MapPath("../../Purchase/xmldata/General.xml")) then



oDOM.load server.MapPath("../temp/transaction/Voucher Entry_PUR_"&Session.SessionID&".xml")
'Response.Write server.MapPath("../temp/transaction/Voucher Entry_PUR_"&Session.SessionID&".xml")
set oNodRoot=oDOM.documentElement

sExp = "//TaxDetails[@RoundOffValue]"
Set CheckNode = oNodRoot.selectNodes(sExp)
IF CheckNode.length <> 0 Then
	CheckVal = "T"
Else
	CheckVal = "F"
End IF


for each oNodTemp in oNodRoot.childNodes
    Response.Write "<p>sPurInvNo = "& sPurInvNo &" sPurInvDt = "& sPurInvDt &" sAppTy = "& sAppTy &" sAppBy = "& sAppBy 
	if oNodTemp.nodeName="Header" then
		for Each oNodEntry in  oNodTemp.childNodes
			if oNodEntry.nodeName="Organization" then
				sOrgId=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="Book" then
				sBookNo=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="PurchaseType" then
				sSalType=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="Party" then
				sParType=oNodEntry.Attributes.Item(0).nodeValue
				sParSubType=oNodEntry.Attributes.Item(1).nodeValue
				sParCode=oNodEntry.Attributes.Item(3).nodeValue
			end if
			if oNodEntry.nodeName="PurInvoice" then
				sPurInvNo=oNodEntry.getAttribute("PurInvNo")
				sPurInvDt=oNodEntry.getAttribute("PurInvDate")
				sAppTy = oNodEntry.getAttribute("Approval")
				sAppBy = oNodEntry.getAttribute("Approver")
				'Response.Write "sPurInvNo = "& sPurInvNo &" sPurInvDt = "& sPurInvDt &" sAppTy = "& sAppTy &" sAppBy = "& sAppBy 
				sRefernceNo=sPurInvNo&"-"&sPurInvDt
			end if
		next
	end if
	if oNodTemp.nodeName="AgentDetails" then
		set oNodAgent=oNodTemp
	end if
'Response.Write "sPurInvNo = "& sPurInvNo &" sPurInvDt = "& sPurInvDt &" sAppTy = "& sAppTy &" sAppBy = "& sAppBy 
	'if oNodTemp.nodeName="Details" then
	'	IF Trim(oNodTemp.Attributes.Item(0).nodeValue) <> "" Then
	'		set oNodDeatils=oNodTemp
	'	End IF
	'end if
	
	if oNodTemp.nodeName="TaxDetails" then
		set oNodTaxRoot=oNodTemp
		'dTotal=oNodTemp.Attributes.Item(0).nodeValue
		dTotal=oNodTemp.Attributes.Item(2).nodeValue
		IF CStr(CheckVal) = "T" Then
			dRoundOff=oNodTemp.Attributes.Item(3).nodeValue
		Else
			dRoundOff = 0
		End IF
	end if
	if oNodTemp.nodeName="AdvanceDetails" then
		set oNodAdvRoot=oNodTemp
	end if


next

sExp = "//Details"
Set TempNode = oNodRoot.selectnodes(sExp)
For iCounter = 0 To TempNode.length - 1
	IF TempNode.Item(iCounter).Attributes.getNamedItem("BasicValue").Value <> "" Then
		set oNodDeatils = TempNode.Item(iCounter)
	End IF
Next

sTransType="PJR"
'Response.Write "sAppTy = "& sAppTy 
IF Trim(sAppTy) = "Y" Then
	sVouStatus = "010101" 'Voucher To be Approved"
Else
	sVouStatus = "010105" 'Voucher approved to be Accounted
End IF

sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue

'Addition of Roundoff Node
Set newElem = oDOM.CreateElement("Tax")
	newElem.SetAttribute "CatCode","0"
	newElem.SetAttribute "TaxCode","0"
	newElem.SetAttribute "TaxMode","0"
	newElem.SetAttribute "TaxFormula","0"
	newElem.SetAttribute "TaxValue","0"
	newElem.SetAttribute "TaxAmount",Cdbl(dRoundOff)
	newElem.SetAttribute "AccHead",sRoundoffHead
	newElem.Text = "ROUND OFF"
oNodTaxRoot.Appendchild newElem

'IF CStr(sRoundoffHead) = "0" Then
'	sExp = "//Entry"
'	Set CheckNode = oNodRoot.selectNodes(sExp)
'	IF CheckNode.length <> 0 Then
'		dInvWtRnd = CheckNode.Item(0).Attributes.Item(2).NodeValue 
'		dInvWtRnd = CDbl(dInvWtRnd) + CDbl(dRoundOff)
'		CheckNode.Item(0).Attributes.Item(2).NodeValue  = dInvWtRnd
'	End IF
'End IF

dim sAccHeadCode,iTemp,sNoSerDate

sNoSerDate = sVoucDate


sQuery="select DrSeriesNo,DrSeriesCode,CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
	"OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
Response.write "<p>"& sQUery
objRs.open sQuery,con
if not objRs.EOF then
	iSeriesNo=objRs(0)
	iSeriesCode=objRs(1)
	iCrSeriesNo=objRs(2)
	iCrSeriesCode=objRs(3)
end if
objRs.close()

con.BeginTrans

iVouNo=""
'iVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)
sCrVouNo=GenSeriesNumber(sOrgId,iCrSeriesNo,iCrSeriesCode,sNoSerDate)


sQuery="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"

objRs.open sQuery,con
	iCrTransNo=objRs(0)
objRs.Close

sQuery=" insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
sQuery=sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus) values"
sQuery=sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sCrVouNo&"',convert(datetime,'"&sVoucDate&"',103),"&dTotal
sQuery=sQuery&",'"&sRefernceNo&"','"&sSalType&"','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"')"

Response.Write sQuery& "<BR><BR>"
con.execute(sQuery)


for each EntryNode in oNodDeatils.childNodes
dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc

	sEntryno=EntryNode.Attributes.Item(0).nodeValue
	dBasicAmount=EntryNode.Attributes.Item(7).nodeValue
	sItemDesc=replace(EntryNode.Attributes.Item(1).nodeValue,"'","''")
	dQty=EntryNode.Attributes.Item(3).nodeValue
	sUOM=EntryNode.Attributes.Item(4).nodeValue
	sAmount=EntryNode.Attributes.Item(2).nodeValue
	dRate=EntryNode.Attributes.Item(6).nodeValue
	dDisPer=EntryNode.Attributes.Item(8).nodeValue
	dDisAmount=EntryNode.Attributes.Item(9).nodeValue
	dItemCode=EntryNode.Attributes.Item(10).nodeValue
	dClassCode=EntryNode.Attributes.Item(11).nodeValue

	sEntryType="D"
	sNarration=""

	for each HeaderNode in EntryNode.childNodes
		if HeaderNode.nodeName="AccHead" then
				sAccCode=HeaderNode.Attributes.Item(0).nodeValue
				sAccType=HeaderNode.Attributes.Item(4).nodeValue
		end if 'End of Check for Account head Node
	next 'End of Entry Node Loop
	
	''added by ragav on Jan 03,2012
	if Trim(sAccCode)="" or IsNull(sAccCode) then sAccCode = "NULL"
	''end

		sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
		sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
		sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
		sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,ItemCode,ClassificationCode ) values ("
		sQuery=sQuery& iCrTransNo&",'"&sOrgId&"'"
		sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
		sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
		sQuery=sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&","&dItemCode&","&dClassCode&" )"
Response.Write sQuery& "<BR><BR>"
con.execute(sQuery)

'--------------Other details Pending---------------------
next'End of Voucher Node Loop

iSno=1
dim sTaxMode
for each EntryNode in oNodTaxRoot.childNodes
	iCatCode=EntryNode.Attributes.Item(0).nodeValue
	iTaxCode=EntryNode.Attributes.Item(1).nodeValue
	sTaxMode=EntryNode.Attributes.Item(2).nodeValue
	dTaxPer=EntryNode.Attributes.Item(4).nodeValue
	dTaxAmount=EntryNode.Attributes.Item(5).nodeValue
	sAccCode=EntryNode.Attributes.Item(6).nodeValue
	if dTaxAmount >=0 then
		sEntryType = "D"
	else
		sEntryType = "C"
		'dTaxAmount=dTaxAmount*-1
		dTaxAmount = Abs(dTaxAmount)
		dTaxPer = abs(dTaxPer)
	End if
	'if CInt(sAccCode)>0 then
		if sTaxMode="P" then
			sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
				"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
				""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,"&dTaxPer&","&dTaxAmount&")"
		else

			sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
				"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
				""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,NULL,"&dTaxAmount&")"
		end if

		Response.Write sQuery& "<BR><BR>"
		con.execute(sQuery)
		iSno=cint(iSno)+1
	'end if
Next

dim iAdvTranNo,dAdvAmount,iCrAdvNo,sAdjType,iCrRecNo
dAdjAmtTotal=0
for each EntryNode in oNodAdvRoot.childNodes
	iAdvTranNo=EntryNode.Attributes.Item(0).nodeValue
	dAdvAmount=EntryNode.Attributes.Item(5).nodeValue
	'iCrAdvNo = EntryNode.Attributes.Item(6).nodeValue
	iCrAdvNo = EntryNode.Attributes.Item(7).nodeValue
	sAdjType = EntryNode.Attributes.Item(8).nodeValue

	'=====================================================================================
	'Added On	:	29/11/2004
	'Reason		:	To update the adjustment based on the selection Bank or Cash
	'			Debit Notes.
	'=====================================================================================
	Dim sAdjNarr
	sAdjNarr = "Adjusted to Purchase Invoice No "&sCrVouNo&" DT: "& sVoucDate &" Amt: "& dTotal
	if CDbl(dAdvAmount)>0 then
		IF CStr(sAdjType) = "I" and CDbl(dAdvAmount) <> 0 Then ' Voucher Like Cask/Bank Payment
			dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)

'----------- UPDATION FOR THE ADVANCE ADJUSTMENT DETAILS ----------------------
'----------- INCLUDED ON 04/05/2004 ----------------------
			sQuery = "Select CreatedAdvanceNo From ACC_T_ADVANCEPAYMENTS Where AdvanceNumber = "&iCrAdvno&" "
			objRs.Open sQuery,Con
			IF Not objRs.EOF Then
				iCrAdvno = objRs(0)
			End IF
			objRs.Close
			'sQuery = "update Acc_T_CreatedAdvances set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&dAdvAmount &_
			'		 " where CreatedTransNo = "&iCrAdvNo

			sQuery = "update Acc_T_CreatedAdvances set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&dAdvAmount &_
					 " where CreatedAdvanceNo = "&iCrAdvNo


			Response.Write sQuery &"<br><br>"
			con.execute(sQuery)

			sQuery = "INSERT INTO ACC_T_CREATEDADVANCEADJ (CREATEDADVANCENO, CREATEDTRANSNO, ADJUSTEDON,  "&_
					 "AMOUNTADJUSTED, NARRATION) VALUES "&_
					 "("&iCrAdvNo&", "&iCrTransNo&", Convert(Datetime,'"&sVoucDate&"',103), "&dAdvAmount&", '"&sAdjNarr&"') "

			Response.Write sQuery &"<br><br>"
			con.execute(sQuery)

		Elseif CStr(sAdjType) = "D" and CDbl(dAdvAmount) <> 0 Then ' If the adjustment is of Debit Notes.
			dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)
			sQuery = "Select DRCreatedReceivable From Acc_T_Receivables Where ReceivableNumber = "&iAdvTranNo&" "

			Response.Write sQuery &"<br><br>"
			objRs.Open sQuery,Con
			IF Not objRs.EOF Then
				iCrRecNo = objRs(0)
			Else
				iCrRecNo = 0
			End IF
			objRs.Close

			sQuery = "insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
					 "ReceivedOn,AmountReceived) Values ("&iCrRecNo&","&iCrTransNo&","&_
					 "getdate(),"&dAdvAmount&")"

			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery

		End IF
'------------ END OF UPDATION ---------------------
	end if

Next


dim sPayNarration,iCrPayableNo
if CDbl(dTotal)>CDbl(dAdjAmtTotal) then
	sQuery="select isnull(max(PayablesNumber),0)+1 from Acc_T_CreatedPayables"
	Response.Write "<p>"&sQuery
	objRs.open sQuery,con
		iCrPayableNo=objRs(0)
	objRs.Close
	sPayNarration="PUR INV No:"&sPurInvNo &" Dt:"&sPurInvDt

	sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber, CreatedTransNo, OUDefinitionID,"&_
			"VoucherDate, PartyType, PartySubType, PartyCode, PartyBillNumber,"&_
			" PartyBillDate, AmountPayable, AmountPaid,Narration)values("&iCrPayableNo&","&iCrTransNo&",'"&sOrgId&"',"&_
			"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sPurInvNo&"',"&_
			"convert(datetime,'"&sPurInvDt&"',103),"&dTotal&",0,'"&sPayNarration&"')"
    Response.Write "<p>"&sQuery
	con.execute(sQuery)

'----------- UPDATION FOR THE ADVANCE ADJUSTMENT DETAILS ----------------------
'----------- INCLUDED ON 04/05/2004 ----------------------


	sQuery="insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
				"PaidOn,AmountPaid) Values ("&iCrPayableNo&","&iCrTransNo&","&_
				"convert(datetime,'"&sVoucDate&"',103),"&dAdjAmtTotal&")"
	Response.Write "<p>"&sQuery
	con.execute(sQuery)

'------------ END OF UPDATION ---------------------

	sQuery = "Insert Into Acc_T_VouchersForApproval (CreatedTransNo,ApprovalLevel,ToBeApprovedBy) "&_
			 "Values("&iCrTransNo&",1,"&sAppBy&" ) "
	Response.Write "<p>"&sQuery
	Con.Execute sQuery


end if

'================== Checking for Debit and Credit Tally ===================================
	Dim dDebTotal,dCreTotal
	'sQuery = "Select VoucherAmount,CrDrIndication From Acc_T_CreatedVoucherHeader  "&_
	'			 "Where CreatedTransNo = "&iCrTransNo&" "
	'Objrs.Open sQuery,Con
	'IF Not Objrs.Eof Then
	'	IF Cstr(Objrs(1)) = "C" Then
	'		dCreTotal = Cdbl(Objrs(0))
	'	Else
	'		dDebTotal = Cdbl(Objrs(0))
	'	End IF
	'End IF
	'objrs.Close

	'sQuery = "Select Sum(Amount),TransCrDrIndication From Acc_T_CreatedVoucherDetails Where "&_
	'		 "CreatedTransNo = "&iCrTransNo&" Group By TransCrDrIndication "

	'Objrs.Open sQuery,Con
	'Do While Not Objrs.Eof
	'	IF Cstr(Objrs(1)) = "C" Then
	'		dCreTotal = Cdbl(dCreTotal) + Cdbl(Objrs(0))
	'	Else
	'		dDebTotal = Cdbl(dDebTotal) + Cdbl(Objrs(0))
	'	End IF
	'	Objrs.MoveNext
	'loop
	'objrs.Close

	''''''''''''''''''''''''''''''''''''''''''' This is Not Required Fro the Created Voucher Since the accounted amount gets added with the item amount this entries is only for amedment stage only
	
	'sQuery = "Select Sum(TaxAmount),TransCrDrIndication From Acc_T_CreatedVoucherTaxDet Where "&_
	'		 "CreatedTransNo = "&iCrTransNo&" Group By TransCrDrIndication "

	'Objrs.Open sQuery,Con
	'Do While Not Objrs.Eof
	'	IF Cstr(Objrs(1)) = "C" Then
	'		dCreTotal = Cdbl(dCreTotal) + Cdbl(Objrs(0))
	'	Else
	'		dDebTotal = Cdbl(dDebTotal) + Cdbl(Objrs(0))
	'	End IF
	'	Objrs.MoveNext
	'loop
	'objrs.Close


	'Response.Write dCreTotal &"   " & dDebTotal &"<br><br>"
	
	'IF Cdbl(dCreTotal) <> Cdbl(dDebTotal) Then
	'	'Response.Clear
	'	Response.Write "Error While Creating Purchase Invoice "
	'	Response.Write "<br>"
	'	'Response.Write "Debit Total " & FormatNumber(dDebTotal,2,,,0) &" Credit Total " & FormatNumber(dCreTotal,2,,,0)
	'	Con.RollbackTrans
	'	Response.End
	'End IF
		
'================== Checking for Debit and Credit Tally ===================================


'Con.RollbackTrans
'Response.End
Con.CommitTrans


'============================================================================================

'******************************** Cost and ANal Procedure Call ************************

dim sCCGroup,sAddCode,sAddRatio,sAddAmount,iCtr

sExp = "//Entry"
Set HeaderNode = oNodRoot.selectNodes(sExp)
IF HeaderNode.length <> 0 Then

	for iCtr = 0 To HeaderNode.length - 1

		sEntryno=HeaderNode.Item(iCtr).Attributes.getNamedItem("No").value
		'sEntryType=HeaderNode.Item(iCtr).Attributes.getNamedItem("CRDR").value
		'sAmount=HeaderNode.Item(iCtr).Attributes.getNamedItem("Amount").value
		'sAccUnit=HeaderNode.Item(iCtr).Attributes.getNamedItem("AccUnit").value

		sExp="//Entry[@No='"&sEntryno&"']/AccHead"
		set tempNode=oNodRoot.selectNodes(sExp)
		sAccCode=tempNode.item(0).Attributes.getNamedItem("No").value

		set nodANL=oDOM.createElement("Root")
		nodANL.setAttribute "TransNo",iCrTransNo
		nodANL.setAttribute "EntryNo",sEntryno
		nodANL.setAttribute "UnitCode", sorgid
		nodANL.setAttribute "GlHead",sAccCode
		nodANL.setAttribute "ACTFlag","C"



		sExp="//Entry[@No='"&sEntryno&"']/CostCenter"
		set tempNode=oNodRoot.selectNodes(sExp)
		if tempNode.length >0 then
			set EntryNode=tempNode.item(0).cloneNode(true)
			nodANL.appendChild(EntryNode)
			bAddFlag=true
		end if

		sExp="//Entry[@No='"&sEntryno&"']/Analytical"
		set tempNode=oNodRoot.selectNodes(sExp)
		if tempNode.length >0 then
			set EntryNode=tempNode.item(0).cloneNode(true)
			nodANL.appendChild(EntryNode)
			bAddFlag=true
		end if
		if bAddFlag then
		  Dim adoConn

		   Set adoConn = Server.CreateObject("ADODB.Connection")
		   adoConn.ConnectionString = con
		   adoConn.CursorLocation = 3
		   adoConn.Open

		   sQuery = "Proc_VouCCANALUpdate"


		   Dim adoCmd
		   Set adoCmd = Server.CreateObject("ADODB.Command")
		   Set adoCmd.ActiveConnection =adoConn
		   adoCmd.CommandText = sQuery
		   adoCmd.CommandType = 4 'adCmdStoredProc
		   adoCmd.Parameters.Append adoCmd.CreateParameter("@XMLDoc",201,1,len(nodANL.xml),nodANL.xml)



		   Dim adoRS
		   Set adoRS = adoCmd.Execute()

		end if

	next 'End of Entry Node Loop
End IF

'******************************** Coast and ANal Procedure Call ************************


'----To code for Tax Form deatils----------------

if con.Errors.count <>0 then
	con.RollbackTrans

	for iCounter=0 to con.Errors.count-1
		Response.Write con.Errors(iCounter).Description &"<br>"
		Response.Write con.Errors(iCounter).Source &"<br>"
	next
	'Redirect to Error Handling System
else

	'con.CommitTrans

	'oDOM.load server.MapPath("../temp/transaction/Voucher Entry_PUR_"&Session.SessionID&".xml")

	Set newElem  = oDOM.createAttribute("CreatedTransNo")
	newElem.value = iCrTransNo
	oNodRoot.setAttributeNode(newElem)

	Set newElem  = oDOM.createAttribute("CreatedVouNo")
	newElem.value = sCrVouNo
	oNodRoot.setAttributeNode(newElem)

'	Set newElem  = oDOM.createAttribute("TransNo")
'	newElem.value = iTransNo
'	oNodRoot.setAttributeNode(newElem)

'	Set newElem  = oDOM.createAttribute("VouNo")
'	newElem.value = iVouNo
'	oNodRoot.setAttributeNode(newElem)

'This been Blocked
'	oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")	'block is Removed By UmaMaheswari S, On June 23,2011
	
'	if Objfs.FileExists(Server.MapPath("../temp/transaction/Voucher Entry_PUR_"&Session.SessionID&".xml")) then
'	    objfs.DeleteFile(Server.MapPath("../temp/transaction/Voucher Entry_PUR_"&Session.SessionID&".xml"))
'	end if 'if Objfs.FileExists(Server.MapPath("../temp/transaction/Voucher Entry_PUR_"&Session.SessionID&".xml")) then
	Response.Redirect ("VouPURDisplay.asp?TransNo="&iCrTransNo)
end if

%>
