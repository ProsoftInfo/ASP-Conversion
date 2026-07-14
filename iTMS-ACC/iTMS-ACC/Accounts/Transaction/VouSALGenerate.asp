<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	VouSALGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	February  21, 2003
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/NoSeries.asp"-->
<%
'XML DOM Variables
Dim oDOM,objFSO,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTaxRoot,oNodAdvRoot,oNodTemp,sQuery
dim sParCode,sParSubType,sParType
dim EntryNode,HeaderNode,nodANL,newElem,sExp
dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
dim sOrgId,sBookNo,sVouType,iVouNo,sVouName
dim sVouCode,sApprove,sVoucDate,sAccUnit
dim dTotal,sTransType,dCRAmt,dDRAmt
dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo
dim sDocType,sVouStatus
dim iSeriesNo,iSeriesCode
dim sSalType,sRefernceNo,iCounter
dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sSalInvNo,sSalInvDt,dRoundOff,sRoundoffHead
dim iCrTransNo,sCrVouNo,iCrSeriesNo,iCrSeriesCode
Dim sAppTy,sAppBy,sGroupCode,bAddFlag,sAdvNarr,dEntRndVal
dim dAdjAmtTotal
Dim dNoofPack,dPackTy,dRatePer,dItemCode,dClassCode

sVouCode="05"
sVouType="D"
sTransType="SJR"

set objRs  = server.CreateObject("adodb.recordset")
set objFSO = CreateObject("Scripting.FileSystemObject")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
if objFSO.FileExists(server.MapPath("../../Sales/xmldata/General.xml")) then
    oDOM.Load server.MapPath("../../Sales/xmldata/General.xml")
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
end if 


oDOM.load server.MapPath("../temp/transaction/Voucher Entry_SAL_"&Session.SessionID&".xml")

set oNodRoot=oDOM.documentElement

sExp = "//Organization"
Set TempNode = oNodRoot.selectNodes(sExp)
IF TempNode.length <> 0 Then
	sOrgId =  TempNode.Item(0).Attributes.getNamedItem("OrgId").Value
End IF


for each oNodTemp in oNodRoot.childNodes
	if oNodTemp.nodeName="Header" then
		for Each oNodEntry in  oNodTemp.childNodes
			if oNodEntry.nodeName="Book" then
				sBookNo=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="SalesType" then
				sSalType=oNodEntry.Attributes.Item(0).nodeValue
			end if
			if oNodEntry.nodeName="Party" then
				sParType=oNodEntry.Attributes.Item(0).nodeValue
				sParSubType=oNodEntry.Attributes.Item(1).nodeValue
				sParCode=oNodEntry.Attributes.Item(3).nodeValue
			end if
			if oNodEntry.nodeName="SaleInvoice" then
				sSalInvNo=oNodEntry.Attributes.Item(0).nodeValue
				sSalInvDt=oNodEntry.Attributes.Item(1).nodeValue
				sRefernceNo=oNodEntry.Attributes.Item(2).nodeValue
				'sAppTy = Trim(oNodEntry.Attributes.Item(3).nodeValue)
				'sAppBy = Trim(oNodEntry.Attributes.Item(4).nodeValue)

			end if

		next
	end if
	if oNodTemp.nodeName="AgentDetails" then
		set oNodAgent=oNodTemp
	end if

	if oNodTemp.nodeName="Details" then
		set oNodDeatils=oNodTemp
	end if
	if oNodTemp.nodeName="TaxDetails" then
		set oNodTaxRoot=oNodTemp
		dTotal=oNodTemp.Attributes.Item(2).nodeValue
		dRoundOff=oNodTemp.Attributes.Item(3).nodeValue
	end if
	if oNodTemp.nodeName="AdvanceDetails" then
		set oNodAdvRoot=oNodTemp
	end if
next

sAppTy = "Y" 
sAppBy = getUserID()

sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue

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



IF CStr(sAppTy) = "Y" Then
	sVouStatus = "010101" 'Voucher created for approval
Else
	sVouStatus = "010103" 'Voucher created to be approved
End IF


con.BeginTrans

'sQuery="select isnull(max(TransactionNumber),0)+1 from Acc_T_VoucherHeader"
'objRs.open sQuery,con
	'iTransNo=objRs(0)
'objRs.Close

sQuery="select isnull(max(CreatedTransNo),0)+1 from Acc_T_CreatedVoucherHeader"
objRs.open sQuery,con
	iCrTransNo=objRs(0)
objRs.Close

sQuery=" insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
sQuery=sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus) values"
sQuery=sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sSalInvNo&"',convert(datetime,'"&sSalInvDt&"',103),"&dTotal
sQuery=sQuery&",'"&sRefernceNo&"','"&sSalType&"','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"')"

Response.Write sQuery& "<BR><BR>"
'Response.End
con.execute(sQuery)

for each EntryNode in oNodDeatils.childNodes
dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc
	dItemCode = 0
	dClassCode = 0

	sEntryno=EntryNode.Attributes.Item(0).nodeValue
	sAmount=EntryNode.Attributes.Item(2).nodeValue
	sItemDesc=Replace(EntryNode.Attributes.Item(1).nodeValue,"'","''")
	dQty=EntryNode.Attributes.Item(3).nodeValue
	sUOM=EntryNode.Attributes.Item(4).nodeValue
	dBasicAmount=EntryNode.Attributes.Item(7).nodeValue
	dRate=EntryNode.Attributes.Item(6).nodeValue
	dDisPer=EntryNode.Attributes.Item(8).nodeValue
	dDisAmount=EntryNode.Attributes.Item(9).nodeValue
	dEntRndVal = EntryNode.Attributes.Item(10).nodeValue
	dNoofPack = EntryNode.Attributes.Item(11).nodeValue
	dPackTy = EntryNode.Attributes.Item(12).nodeValue
	dRatePer = EntryNode.Attributes.Item(13).nodeValue

	dItemCode = EntryNode.Attributes.Item(14).nodeValue
	dClassCode = EntryNode.Attributes.Item(15).nodeValue

	IF Cstr(dEntRndVal) = "" Then
		dEntRndVal = 0
	End IF

	IF Cstr(dNoofPack) = "" Then
		dNoofPack = 0
	End IF

	IF Cstr(dClassCode) = "" Then
		dClassCode = 0
	End IF

	IF Cstr(dItemCode) = "" Then
		dItemCode = 0
	End IF


	sEntryType="C"
	sNarration=""

	for each HeaderNode in EntryNode.childNodes
		if HeaderNode.nodeName="AccHead" then
				sAccCode=HeaderNode.Attributes.Item(0).nodeValue
				sAccType=HeaderNode.Attributes.Item(4).nodeValue
		end if 'End of Check for Account head Node
	next 'End of Entry Node Loop

		sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
		sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
		sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
		sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,RoundOffvalue,NoofPack,PackCode,RatePer,ItemCode,ClassificationCode ) values ("
		sQuery=sQuery& iCrTransNo&",'"&sOrgId&"'"
		sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
		sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
		sQuery=sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&", "&dEntRndVal&","&dNoofPack&","&dPackTy&","&dRatePer&","&dItemCode&","&dClassCode&" )"
	Response.Write sQuery& "<BR>"
	con.execute(sQuery)

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
		sEntryType = "C"
	else
		sEntryType = "D"
		dTaxAmount=dTaxAmount*-1
	End if

	if sTaxMode="P" then
		sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
			"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
			""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,"&dTaxPer&","&dTaxAmount&")"
	else
		sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
			"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
			""&iCrTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,NULL,"&dTaxAmount&")"
	end if
Response.Write sQuery& "<BR>"
con.execute(sQuery)
iSno=cint(iSno)+1
Next


dim iAdvTranNo,dAdvAmount,iCrAdvno,sAdjType,iCrPayNo
dAdjAmtTotal=0
if  IsObject(oNodAdvRoot) then
	for each EntryNode in oNodAdvRoot.childNodes
		iAdvTranNo=EntryNode.Attributes.Item(0).nodeValue
		dAdvAmount=EntryNode.Attributes.Item(5).nodeValue
		iCrAdvno = EntryNode.Attributes.Item(7).nodeValue
		sAdjType = EntryNode.Attributes.Item(8).nodeValue


		IF CStr(sAdjType) = "B" Then
			dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)

'----------- UPDATION FOR THE ADVANCE ADJUSTMENT DETAILS ----------------------
'----------- INCLUDED ON 05/05/2004 ----------------------
			sQuery = "Select CreatedAdvanceNo From ACC_T_ADVANCEPAYMENTS Where AdvanceNumber = "&iCrAdvno&" "
			objRs.Open sQuery,Con
			IF Not objRs.EOF Then
				iCrAdvno = objRs(0)
			End IF
			objRs.Close

			sQuery = "update Acc_T_CreatedAdvances set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&dAdvAmount &_
					 " where CreatedAdvanceNo = "&iCrAdvno

			Response.Write sQuery &"<br>"
			con.execute(sQuery)

			sAdvNarr = "Sales Invoice No: "&sSalInvNo&" DT: " & sVoucDate
			sQuery = "INSERT INTO ACC_T_CreatedAdvanceAdj (CreatedAdvanceNo, CreatedTransNo,  "&_
					 "AdjustedOn, AmountAdjusted, Narration)  "&_
					 "VALUES ("&iCrAdvno&", "&iCrTransNo&", Convert(Datetime,'"&sVoucDate&"',103), "&dAdvAmount&", '"&sAdvNarr&"') "

			Response.Write sQuery &"<br>"
			con.execute(sQuery)

'------------ END OF UPDATION ---------------------

	'		sQuery="update Acc_T_AdvancePayments set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&dAdvAmount &_
	'		", CreatedTransNo = "&iCrTransNo&" "&_
	'		" where TransactionNumber="&iAdvTranNo
	'		con.execute(sQuery)
		Elseif CStr(sAdjType) = "SR" and CDbl(dAdvAmount) <> 0 Then
			dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)

			sQuery = "Select CRCreatedPayable From Acc_T_Payables Where PayablesNumber = "&iAdvTranNo&" "

			Response.Write sQuery &"<br><br>"
			objRs.Open sQuery,Con
			IF Not objRs.EOF Then
				iCrPayNo = objRs(0)
			End IF
			objRs.Close
			sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
					 "PaidOn,AmountPaid) Values ("&iCrPayNo&","&iCrTransNo&","&_
					 "getdate(),"&dAdvAmount&")"

			Response.Write sQuery &"<br><br>"

			con.execute(sQuery)
		Elseif CStr(sAdjType) = "SC" and CDbl(dAdvAmount) <> 0 Then
			dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)

			sQuery = "Select CRCreatedPayable From Acc_T_Payables Where PayablesNumber = "&iCrAdvno&" "

			Response.Write sQuery &"<br><br>"
			objRs.Open sQuery,Con
			IF Not objRs.EOF Then
				iCrPayNo = objRs(0)
			Else
				iCrPayNo = 0
			End IF
			objRs.Close

			IF CStr(iCrPayable) <> "0" Then

				sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
						 "PaidOn,AmountPaid) Values ("&iCrPayNo&","&iCrTransNo&","&_
						 "getdate(),"&dAdvAmount&")"

				Response.Write sQuery &"<br><br>"

				con.execute(sQuery)
			End IF
		Elseif CStr(sAdjType) = "OT" and CDbl(dAdvAmount) <> 0 Then
			dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)

			sQuery = "Select CRCreatedPayable From Acc_T_Payables Where PayablesNumber = "&iCrAdvno&" "

			Response.Write sQuery &"<br><br>"
			objRs.Open sQuery,Con
			IF Not objRs.EOF Then
				iCrPayNo = objRs(0)
			End IF
			objRs.Close
			sQuery = "insert into Acc_T_CreatedPybleAdjDet(PayablesNumber,CreatedTransNo,"&_
					 "PaidOn,AmountPaid) Values ("&iCrPayNo&","&iCrTransNo&","&_
					 "getdate(),"&dAdvAmount&")"

			Response.Write sQuery &"<br><br>"

			con.execute(sQuery)

		End IF
	Next
end if

Response.Write "============================== " &"<br>"

dim iCrPayable,sPayNarration,iCrPayableNo
sPayNarration="SALE INV NO:"&sSalInvNo&" Dt:"&sSalInvDt


sQuery="select isnull(max(ReceivableNumber),0)+1 from Acc_T_CreatedReceivables"
objRs.open sQuery,con
	iCrPayable=objRs(0)
objRs.Close

sQuery="INSERT INTO Acc_T_CreatedReceivables(ReceivableNumber, CreatedTransNo, OUDefinitionID,"&_
		"VoucherDate, PartyType, PartySubType, PartyCode, PartyInvoiceNumber,"&_
		" PartyInvoiceDate, AmountReceivable, AmountReceived,Narration)values("&iCrPayable&","&iCrTransNo&",'"&sOrgId&"',"&_
		"convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sSalInvNo&"',"&_
		"convert(datetime,'"&sSalInvDt&"',103),"&dTotal&",0,'"&sPayNarration&"')"

Response.Write sQuery &"<br><br>"
con.execute(sQuery)

'----------- UPDATION FOR THE ADVANCE ADJUSTMENT DETAILS ----------------------
'----------- INCLUDED ON 04/05/2004 ----------------------
if CDbl(dAdjAmtTotal) > 0 then
	sQuery="insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
				"ReceivedOn,AmountReceived) Values ("&iCrPayable&","&iCrTransNo&","&_
				"convert(datetime,'"&sSalInvDt&"',103),"&dAdjAmtTotal&")"

	Response.write squery
	con.execute(sQuery)
end if
'------------ END OF UPDATION ---------------------
'Response.Clear

IF CStr(sAppTy) <> "N" Then
	sQuery = "Insert Into Acc_T_VouchersForApproval (CreatedTransNo,ApprovalLevel,ToBeApprovedBy) "&_
 			 "Values("&iCrTransNo&",1,"&sAppBy&" ) "
	Con.Execute sQuery

	Response.Write "<BR><BR>"
End IF

Dim iRecCnt

sQuery = "Select Count(1) From Acc_T_CreatedReceivables Where CreatedTransNo = "&iCrTransNo&" "
Objrs.Open sQuery
IF Not Objrs.Eof Then
	iRecCnt = Objrs(0)
End IF
Objrs.Close

'Con.RollbackTrans
'Response.End


IF Cstr(iRecCnt) = "0" Then
	Con.RollbackTrans
	Response.Clear
	Response.Write "Error While Updating Contact Prosoft "
	Response.End
Else
	con.CommitTrans
End IF

'Con.RollbackTrans
'Response.End

'----To code for Tax Form deatils----------------

'******************************** Cost and ANal Procedure Call ************************

dim sCCGroup,sAddCode,sAddRatio,sAddAmount,iCtr,TempNode

sExp = "//Entry"
Set HeaderNode = oNodRoot.selectNodes(sExp)
IF HeaderNode.length <> 0 Then

	for iCtr = 0 To HeaderNode.length - 1

		sEntryno=HeaderNode.Item(iCtr).Attributes.getNamedItem("No").value

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



if con.Errors.count <>0 then
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else



	'Set newElem  = oDOM.createAttribute("CreatedTransNo")
	'newElem.value = iCrTransNo
	'oNodRoot.setAttributeNode(newElem)

	'Set newElem  = oDOM.createAttribute("CreatedVouNo")
	'newElem.value = sSalInvNo
	'oNodRoot.setAttributeNode(newElem)

	'Set newElem  = oDOM.createAttribute("TransNo")
	'newElem.value = iTransNo
	'oNodRoot.setAttributeNode(newElem)

	'Set newElem  = oDOM.createAttribute("VouNo")
	'newElem.value = sSalInvNo
	'oNodRoot.setAttributeNode(newElem)

	'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")	'Unblocked By UmaMaheswari S,On 28 th June 2011

	Response.Redirect ("VouSALDisplay.asp?TransNo="&iCrTransNo)

end if
%>