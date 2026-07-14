<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AmdInvPurInvoiceInsert_New.asp
	'Module Name				:	Purchase (Transaction - Invoice entry)
	'Author Name				:	Kalaiselvi R
	'Created On					:	January 27, 2006
	'Modified By				:	Ragavendran R
	'Modified On				:	Jun 14,2010
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
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/Purpopulate.asp"-->
<!--#include virtual="/include/NoSeries.asp"-->
<!--#include virtual="/include/sessionverify.asp"-->

<%
'declaring variables

'Invoice Header Variables
Dim sOrgID,sPartyType,iPartySubtype,iPartyCode,sPurType,sCurrency,sInvAgainst,saRefNum,sRefNum
Dim sInvDt,iCurrency,iPayTerms,iBasisPricing,iDespatchMode,iPayMode,iTransporterCode
Dim iDestPort,iLoadPort,sIssueBank,iBeficBank,sRemarks,iTotBasValue,iTotDisBasValue,sQuery
Dim iTotTaxValue,iTotInvValue,nInvNo,iTotNetVal,sSuppInvNo,sSuppInvDt
Dim sTransFlag, dRoundOffValue,sTransporterName,sInvoiceFlag,sByUnit,sOrgFlag
Dim sConvFactor, sConvAson, iOperCurrency,sSqlCmd,sBillType,sAddDescr,sAttributeList

Dim sVouStatus,sVouCode,sVouType,sTransType,sEntryType,sSql,sPartyName
Dim sItemDesc,sUOM,sNarration,sRCrtype,sCrVouNo,sRateUoM,sAmdBy
Dim sAddCode,sAddRatio,sAddAmount,sTaxMode,sPayNarration,sInvCatCode
Dim ndTaxDet,nOrgAccHdNo,ndTaxPackDet,nTaxPackCode,nTaxPackRate


Dim iTaxCat,iTaxCode,dTaxPer,dTaxAmount,nAccCode,iFormNo,iPayableNo,iCnt,iPurTranNum
Dim nEntryNo,iNetBasic,dQty,nAmount,dRate,iDisPer,dDisAmount,iItemCode,iClassCode
Dim iSno,dAdjAmtTotal,iAdvTranNo,dAdvAmount
Dim iCrTransNo,iRegCode,iTaxCrEligible,nAmdNo,sVat
Dim dRatePerQtyUoM,iItemBasicVal,iItemValue,iItemCalcRate,iTempPurType

Dim dtVoucDate,dtSuppInvDate

Dim blnPayableFlag,blnTaxFlag,blnLRFlag,blnAdvDetExist


Dim rsTemp,objFS,oDOM,objRs,oDOMAccData

Dim oNodRoot,oNodTemp,oNodEntry,oNodDeatils,oNodTaxRoot,oNodAdvRoot,EntryNode
Dim newElem,HeaderNode,nodANL,AccHeadElem
Dim oNodAccRoot

Set objFS = Server.CreateObject("Scripting.FileSystemObject")

sVouStatus = "010101" 'Voucher created for Approval
sVouCode = "04"
sVouType = "C"
sTransType="PJR"
sEntryType="D"

Set objRs  = server.CreateObject("adodb.recordset")
Set rsTemp = server.CreateObject("adodb.recordset")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set oDOMAccData = server.CreateObject("Microsoft.XMLDOM")

oDOM.load server.MapPath("../temp/transaction/AmdNewInvItemValue_PUR_"&Session.SessionID&".xml")

blnTaxFlag = false
blnLRFlag = false
blnAdvDetExist = false

Set oNodRoot=oDOM.documentElement

For each oNodTemp in oNodRoot.childNodes
	if oNodTemp.nodeName="InvoiceHeader" then
		for Each oNodEntry in  oNodTemp.childNodes
			if oNodEntry.nodeName="Header" then
								
				sPartyName	= oNodEntry.Attributes.getNamedItem("Party").value
				
				sOrgID		= oNodEntry.Attributes.Item(0).nodeValue
				sByUnit		= sOrgID
				sPurType	= oNodEntry.Attributes.Item(2).nodeValue
				sInvAgainst = oNodEntry.Attributes.Item(4).nodeValue
				sRefNum		= oNodEntry.Attributes.Item(5).nodeValue
				iPartyCode	= oNodEntry.Attributes.Item(6).nodeValue
				sPartyType	= oNodEntry.Attributes.Item(7).nodeValue
				iPartySubtype  = oNodEntry.Attributes.Item(8).nodeValue
				iCurrency	= oNodEntry.Attributes.Item(9).nodeValue
				iDespatchMode  = oNodEntry.Attributes.Item(10).nodeValue
				iPayMode	= oNodEntry.Attributes.Item(11).nodeValue
				iPayTerms	= oNodEntry.Attributes.Item(12).nodeValue
				sIssueBank	= oNodEntry.Attributes.Item(13).nodeValue
				iBeficBank	= oNodEntry.Attributes.Item(14).nodeValue
				iBasisPricing = oNodEntry.Attributes.Item(15).nodeValue
				iTransporterCode  = oNodEntry.Attributes.Item(16).nodeValue
				iLoadPort	= oNodEntry.Attributes.Item(17).nodeValue
				iDestPort	= oNodEntry.Attributes.Item(18).nodeValue
				sRemarks	= PackQuote(oNodEntry.Attributes.Item(19).nodeValue)
				sSuppInvNo = PackQuote(oNodEntry.Attributes.Item(20).nodeValue)
				dtSuppInvDate = oNodEntry.Attributes.Item(21).nodeValue
				sTransFlag = oNodEntry.getAttribute("TransporterFlag")
				sInvoiceFlag = oNodEntry.getAttribute("InvoiceFlag")
				sOrgFlag = oNodEntry.getattribute("OriginalInvoice")
				'	sInvDt		= day(date()) & "/" & month(date()) & "/" & year(date())
				'sInvDt = formatdate(date())
				
				sCrVouNo = sSuppInvNo

				sInvDt = oNodEntry.Attributes.getNamedItem("SuppInvDt").value

				'Newly Added by Tajudeen on 28-04-2005
				sConvFactor = oNodEntry.getAttribute("ConversationRate")
				sConvAson = oNodEntry.getAttribute("ConversationAsOn")
				iOperCurrency = oNodEntry.getAttribute("OperatingCurrency")

				nInvNo			= oNodEntry.getAttribute("InvoiceNumber")
				
				dRoundOffValue  = oNodEntry.getAttribute("RoundOff")
				iTotInvValue	= oNodEntry.getAttribute("InvValue")
				sInvCatCode = oNodEntry.getAttribute("InvCategory")
				sBillType = oNodEntry.getAttribute("BillType")
			end if
			
		next
	end if 'if oNodTemp.nodeName="InvoiceHeader" then

	if oNodTemp.nodeName="Voucher" then
		set oNodAccRoot = oNodTemp
		oDOMAccData.appendChild oNodAccRoot
	end if 
	
	if oNodTemp.nodeName="ItemDetails" then
		set oNodDeatils=oNodTemp
	end if
	if oNodTemp.nodeName="TaxDetails" then
		set oNodTaxRoot=oNodTemp
	end if
	if oNodTemp.nodeName="AdvanceDetails" then
		set oNodAdvRoot=oNodTemp
		blnAdvDetExist = true
	end if

	If oNodTemp.nodeName="ExemptionForms" then
		blnTaxFlag = true
		set TaxFormRoot = oNodTemp
	end if
	
	if oNodTemp.nodeName="AccountHead" then
		set AccHeadElem = oNodTemp
	end if 		
Next

dtVoucDate=formatdate(date())


if trim(iDespatchMode)	= "" then iDespatchMode = 0
if trim(iPayMode)	= "" then iPayMode = 0
if trim(iPayTerms)	= "" then iPayTerms = 0
if trim(sIssueBank)	= "" then sIssueBank = 0
if trim(iBeficBank)	= "" then iBeficBank = 0
if trim(iBasisPricing)	= "" then iBasisPricing = 0
if trim(iTransporterCode)	= "" then iTransporterCode = 0
if trim(iLoadPort)	= "" then iLoadPort = 0
if trim(iDestPort)	= "" then iDestPort = 0



''Inv Against
IF sInvAgainst="Other" then sInvAgainst="0"
IF sInvAgainst="Receipt" then sInvAgainst="1"
IF sInvAgainst="Order" then sInvAgainst="2"

if trim(sPartyType) = "" then sPartyType = ""
if trim(iPartyCode) = "" then iPartyCode = "0"

iTotBasValue = 0
iTotNetVal = 0
iTotTaxValue = 0
iTotDisBasValue = 0

if trim(sPurType) = "0" then
	For each oNodTemp in oNodRoot.childNodes
		if oNodTemp.nodeName="TaxDetails" then
			set  oNodTaxRoot = oNodTemp
			iTotBasValue	= cdbl(iTotBasValue)+ oNodTaxRoot.attributes.item(0).nodevalue
			iTotNetVal		= cdbl(iTotNetVal)	+ oNodTaxRoot.attributes.item(1).nodevalue
			iTotTaxValue	= cdbl(iTotTaxValue)+ oNodTaxRoot.attributes.item(2).nodevalue
		end if 
	Next		
	iTotDisBasValue = cdbl(iTotBasValue) - cdbl(iTotNetVal)
else
	iTotBasValue = oNodTaxRoot.attributes.item(0).nodevalue
	iTotNetVal =  oNodTaxRoot.attributes.item(1).nodevalue
	iTotTaxValue = oNodTaxRoot.attributes.item(2).nodevalue
	iTotDisBasValue = cdbl(iTotBasValue) - cdbl(iTotNetVal)
end if 

con.BeginTrans


sAmdBy = session("employeeNumber")	
	
'appending data from Invoice header table to Amendment table

''Auto Increment Invoice Number
With rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.Source =  "SELECT ISNULL(MAX(AmendmentNumber) + 1, 1) AS InvoiceNo FROM RCV_A_InvoiceHeader"
	.ActiveConnection = con
	.Open
End With
Set rsTemp.ActiveConnection = nothing

if not rsTemp.EOF then
	nAmdNo	=	rsTemp(0)
end if 	
rsTemp.Close


sSqlCmd = "Insert Into RCV_A_InvoiceHeader(AmendmentNumber,AmendedBy,AmendedOn," &_
	"InvoiceNumber,InvoiceCode,InvoiceType,SuppInvoiceNo,SuppInvoiceDate,InvoiceRcptUnit,InvoiceAgainst,PurchaseType,ReferenceNumber,InvoiceByUnit,InvoiceDate,PartyType,PartySubType,PartyCode,CurrencyCode,PaymentTerms,BasisOfPricing,DespatchMode,PaymentMode,TransporterCode,DestinationPort,LoadingPort,IssueBank,BenifBank,Remarks,InvBasicValue,InvDiscBasicValue,InvTaxValue,GrossWeight,NettWeight,PackingCases,TotalInvoiceValue,Accounted,RoundOffvalue,TransporterName,OrgSupplierInvoice,OrgSupplierName,OrgSupplierInvNo,OrgSupplierInvDate,OrgSuppDutyCatCode,OrgSuppDutyTaxtCode,OrgSuppDutyAmount,ConversionFactor,ConversionAson,OperatingCurrency) " & _ 
	"Execute('Select "&nAmdNo&","&sAmdBy&",convert(datetime,''"&formatDate(date())&"'',103)," & _
	"InvoiceNumber,InvoiceCode,InvoiceType,SuppInvoiceNo,SuppInvoiceDate,InvoiceRcptUnit,InvoiceAgainst,PurchaseType,ReferenceNumber,InvoiceByUnit,InvoiceDate,PartyType,PartySubType,PartyCode,CurrencyCode,PaymentTerms,BasisOfPricing,DespatchMode,PaymentMode,TransporterCode,DestinationPort,LoadingPort,IssueBank,BenifBank,Remarks,InvBasicValue,InvDiscBasicValue,InvTaxValue,GrossWeight,NettWeight,PackingCases,TotalInvoiceValue,Accounted,RoundOffvalue,TransporterName,OrgSupplierInvoice,OrgSupplierName,OrgSupplierInvNo,OrgSupplierInvDate,OrgSuppDutyCatCode,OrgSuppDutyTaxtCode,OrgSuppDutyAmount,ConversionFactor,ConversionAson,OperatingCurrency" & _
	" From  RCV_T_InvoiceHeader Where InvoiceNumber=" & nInvNo & "')"
Response.Write "<p>" & sSqlCmd
con.execute(sSqlCmd)


'Inserting Record in RCV_A_InvoiceDetails
sSqlCmd = "Insert into RCV_A_InvoiceDetails(AmendmentNumber,InvoiceNumber,EntryNo,OrganisationCode,ClassificationCode,ItemCode,TempItemCode,AccountHead,InvItemUOM,InvoiceQuantity,InvoiceRate,ItemBasicValue,ItemDiscountPercent,ItemDiscountValue,ItemNettBasicValue,ItemValue,ItemRate,RateUoM,PurchaseType)" &_
							" execute ('Select "& nAmdNo &",InvoiceNumber,EntryNo,OrganisationCode,ClassificationCode,ItemCode,TempItemCode,AccountHead,InvItemUOM,InvoiceQuantity,InvoiceRate,ItemBasicValue,ItemDiscountPercent,ItemDiscountValue,ItemNettBasicValue,ItemValue,ItemRate,RateUoM,PurchaseType from RCV_T_InvoiceDetails where InvoiceNumber="& nInvNo &"')"
Response.Write "<p> " & sSqlCmd
con.Execute(sSqlCmd)


'Inserting Record in RCV_A_InvoiceTaxDetails
sSqlCmd = "Insert into RCV_A_InvoiceTaxDetails(AmendmentNumber,InvoiceNumber,TaxCategoryCode,TaxCode,TaxCreditEligibility,TaxAmount,TaxPercentage,PurchaseType)" &_
								"execute ('Select "& nAmdNo &",InvoiceNumber,TaxCategoryCode,TaxCode,TaxCreditEligibility,TaxAmount,TaxPercentage,PurchaseType from RCV_T_InvoiceTaxDetails where InvoiceNumber="& nInvNo &"')"
Response.Write "<p> " & sSqlCmd
con.Execute(sSqlCmd)

'Inserting Record in RCV_A_TransportDetail
sSqlCmd = "Insert into RCV_A_TransportDetail(AmendmentNumber,RefferenceNo,LRNumber,UCNumber,EntryStage,LRDate,VehicleNumber,GrossWeight,NettWeight,PackingCases) " &_
							  "execute ('Select "& nAmdNo &",RefferenceNo,LRNumber,UCNumber,EntryStage,LRDate,VehicleNumber,GrossWeight,NettWeight,PackingCases from RCV_T_TransportDetail where RefferenceNo="& nInvNo &"')"
Response.Write "<p> " & sSqlCmd
con.Execute(sSqlCmd)

'Deleting Record from RCV_T_InvoiceDetails
sSqlCmd = "delete from RCV_T_InvoiceDetails where InvoiceNumber="& nInvNo  & ""
Response.Write "<p> " & sSqlCmd
con.Execute(sSqlCmd)

'Deleting Record from RCV_T_InvoiceTaxDetails
sSqlCmd = "Delete from RCV_T_InvoiceTaxDetails where InvoiceNumber="& nInvNo  & ""
Response.Write "<p> " & sSqlCmd
con.Execute(sSqlCmd)

'Deleting Record from RCV_T_TransportDetail
sSqlCmd = "Delete from RCV_T_TransportDetail where RefferenceNo="& nInvNo  & ""
Response.Write "<p> " & sSqlCmd
con.Execute(sSqlCmd)

sSqlCmd = "Delete from RCV_T_InvoiceTaxPacDet where InvoiceNumber = "& nInvNo 
Response.Write "<p>" & sSqlCmd 
con.execute(sSqlCmd)





iCurrency = 1
sConvFactor = 1
sConvAson = sInvDt
iOperCurrency =1
	
'Update RCV_T_InvoiceHeader
sSql = "Update RCV_T_InvoiceHeader set " &_
		" SuppInvoiceNo='" & sSuppInvNo & "',SuppInvoiceDate=convert(datetime,'" & dtSuppInvDate  & "',103),InvoiceRcptUnit='" & sOrgID & "'," & _
		" InvoiceAgainst='" & sInvAgainst & "',PurchaseType=" & sPurType & ",ReferenceNumber='" & sRefNum& "',InvoiceDate=convert(datetime,'" & sInvDt & "',103),PartyType='" & sPartyType & "'," &_
		" PartySubType=" & iPartySubtype & ",PartyCode=" & iPartyCode & ",CurrencyCode=" & iCurrency & ",PaymentTerms=" & iPayTerms & ",BasisOfPricing=" & iBasisPricing & "," &_
		" DespatchMode=" & iDespatchMode & ",PaymentMode=" & iPayMode & ",TransporterCode=" & iTransporterCode & ",DestinationPort=" & iDestPort & ",LoadingPort=" & iLoadPort & "," &_
		" IssueBank=" & sIssueBank & ",BenifBank=" & iBeficBank & ",Remarks='" & sRemarks & "',InvBasicValue=" & iTotBasValue & ",InvDiscBasicValue=" & iTotDisBasValue & "," &_
		" InvTaxValue=" & iTotTaxValue & ",TotalInvoiceValue=" & iTotInvValue & ",InvoiceType='" & trim(sTransFlag) & "',RoundOffvalue=" & trim(dRoundOffValue) & "," &_
		" TransporterName='" & trim(sTransporterName) & "',InvoiceByUnit='" & trim(sByUnit) & "',ConversionFactor=" & sConvFactor & ",ConversionAson=convert(datetime,'" & sConvAson & "',103)," &_
		" OperatingCurrency=" & iOperCurrency & ",InvCategoryCode = '"&sInvCatCode&"',BillType='"&sBillType&"' where InvoiceNumber = " & nInvNo & ""
	' kk test
Response.Write "<br><br>" + ssql
con.execute sSql


'fetch CreatedVoucherNo details from RCV_T_InvoiceHeader table
iCrTransNo = 0
With rsTemp
	.CursorLocation = 3
	.CursorType = 3
	.Source =  "SELECT IsNull(CreatedVoucherNo,0) FROM RCV_T_InvoiceHeader where InvoiceNumber = " & nInvNo & ""
	.ActiveConnection = con
	.Open
End With
Set rsTemp.ActiveConnection = nothing

if not rsTemp.EOF then
	iCrTransNo	=	rsTemp(0)
end if 	
rsTemp.Close

'Response.Write "<p> iCrTransNo = "& trim(iCrTransNo )



sQuery= "Update Acc_T_CreatedVoucherHeader "
sQuery=sQuery & " set OUDefinitionID='"&sOrgId&"',BookCode='"&sVouCode&"',BookNumber=NULL,TransactionType='"&sTransType&"',"
sQuery=sQuery & "PartyType='"&sPartyType&"',PartySubType="&iPartySubtype&",PartyCode="&iPartyCode&",AccountHead=NULL,CreatedVoucherNo='"&sCrVouNo&"',VoucherDate=convert(datetime,'"&dtVoucDate&"',103),VoucherAmount="& iTotInvValue
sQuery=sQuery & ",PayToRecdFrom='"&sPartyName&"',BankInstrumentType='"&sPurType&"',CrDrIndication='"&sVouType&"',CreatedBy="&getUserid&",CreatedOn=getdate(),ApprovedBy=NULL,CreatedVouchStatus='"&sVouStatus&"',FromApplication='2',OtherApplnTransNo=" & nInvNo & ",OtherApplnTableName='RCV_T_INVOICEHEADER',InvCategoryCode = '"&sInvCatCode&"',PurchaseBillType='"&sBillType&"' "
sQuery=sQuery & " where CreatedTransNo = "& iCrTransNo & ""
Response.Write "<br><br> aaaaa = " + sQuery
con.execute(sQuery)

'removing data from Accounts related table(s)
sSqlCmd = "delete from Acc_T_CreatedVoucherDetails where CreatedTransNo="& iCrTransNo  & ""
Response.Write "<p> " & sSqlCmd
con.Execute(sSqlCmd)

sSqlCmd = "delete from Acc_T_CreatedVoucherCCDet where CreatedTransNo="& iCrTransNo  & ""
Response.Write "<p> " & sSqlCmd
con.Execute(sSqlCmd)

sSqlCmd = "delete from Acc_T_CretedVoucherAHDet where CreatedTransNo="& iCrTransNo  & ""
Response.Write "<p> " & sSqlCmd
con.Execute(sSqlCmd)

sSqlCmd = "delete from PUR_T_ExemptionForms where InvoiceNumber="& nInvNo  & ""
Response.Write "<p> " & sSqlCmd
con.Execute(sSqlCmd)

'sSqlCmd = "delete from PUR_T_FormsUtilised where UtilisedInvoiceNumber="& nInvNo  & ""
'Response.Write "<p> " & sSqlCmd
'con.Execute(sSqlCmd)

'sSqlCmd = "delete from RCV_T_ModvatCreditTaken where InvoiceNumber="& nInvNo  & ""
'Response.Write "<p> " & sSqlCmd
'con.Execute(sSqlCmd)

sSqlCmd = "delete from Acc_T_CreatedVoucherTaxDet where CreatedTransNo="& iCrTransNo  & ""
Response.Write "<p> " & sSqlCmd
con.Execute(sSqlCmd)

sSqlCmd = "delete from Acc_T_CreatedPayables where CreatedTransNo="& iCrTransNo  & ""
Response.Write "<p> " & sSqlCmd
con.Execute(sSqlCmd)
Response.Clear

for each EntryNode in oNodDeatils.childNodes

	nEntryNo=EntryNode.getAttribute("EntryNo")
	iNetBasic=EntryNode.getAttribute("NettBasic")
	sItemDesc=replace(EntryNode.getAttribute("ItmDescription"),"'","''")
	dQty=EntryNode.getAttribute("Qty")
	sUOM=EntryNode.getAttribute("Uom")
	
	nAmount = EntryNode.getAttribute("ItemValue")

	dRate = EntryNode.getAttribute("Rate")
	iDisPer = EntryNode.getAttribute("DisPer")
	dDisAmount = EntryNode.getAttribute("DisAmount")
	iItemCode = EntryNode.getAttribute("ItemCode")
	iClassCode = EntryNode.getAttribute("ClassificationCode")

	nAccCode = EntryNode.getAttribute("ItemAccHead")
	sVat = EntryNode.getAttribute("VAT")
	sNarration=""

	'*********
	
	dRatePerQtyUoM = EntryNode.getAttribute("RatePerQtyUoM")
	'iItemBasicVal	= cdbl(dQty) * cdbl(dRatePerQtyUoM)
	iItemBasicVal	= cdbl(dQty) * cdbl(dRate)
	
	iItemValue		= EntryNode.getAttribute("ItemValue")
	iItemCalcRate	= EntryNode.getAttribute("ItemRate")
	sRateUoM		= EntryNode.getAttribute("RateUoM")
	iTempPurType	= EntryNode.getAttribute("PurchaseType")
	sAddDescr		= EntryNode.getAttribute("ItmDescription")
	sAttributeList  = EntryNode.getAttribute("AttributeList")

	if trim(iItemValue) = "" then iItemValue = "0"
	if trim(iItemCalcRate) = "" then iItemCalcRate = "0"

	'if CDbl(nAmount) <= 0  then nAmount = 0
	'if CDbl(iItemValue) <= 0  then iItemValue = 0
	'if CDbl(iItemCalcRate) <= 0  then iItemCalcRate = 0
	'Response.Clear
	Response.Write "["&dRatePerQtyUoM &"<BR>"&iItemCalcRate&"]"
	If iClassCode <> "TEMP" Then
	sSql = "Insert into RCV_T_InvoiceDetails(InvoiceNumber,EntryNo,OrganisationCode,ClassificationCode,ItemCode," &_
		   "InvItemUOM,InvoiceQuantity,InvoiceRate,ItemBasicValue,ItemDiscountPercent,ItemDiscountValue," &_
		    "ItemNettBasicValue,ItemValue,ItemRate,RateUoM,PurchaseType,AccountHead,VatEligibility,AdditionalDescr,ItemAttributes) values (" & nInvNo & "," & nEntryNo & ",'" & sOrgID & "'," & iClassCode & "," &_
		    iItemCode & ",'" & sUOM & "'," & dQty & "," & dRatePerQtyUoM & "," & iItemBasicVal & "," &_
		    iDisPer  & "," & dDisAmount  & "," & iNetBasic  & "," & iItemValue & "," & iItemCalcRate & ",'" & trim(sRateUoM) & "'," & iTempPurType & "," & nAccCode & ",'" & sVat & "',"& Pack(sAddDescr) &",'"& sAttributeList &"' )"
	Else
	sSql = "Insert into RCV_T_InvoiceDetails(InvoiceNumber,EntryNo,OrganisationCode,TempItemCode," &_
		   "InvItemUOM,InvoiceQuantity,InvoiceRate,ItemBasicValue,ItemDiscountPercent,ItemDiscountValue," &_
		    "ItemNettBasicValue,ItemValue,ItemRate,RateUoM,PurchaseType,AccountHead,VatEligibility,AdditionalDescr,ItemAttributes) values (" & nInvNo & "," & nEntryNo & ",'" & sOrgID & "'," &_
		    iItemCode & ",'" & sUOM & "'," & dQty & "," & dRatePerQtyUoM & "," & iItemBasicVal & "," &_
		    iDisPer  & "," & dDisAmount  & "," & iNetBasic  & "," & iItemValue & "," & iItemCalcRate & ",'" & trim(sRateUoM)  & "',"& iTempPurType & ","& nAccCode & ",'" & sVat & "',"& Pack(sAddDescr) &",'"& sAttributeList &"' )"
	End if
	Response.Write "<br><br>" + sSql
	' kk test
	con.execute sSql
	'**********

	
	sQuery="insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
	sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
	sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
	sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,ItemCode,ClassificationCode,InvoiceType,VATEligibility ) values ("
	sQuery=sQuery& iCrTransNo&",'"&sOrgId&"'"
	sQuery=sQuery& ","&nEntryNo&","&nAccCode&",NULL,NULL,NULL,"
	sQuery=sQuery&" '"&sNarration&"',"&nAmount&",'"&sEntryType&"',"&dQty&",'"&sItemDesc&"',"
	sQuery=sQuery&" '"&sUOM&"',"&dRatePerQtyUoM&","&iNetBasic&","&iDisPer&","&dDisAmount&"," & iItemCode & "," & iClassCode & "," & iTempPurType & ",'" & sVat & "')"

	Response.Write "<br><br>" + sQuery
	con.execute(sQuery)


	for each HeaderNode in EntryNode.childNodes
		if 	HeaderNode.nodeName="CostCenter" then
			for each  nodANL in HeaderNode.childNodes
				sAddCode=nodANL.Attributes.Item(0).nodeValue
				sAddRatio=nodANL.Attributes.Item(3).nodeValue
				sAddAmount=nodANL.Attributes.Item(4).nodeValue

				sQuery="INSERT INTO Acc_T_CreatedVoucherCCDet(CreatedTransNo, VoucherEntryNumber, AccountingUnit,"&_
					" AccUnitAccountHead,AccUnitCCHead,"&_
					"CCRatioPercent, CCRatioAmount)"&_
					" VALUES("& iCrTransNo& ","& sEntryno& ",'"&sOrgId&"',"&nAccCode&","&_
					" "& sAddCode &","& sAddRatio &"," & sAddAmount & ")"
				Response.Write "<br><br>" + sQuery
				con.execute(sQuery)

			next
		end if 'End of Check for Cost Center Node
		
		if 	HeaderNode.nodeName="Analytical" then
			for each  nodANL in HeaderNode.childNodes
				sAddCode=nodANL.Attributes.Item(0).nodeValue
				sAddRatio=nodANL.Attributes.Item(3).nodeValue
				sAddAmount=nodANL.Attributes.Item(4).nodeValue

				sQuery="INSERT INTO Acc_T_CretedVoucherAHDet(CreatedTransNo, VoucherEntryNumber, AccountingUnit, "&_
					"AccUnitAccountHead, AccUnitAnalyticalCode,"&_
					"RatioPercentage, RatioAmount)"&_
					" VALUES("& iCrTransNo& ","& sEntryno& ",'"&sOrgId&"',"&nAccCode&","&_
					""&sAddCode&","&sAddRatio&","&sAddAmount&")"
				Response.Write "<br><br>" + sQuery
				con.execute(sQuery)
			
			next
		end if 'End of Check for Analytical Node
	next 'End of Entry Node Loop
	'--------------Other details Pending---------------------

next 'End of Voucher Node Loop

iSno=1


'' from old program - start
''To insert Tax Form details
if blnTaxFlag = true then
	For Each oNodEntry in TaxFormRoot.childNodes

		iFormNo = oNodEntry.Attributes.Item(0).nodeValue
		iTaxCat = oNodEntry.Attributes.Item(1).nodeValue
		iTaxCode = oNodEntry.Attributes.Item(2).nodeValue
		sStatus = oNodEntry.Attributes.Item(3).nodeValue
		sExistForm  = oNodEntry.Attributes.Item(4).nodeValue
		sEntryFormNo  = oNodEntry.Attributes.Item(5).nodeValue
		iFormQty  = oNodEntry.Attributes.Item(6).nodeValue
		iUtilQty = oNodEntry.Attributes.Item(7).nodeValue
		iFormValue = oNodEntry.Attributes.Item(8).nodeValue
		iUtilValue = oNodEntry.Attributes.Item(9).nodeValue
		sValidTill = oNodEntry.Attributes.Item(10).nodeValue
		sRemarks = oNodEntry.Attributes.Item(11).nodeValue
		sSentOn = oNodEntry.Attributes.Item(12).nodeValue
		sUtilInvNo = oNodEntry.Attributes.Item(13).nodeValue

		If trim(iUtilQty) = "" then iUtilQty = "0"
		If trim(iUtilValue) = "" then iUtilValue = "0"
		''Util Qty/Value to be updated.....

		sSql = "Insert into PUR_T_ExemptionForms(FormNumber,InvoiceNumber,OUDefinitionID,TaxCode," &_
				"TaxCategoryCode,FormStatus,FormNo,QuantityInForm,ValueInForm,QuantityUtilised,ValueUtilised," &_
				"ValidUntil,SentOn,FormRemaks) values(" & iFormNo & "," & nInvNo & ",'" &  sOrgID & "'," &_
				iTaxCode & "," & iTaxCat & ",'" & sStatus & "','" & sEntryFormNo & "'," &  iFormQty & "," &_
				iFormValue  & "," & iUtilQty & "," & iUtilValue & ",convert(datetime,'" & sValidTill  & "',103)," &_
				"convert(datetime,'" & sSentOn & "',103),'" & sRemarks & "')"
		'Response.Write ssql
		'con.execute sSql

		'Update in Utilised form table if Form Status is "Existing Forms"
		If trim(sStatus) = "E" Then
			sSql = "Insert into PUR_T_FormsUtilised(FreshFormNumber,FreshInvoiceNumber,UtilisedFormNumber,UtilisedInvoiceNumber)" &_
				" Values(" & iFormNo & "," & trim(sUtilInvNo) & "," & sEntryFormNo & "," & nInvNo & ")"

			'Response.Write ssql
			con.execute sSql

			sSql = "Update PUR_T_ExemptionForms set QuantityUtilised = QuantityUtilised+" &  iFormQty & ", ValueUtilised = ValueUtilised+" &_
				iFormValue + " where FormNumber="  & iFormNo & " and InvoiceNumber=" & sUtilInvNo & " and TaxCode=" & iTaxCode & " and TaxCategoryCode=" & iTaxCat &_
				" and FormStatus='S' and FormNo='" & sEntryFormNo & "'"

			con.execute sSql

		End if 'If trim(sStatus) = "E" Then
	Next
End if 'if blnTaxFlag = true then

'' To Insert into Transport Table
If blnLRFlag = true then
	If LRRoot.hasChildNodes Then
		sTransportStage = LRRoot.attributes.item(0).nodevalue
		if trim(sTransportStage) <> "R" then
			For Each oNodEntry in LRRoot.childNodes
				sLRNum  = oNodEntry.attributes.item(0).nodevalue
				dtLRDate  = formatdate(oNodEntry.attributes.item(1).nodevalue)
				sVehicleNum = oNodEntry.attributes.item(2).nodevalue
				iLRGross  = oNodEntry.attributes.item(3).nodevalue
				iLRNett  = oNodEntry.attributes.item(4).nodevalue
				iLRPackCases = oNodEntry.attributes.item(5).nodevalue

				sSql = "Insert into RCV_T_TransportDetail (RefferenceNo,LRNumber,EntryStage," &_
				"LRDate,VehicleNumber,GrossWeight,NettWeight,PackingCases) Values(" & nInvNo & ",'" &_
				sLRNum  & "','" & sTransportStage & "',convert(datetime,'" & dtLRDate & "',103),'" & sVehicleNum & "'," & iLRGross &_
				"," & iLRNett & "," & iLRPackCases & ")"
				'Response.Write sSql
				Con.execute sSql
			Next
		end if 'if trim(sTransportStage) <> "R" then
	End if 'If LRRoot.hasChildNodes Then
End if 'If blnLRFlag = true then


''-------To Update Reference table ------------
If sInvAgainst = "1" Then	'Receipt
	''for multiple receipt details insert
	saRefNum = split(sRefNum,",")
	'Response.Write sRefNum
	for iCnt = 0 to ubound(saRefNum)
		sRefNum = saRefNum(iCnt)

	'' To check if invoice is already made for the receipt , if invoice exists to insert a new record for
	'' the receipt reference in REf. Num detail table changing the New Invoice No.alone
	'' else if no invoice is raised for the receipt to update the inv. no. against the receipt no

		Dim iRefInvNo
		sSql = "Select isnull(InvoiceNumber,0) from PUR_T_RefferenceNumberDet where ReceiptNumber=" & sRefNum  & ""
		'Response.Write sSql + vbcr
		With rsTemp
			.CursorLocation = 3
			.CursorType = 3
			.Source =  sSql
			.ActiveConnection = con
			.Open
		End With
		Set rsTemp.ActiveConnection = nothing

		if not rsTemp.EOF then
			iRefInvNo = rsTemp(0)
		end if
		rsTemp.Close

		if cint(iRefInvNo) = 0 then	'No Invoice has been raised for the receipt
			sSql = "Update PUR_T_RefferenceNumberDet Set InvoiceNumber=" & nInvNo &_
					 " where ReceiptNumber=" & sRefNum  & ""

		Else	' ALready Invoice is raised for the receipt

			sSql = "Insert into PUR_T_RefferenceNumberDet(OCNumber,GRNNumber,ReceiptNumber,InspectionNumber,InvoiceNumber) " &_
					" Select Distinct OCNumber,GRNNumber,ReceiptNumber,InspectionNumber," & nInvNo & " from PUR_T_RefferenceNumberDet " &_
					" where ReceiptNumber=" & sRefNum  & ""
		End if
	'	Response.Write sSql
		con.execute sSql
	next

ElseIf sInvAgainst = "2" Then	'Order
	sSql = "Update PUR_T_RefferenceNumberDet Set InvoiceNumber=" & nInvNo &_
			 " where PONumber=" & sRefNum  & ""
	con.execute sSql
ElseIf sInvAgainst = "0" Then	'Other
	sSql = "insert into PUR_T_RefferenceNumberDet (InvoiceNumber) values(" & nInvNo & ")"
	con.execute sSql
End If

'' from old program - end 


''To insert tax details
For each oNodTemp in oNodRoot.childNodes
	if oNodTemp.nodeName="TaxDetails" then
			set  oNodTaxRoot = oNodTemp
		
			'Response.Write "<p><p> oNodTaxRoot.NodeName = "  & oNodTaxRoot.NodeName
			iTempPurType	= oNodTaxRoot.getAttribute("PurchaseType")
			'Response.Write "<p><p> iTempPurType = " & trim(iTempPurType)
			

			For each EntryNode in oNodTaxRoot.childNodes

				iTaxCat=EntryNode.getAttribute("CatCode")
				iTaxCode=EntryNode.getAttribute("TaxCode")
				sTaxMode=EntryNode.getAttribute("TaxMode")
				dTaxPer=EntryNode.getAttribute("TaxValue")
				dTaxAmount=EntryNode.getAttribute("TaxAmount")
				nAccCode=EntryNode.getAttribute("AccHead")
				iTaxCrEligible = EntryNode.Attributes.Item(7).nodeValue
				iRegCode  = EntryNode.Attributes.Item(8).nodeValue	
				
				
				
				''Round off tax node not to be considered in insertion - 09 Aug 04
				If trim(iTaxCat) <> "0" and trim(iTaxCode) <> "0" then

					if trim(sTaxMode) <> "P" then dTaxPer = 0

					sSql = "Insert into RCV_T_InvoiceTaxDetails(InvoiceNumber,TaxCategoryCode,TaxCode,TaxCreditEligibility," &_
					"TaxAmount,TaxPercentage,PurchaseType) values(" & nInvNo & "," & iTaxCat  & "," & iTaxCode  & ",'" &  iTaxCrEligible  & "'," &_
					dTaxAmount & "," & dTaxPer & ","& iTempPurType & ")"
					Response.Write "<p><p> sSql = " &  sSql
					' kk test
					con.execute sSql

					''To insert into Modvat Credit Table
					If iTaxCrEligible = "1" Then

						'To fetch No. of credit years for the given Tax Register
						sSql = "Select NoOfYears from APP_R_ExciseRegisters Where RegisterCode=" & iRegCode & ""
						With rsTemp
							.CursorLocation = 3
							.CursorType = 3
							.Source =  sSql
							.ActiveConnection = con
							.Open
						End With
						Set rsTemp.ActiveConnection = nothing

						If not rsTemp.EOF then
							iRegYear = rsTemp(0)
						End If
						rsTemp.Close

						For iCnt = 1 to iRegYear
							'To fetch Tax percentage for each year of modvat credit taken
							sSql = "Select ApplicablePercentage from APP_R_ExciseRegisterPercentage Where RegisterCode=" & iRegCode & " and YearNumber=" & iCnt & ""
							With rsTemp
								.CursorLocation = 3
								.CursorType = 3
								.Source =  sSql
								.ActiveConnection = con
								.Open
							End With
							Set rsTemp.ActiveConnection = nothing

							If 	not rsTemp.EOF then
								iRegPercent = rsTemp(0)
							End If
							rsTemp.Close

							iModvatTaxAmt = iTaxAmt *(cdbl(iRegPercent)/100)

							''Insert into Modvat Credit Taken table for First Year and
							''in Modvat Pending Table for remaining years
							If iCnt = 1 then
								iForYear = Year(date())
								sSql = "Insert into RCV_T_ModvatCreditTaken(InvoiceNumber,TaxCategoryCode,TaxCode,RegisterCode," &_
									   "ForYear,TaxAmount) values(" & nInvNo & "," & iTaxCat & "," & iTaxCode & "," & iRegCode & "," &_
									   iForYear & "," & iModvatTaxAmt & ")"

								con.execute sSql
							Else
								iForYear = iForYear + 1
								sSql = "Insert into RCV_T_ModvatPending(InvoiceNumber,TaxCategoryCode,TaxCode,RegisterCode," &_
									   "TaxForYear,TaxPercent,TaxAmount) values(" & nInvNo & "," & iTaxCat & "," & iTaxCode & "," & iRegCode & "," &_
									   iForYear & "," & iRegPercent & "," & iModvatTaxAmt & ")"

								con.execute sSql
							End if

						Next
					End if
				End if ' If trim(iTaxCat) <> "0" and trim(iTaxCode) <> "0" then
				
				
				''To fetch Tax form No. if its available in Exemption forms
				sSql = "Select isnull(FormNumber,'') from Pur_T_ExemptionForms Where InvoiceNumber=" & trim(nInvNo) & "" &_
						" and OUDefinitionID = '" & trim(sOrgId) & "' and TaxCode=" & trim(iTaxCode) & " and TaxCategoryCode = " & iTaxCat & ""
				with rsTemp
					.CursorLocation = 3
					.CursorType = 3
					.Source = sSql
					.ActiveConnection = con
					.Open
				end with
				'Response.write sSql
				set rsTemp.ActiveConnection = nothing

				If not rsTemp.EOF then
					iFormNo = rsTemp(0)
				Else
					iFormNo = ""
				End if
				rsTemp.Close


			'	if CInt(nAccCode)>0 then
				
					if iTaxCat = "0" then	' for round off
							if dTaxAmount >=0 then
								sRCrtype = "D"
							else
								sRCrtype = "C"
							End if	
							sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
								"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
								""&iCrTransNo&","&nAccCode&",'"&sRCrtype&"',"&iSno&",NULL,NULL,"& iTempPurType & ",NULL,NULL,"&Abs(dTaxAmount)&")"
							Response.Write "<br><br>" + sQuery
							con.execute(sQuery)
				
							EntryNode.SetAttribute "TransAmt" , dTaxAmount
							
					else	
						' The C & D values are swapped on 15th July 2005 to be in sync with Accounts module Purchase voucher
						if dTaxAmount >=0 then
							sRCrtype = "D"
						else	'' For Tax Charges with negative value
							sRCrtype = "C"
							dTaxAmount = abs(dTaxAmount)
						End if
						
						if sTaxMode="P" then
							if trim(iFormNo) = "" or trim(iFormNo) = "0" then
								sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
									"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
									""&iCrTransNo&","&nAccCode&",'"&sRCrtype&"',"&iSno&","&iTaxCat&","&iTaxCode&","&iTempPurType&",NULL,"&dTaxPer&","&dTaxAmount&")"
							else
								sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
									"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
									""&iCrTransNo&","&nAccCode&",'"&sRCrtype&"',"&iSno&","&iTaxCat&","&iTaxCode&","&iTempPurType&"," & trim(iFormNo) & ","&dTaxPer&","&dTaxAmount&")"
							end if			
						else
							if trim(iFormNo) = "" or trim(iFormNo) = "0" then
								sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
									"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
									""&iCrTransNo&","&nAccCode&",'"&sRCrtype&"',"&iSno&","&iTaxCat&","&iTaxCode&","&iTempPurType&",NULL,NULL,"&dTaxAmount&")"
							else
								sQuery="INSERT INTO Acc_T_CreatedVoucherTaxDet(CreatedTransNo, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
									"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
									""&iCrTransNo&","&nAccCode&",'"&sRCrtype&"',"&iSno&","&iTaxCat&","&iTaxCode&","&iTempPurType&"," & trim(iFormNo) & ",NULL,"&dTaxAmount&")"			
							end if	
						end if

						Response.Write "<br><br>" + sQuery
						con.execute(sQuery)

					End if
				iSno=iSno + 1
			'	end if 'if CInt(nAccCode)>0 then
			Next

		
	end if 'if oNodTemp.nodeName="TaxDetails" then
	
next 'For each oNodTemp in oNodRoot.childNodes

dAdjAmtTotal=0
if blnAdvDetExist then
	For each EntryNode in oNodAdvRoot.childNodes
		iAdvTranNo=EntryNode.Attributes.Item(0).nodeValue
		dAdvAmount=EntryNode.Attributes.Item(5).nodeValue
		if CDbl(dAdvAmount)>0 then
			dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)

			sQuery="update Acc_T_AdvancePayments set AdvanceAdjusted=isnull(AdvanceAdjusted,0)+"&dAdvAmount &_
				" where TransactionNumber="&iAdvTranNo
			Response.Write "<br><br>" + sQuery
			con.execute(sQuery)
		end if
	Next
end if  'if blnAdvDetExist then

blnPayableFlag = false

If CDbl(iTotInvValue)>CDbl(dAdjAmtTotal) Then
	
	blnPayableFlag = True
	sQuery="select isnull(max(PayablesNumber),0)+1 from Acc_T_CreatedPayables"
	objRs.open sQuery,con
		iPayableNo=objRs(0)
	objRs.Close
	
	sPayNarration = "PUR INV for : " + cstr(sSuppInvNo)  + " , Date : " + cstr(dtSuppInvDate)
	
	sQuery="INSERT INTO Acc_T_CreatedPayables(PayablesNumber, CreatedTransNo, OUDefinitionID,"&_
			"VoucherDate, PartyType, PartySubType, PartyCode, PartyBillNumber,"&_
			" PartyBillDate, AmountPayable, AmountPaid,Narration)values("&iPayableNo&","&iCrTransNo&",'"&sOrgId&"',"&_
			"convert(datetime,'"&dtVoucDate&"',103),'"&sPartyType&"',"&iPartySubtype&","&iPartyCode&",'"&sSuppInvNo&"',"&_
			"convert(datetime,'"&dtSuppInvDate&"',103),"& iTotInvValue &","&dAdjAmtTotal&",'" & trim(sPayNarration) & "')"
	Response.Write "<br><br>" + sQuery
	con.execute(sQuery)
	
End if


''=====Inserting the Tax Packing Details ================
''Added by Ragav on Jun 14
For each oNodTemp in oNodRoot.childNodes
	if strcomp(oNodTemp.nodeName,"TaxDetails")=0 then
		for each ndTaxDet in oNodTemp.childNodes
			if strcomp(ndTaxDet.nodeName,"Tax")=0 then
				nOrgAccHdNo = ndTaxDet.getAttribute("OrgTaxAccHdNo")
				if ndTaxDet.hasChildNodes() then
					for each ndTaxPackDet in ndTaxDet.childNodes
						if ndTaxPackDet.nodeName="Taxpack" then
							nTaxPackCode = ndTaxPackDet.getAttribute("PackCode")
							nTaxPackRate = ndTaxPackDet.getAttribute("Packrate")
							sQuery = "Insert into RCV_T_InvoiceTaxPacDet (InvoiceNumber,OrgTaxAccHdNo,PackingCode,RatePerPack)"&_
									 "values ("&nInvNo&","&nOrgAccHdNo&","&nTaxPackCode&","&nTaxPackRate&")"
							Response.Write sQuery
							con.execute sQuery
						end if 'if StrComp(ndTaxPackDet.nodeName,"TaxPack")=0 then
					next
				end if 'if ndTaxDet.hasChildNodes() then
			end if
		next
	end if
Next
'========================================================





'' To set Main Purchase Inv accounting XML's attributes
Set newElem  = oDOM.createAttribute("CreatedTransNo")
newElem.value = iCrTransNo
oNodRoot.setAttributeNode(newElem)

Set newElem  = oDOM.createAttribute("CreatedVouNo")
newElem.value = sCrVouNo
oNodRoot.setAttributeNode(newElem)

Set newElem  = oDOM.createAttribute("TransNo")
newElem.value = iCrTransNo
oNodRoot.setAttributeNode(newElem)

Set newElem  = oDOM.createAttribute("VouNo")
newElem.value = sCrVouNo
oNodRoot.setAttributeNode(newElem)

iPurTranNum = iCrTransNo 

'-------- INSERTION FOR PURCHASE INVOICE ACCOUNTING OVER HERE ------------------------------------------


''To set accounted flag in invoice header
sQuery= " Update Rcv_T_Invoiceheader Set Accounted = 'Y' where InvoiceNumber=" & trim(nInvNo) & ""
Response.Write "<br><br>" + sQuery
con.execute(sQuery)
'-------------------------------------------------------------------------------------------------------

'----To code for Tax Form deatils----------------
'Response.End 
if con.Errors.count <>0 then
	con.RollbackTrans

	for iCounter=0 to con.Errors.count-1
		Response.Write con.Errors(iCounter).Description &"<br><br>"
		Response.Write con.Errors(iCounter).Source &"<br><br>"
	next
	'Redirect to Error Handling System
else


'	con.RollbackTrans
'	Response.End 
	
	Response.clear
	con.CommitTrans

	'storing acchead information in Xml
	oNodAccRoot.appendChild AccHeadElem
	
	'updating Invoice Number in A/c xml
	if oNodAccRoot.hasChildNodes then
		oNodAccRoot.setAttribute "CreatedTransNo",iCrTransNo
		oNodAccRoot.setAttribute "CreatedVouNo",sCrVouNo
		for each HeaderNode in oNodAccRoot.ChildNodes
			if HeaderNode.nodename = "Header" then
				For each EntryNode in HeaderNode.childNodes
					If EntryNode.nodename = "PurInvoice" then
						EntryNode.setAttribute "InvoiceNo",nInvNo
					end if 
				next
			end if 'if HeaderNode.nodename = "Header" then
		next
	end if 'if oNodAccRoot.hasChildNodes then
			
	'oDOM.Save server.MapPath("../../Accounts/xmldata/Voucher/"&iPurTranNum&".xml")
	'oDOMAccData.save server.MapPath("../../Accounts/xmldata/Voucher_XML/"&iPurTranNum&".xml")
		
	
	''---------------------------------------------------------------------
	
	If objFS.FileExists(server.MapPath("../temp/transaction/AmdNewInvItemValue_PUR_"&Session.SessionID&".xml")) Then 	
		objFS.deletefile(Server.MapPath("../temp/transaction/AmdNewInvItemValue_PUR_"&Session.SessionID&".xml"))
	End if
	
	Response.Redirect "CommonMessage.asp?Title=Receipt Status&Heading=Receipt Invoice Amendment Completed <br><br>" & nInvNo & " &Redirect=AMDINVOICENOSELECTION.ASP"
end if

%>

<%
Function getAccountHeadName(iAccHead)

Dim objAcc,sSql
Set objAcc = Server.CreateObject("Adodb.Recordset")

sSql = "Select AccountDescription,AccountHeadCode from Acc_M_GLAccountHead where AccountHead=" & iAccHead & ""
with objAcc
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSql
	.ActiveConnection = con
	.Open
end with
set objAcc.ActiveConnection = nothing

If not objAcc.EOF then
	getAccountHeadName = objAcc(0)
Else
	getAccountHeadName = iAccHead  
End if 
objAcc.Close  

End function
%>
