<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AppOtherSALUpdate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	September  30, 2003
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
<%
'XML DOM Variables
Dim oDOM,oNodRoot,oNodDeatils,oNodEntry,objRs,oNodAgent,oNodTaxRoot,oNodAdvRoot,oNodTemp,sQuery
dim sParCode,sParSubType,sParType,sParName
dim EntryNode,HeaderNode,nodANL,newElem
dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
dim sOrgId,sBookNo,sVouType,iVouNo,sVouName
dim sVouCode,sApprove,sVoucDate,sAccUnit
dim dTotal,sTransType,dCRAmt,dDRAmt
dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo
dim sDocType,sVouStatus
dim iSeriesNo,iSeriesCode
dim sSalType,sRefernceNo,iCounter
dim iPayableNo,iCatCode,iTaxCode,dTaxPer,dTaxAmount,sSalInvNo,sSalInvDt
dim iCrTransNo,sCrVouNo,iCrSeriesNo,iCrSeriesCode,sBookName
dim sMessage,sMessage1,dRatePer,sInvType,iSaleno,sItemcode,sClasscode
dim dAdjAmtTotal,sNewItemDesc,sNewNarr,sVouEntTy,iNoPacks,sAgType,TaxEntryNode
dim dtAccDate,sPara,sTaxAccCode
Dim iAdjRecNo,sAdjType,sExp,CheckNode,iCrAdvNo,iCrPayableNo

sVouStatus="010104" 'Approved and Accounted
sVouCode="05"
sVouType="D"
sTransType="SJR"
sAccUnit = Request.Form("selUnitId")
sInvType = Request("hExpInv")
dtAccDate = Request("hAccDate")
sPara = Request("hPara")
If trim(sPara) = "App" then sVouStatus = "010103"
IF CStr(sAccUnit) = "" Then
	sAccUnit = sOrgId
	sAccUnit = session("organizationcode")
End IF


set objRs  = server.CreateObject("adodb.recordset")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

iCrTransNo=Request("hTransNo")
sBookNo=Request("selBook")
sBookName=Request("hBookName")

con.BeginTrans

oDOM.load  server.MapPath("../temp/transaction/Voucher Entry_OthSal_"&Session.SessionID&".xml")

If trim(sPara) = "Acc" or trim(sPara) = "" then 'For Accounting --condition Added by Maheswari on July 1st 2008

	sQuery = "Select CreatedTransNo From Acc_T_VoucherHeader Where CreatedTransNo = "&iCrTransNo&" "
	objRs.Open sQuery,Con
	IF Not objRs.EOF Then
		sVouEntTy = "Y" 'Already Accounted
	Else
		sVouEntTy = "N" 'Can Be Accounted
	End IF
	objRs.Close
'Response.Write "<p> sVouEntTy= " & sVouEntTy
	IF Cstr(sVouEntTy) = "N" Then

		sQuery = "Select isNull(NoOfPack,0) From Acc_T_CreatedVoucherDetails Where CreatedTransNo = "&iCrTransNo&" "
		objRs.Open sQuery,Con
		IF Not objRs.EOF Then
			iNoPacks = objRs(0)
		End IF
		objRs.Close



		sQuery = "Select OtherApplnTransNo From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&iCrTransNo&" "
		Objrs.Open sQuery,Con
		IF Not Objrs.Eof Then
			iSaleno = Objrs(0)
		Else
			iSaleno = 0
		End IF
		Objrs.close

		IF Cstr(sInvType) = "Y" Then



			sQuery = "Delete From Acc_T_CreatedReceivables WHere CreatedTransNo = "&iCrTransNo
			Con.Execute sQuery

			sQuery = "Delete From Acc_T_CreatedVoucherTaxDet WHere CreatedTransNo = "&iCrTransNo
			Con.Execute sQuery


			sQuery = "Delete From Acc_T_CreatedVoucherDetails WHere CreatedTransNo = "&iCrTransNo
			Con.Execute sQuery

			sQuery = "Delete From Acc_T_CreatedVoucherHeader WHere CreatedTransNo = "&iCrTransNo
			Con.Execute sQuery





		End IF


		set oNodRoot=oDOM.documentElement

		for each oNodTemp in oNodRoot.childNodes
			if oNodTemp.nodeName="Header" then
				for Each oNodEntry in  oNodTemp.childNodes
					if oNodEntry.nodeName="Organization" then
						sOrgId=oNodEntry.Attributes.Item(0).nodeValue
					end if
					if oNodEntry.nodeName="Book" then
						oNodEntry.Attributes.Item(0).nodeValue=sBookNo
						oNodEntry.Text=sBookName
					end if
					if oNodEntry.nodeName="SalesType" then
						sSalType=oNodEntry.Attributes.Item(0).nodeValue
					end if
					if oNodEntry.nodeName="Party" then
						sParType=oNodEntry.Attributes.Item(0).nodeValue
						sParSubType=oNodEntry.Attributes.Item(1).nodeValue
						sParCode=oNodEntry.Attributes.Item(3).nodeValue
						sAgType = oNodEntry.Attributes.Item(4).nodeValue
						sParName=oNodEntry.text
					end if
					if oNodEntry.nodeName="SaleInvoice" then
						sSalInvNo=oNodEntry.Attributes.Item(0).nodeValue
						sSalInvDt=oNodEntry.Attributes.Item(1).nodeValue
						sRefernceNo=oNodEntry.Attributes.Item(2).nodeValue
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
				dTotal=oNodTemp.Attributes.Item(0).nodeValue
			end if
			if oNodTemp.nodeName="AdvanceDetails" then
				set oNodAdvRoot=oNodTemp
			end if
		next
		oNodRoot.setAttribute "TransNo",iTransNo


		sNewNarr = "Sale Inv No:"&sSalInvNo&" Dt:"&sSalInvDt

		'Dim sExp,TempNode
		'sExp = "//AdvanceDetails"
		'Set TempNode = oNodRoot.selectNodes(sExp)
		'Response.Write TempNode.length
		'Response.End
		'Set oNodAdvRoot = TempNode.Item(0)

		sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue



		sQuery="select isnull(max(TransactionNumber),0)+1 from Acc_T_VoucherHeader"
		objRs.open sQuery,con
			iTransNo=objRs(0)
		objRs.Close

		IF Cstr(sInvType) = "Y" Then
			sQuery = "Insert into Acc_T_CreatedVoucherHeader (CreatedTransNo,OUDefinitionID,BookCode,BookNumber,TransactionType,"
			sQuery = sQuery&"PartyType,PartySubType,PartyCode,AccountHead,CreatedVoucherNo,VoucherDate,VoucherAmount,"
			sQuery = sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,CreatedVouchStatus,FromApplication,OtherApplnTransNo,OtherApplnTableName) values"
			sQuery = sQuery&"("&iCrTransNo&",'"&sOrgId&"','"&sVouCode&"',NULL,'"&sTransType&"',"
			sQuery = sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sSalInvNo&"',convert(datetime,'"&sSalInvDt&"',103),"&dTotal
			sQuery = sQuery&",'"&sRefernceNo&"','"&sSalType&"','"&sVouType&"',"&getUserid&",getdate(),NULL,'"&sVouStatus&"','3',"&iSaleno&",'Sal_T_Invoiceheader')"

			Response.Write sQuery& "<BR><BR>"
			con.execute(sQuery)
		End IF


		sQuery=" insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"
		sQuery=sQuery&"PartyType,PartySubType,PartyCode,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"
		sQuery=sQuery&"PayToRecdFrom,BankInstrumentType,CrDrIndication,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus) values"
		sQuery=sQuery&"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"
		sQuery=sQuery&"'"&sParType&"',"&sParSubType&","&sParCode&",NULL,'"&sSalInvNo&"','"&sSalInvNo&"',"&iCrTransNo&",convert(datetime,'"&sSalInvDt&"',103),"&dTotal
		sQuery=sQuery&",'"&sParName&"','"&sSalType&"','"&sVouType&"',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"')"

		Response.Write sQuery &"<br><br>"
		con.execute(sQuery)

		'sAmount = oNodDeatils.Attributes.getNamedItem("ActualValue").Value
		for each EntryNode in oNodDeatils.childNodes
		dim dQty,sUOM,dBasicAmount,dRate,dDisPer,dDisAmount,sItemDesc,dItem,dClass

			sEntryno=EntryNode.Attributes.Item(0).nodeValue
			sAmount=EntryNode.Attributes.Item(2).nodeValue
			sItemDesc=Replace(EntryNode.Attributes.Item(1).nodeValue,"'","''")
			dQty=EntryNode.Attributes.Item(3).nodeValue
			sUOM=EntryNode.Attributes.Item(4).nodeValue
			dBasicAmount=EntryNode.Attributes.Item(7).nodeValue
			dRate=EntryNode.Attributes.Item(6).nodeValue
			dDisPer=EntryNode.Attributes.Item(8).nodeValue
			dDisAmount=EntryNode.Attributes.Item(9).nodeValue
			dRatePer = EntryNode.Attributes.Item(16).nodeValue
			dItem = EntryNode.Attributes.Item(10).nodeValue
			dClass = EntryNode.Attributes.Item(11).nodeValue
			
			''blocked by ragav on Dec 08,2011 for updating the second time the discout value so blocked
			'sAmount = cdbl(sAmount) - cdbl(dDisAmount)	'Added By UmaMaheswari S,On July 30,2011
			''end 

			sEntryType="C"
			sNarration=""

			for each HeaderNode in EntryNode.childNodes
				if HeaderNode.nodeName="AccHead" then
						sAccCode=HeaderNode.Attributes.Item(0).nodeValue
						sAccType=HeaderNode.Attributes.Item(4).nodeValue
				end if 'End of Check for Account head Node
			next 'End of Entry Node Loop

			sQuery = "Select M.PackingName,Count(S.PackingCode) From Sal_T_InvPackDetails S, "&_
					 "APP_M_PackingType M Where S.PackingCode = M.PackingCode and  "&_
					 "S.SaleTransactionNo = "&iSaleno&" and ItemCode = "&dItem&" and ClassificationCode = "&dClass&"  "&_
					 "Group By M.PackingName "

			Response.Write sQuery &"<br><br><br>"

			objRs.Open sQuery,con
			Do While Not objRs.EOF
				sNewItemDesc = sNewItemDesc & objRs(1) &" " & objRs(0) &" "
				objRs.MoveNext
			Loop
			objRs.Close



			'sNewItemDesc = sItemDesc&" - "&sNewItemDesc
			sNewItemDesc = sItemDesc 
			Response.Write sNewItemDesc &"<br><br>"


				IF Cstr(sInvType) = "Y" Then
					sQuery = "Insert into Acc_T_CreatedVoucherDetails (CreatedTransNo,AccountingUnit,"
					sQuery = sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
					sQuery = sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
					sQuery = sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,Itemcode,ClassificationCode,RatePer,NoOfPack ) values ("
					sQuery = sQuery& iCrTransNo&",'"&sOrgId&"'"
					sQuery = sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
					sQuery = sQuery&" '"&sNewItemDesc&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sNewItemDesc&"',"
					sQuery = sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&","&dItem&","&dClass&", "&dRatePer&","&iNoPacks&")"
					Response.Write sQuery& "<BR><BR>"
					con.execute(sQuery)
				End IF
				
				if CDbl(sEntryno)=1 then
		            if oNodTaxRoot.hasChildNodes() then
	                    for each TaxEntryNode in oNodTaxRoot.childNodes

			                iCatCode=TaxEntryNode.Attributes.Item(0).nodeValue
			                iTaxCode=TaxEntryNode.Attributes.Item(1).nodeValue
			                sTaxMode=TaxEntryNode.Attributes.Item(2).nodeValue
			                dTaxPer=TaxEntryNode.Attributes.Item(4).nodeValue
			                dTaxAmount=TaxEntryNode.Attributes.Item(5).nodeValue
			                sTaxAccCode=TaxEntryNode.Attributes.Item(6).nodeValue

			                IF Cstr(dTaxAmount) = "" Then
				                dTaxAmount = 0
			                End IF
			                if CDbl(sTaxAccCode)=0 then
		                        sAmount = cdbl(sAmount) + cdbl(dTaxAmount)
		                    end if
		                Next
				    end if 'if oNodTaxRoot.hasChildNodes() then
				end if 'if CDbl(sEntryno)=1 then
			
				sQuery="insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
				sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
				sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,InvoicedQuantity,ItemDescription,"
				sQuery=sQuery&" InvoicedUoM,InvoicedRate,BasicAmount,DiscountPercent,DiscountAmount,RatePer,ItemCode,ClassificationCode,NoOfPack ) values ("
				sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
				sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
				sQuery=sQuery&" '"&sNewItemDesc&"',"&sAmount&",'"&sEntryType&"',"&dQty&",'"&sNewItemDesc&"',"
				sQuery=sQuery&" '"&sUOM&"',"&dRate&","&dBasicAmount&","&dDisPer&","&dDisAmount&", "&dRatePer&", "&dItem&", "&dClass&","&iNoPacks&" )"

		Response.Write sQuery &"<br><br>"
		con.execute(sQuery)
		'Response.End 

		dim sCCGroup,sAddCode,sAddRatio,sAddAmount

		for each HeaderNode in EntryNode.childNodes
			if 	HeaderNode.nodeName="CostCenter" then
				for each  nodANL in HeaderNode.childNodes
					sAddCode=nodANL.Attributes.Item(0).nodeValue
					sAddRatio=nodANL.Attributes.Item(3).nodeValue
					sAddAmount=nodANL.Attributes.Item(4).nodeValue

					sQuery="INSERT INTO Acc_T_CCVoucherDet(TransactionNumber, VoucherEntryNumber, AccountingUnit,"&_
						" AccUnitAccountHead,AccUnitCCHead,"&_
						"CCRatioPercent, CCRatioAmount)"&_
						" VALUES("& iTransNo& ","& sEntryno& ",'"&sOrgId&"',"&sAccCode&","&_
						" "& sAddCode &","& sAddRatio &"," & sAddAmount & ")"
					'Response.Write sQuery& "<BR>"
					con.execute(sQuery)
				next
			end if 'End of Check for Cost Center Node
			if 	HeaderNode.nodeName="Analytical" then
				for each  nodANL in HeaderNode.childNodes
					sAddCode=nodANL.Attributes.Item(0).nodeValue
					sAddRatio=nodANL.Attributes.Item(3).nodeValue
					sAddAmount=nodANL.Attributes.Item(4).nodeValue

					sQuery="INSERT INTO Acc_T_AnalyticalVoucherDet(TransactionNumber, VoucherEntryNumber, AccountingUnit, "&_
						"AccUnitAccountHead, AccUnitAnalyticalCode,"&_
						"RatioPercentage, RatioAmount)"&_
						" VALUES("& iTransNo& ","& sEntryno& ",'"&sOrgId&"',"&sAccCode&","&_
						""&sAddCode&","&sAddRatio&","&sAddAmount&")"
					'Response.Write sQuery& "<BR>"
					con.execute(sQuery)
				next
			end if 'End of Check for Analytical Node
		next 'End of Entry Node Loop

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

			IF Cstr(dTaxAmount) = "" Then
				dTaxAmount = 0
			End IF

		Response.write dTaxAmount &"<br>"

		if 	CDbl(dTaxAmount) >0 then
			sEntryType="C"
		else
			dTaxAmount=CDbl(dTaxAmount)*-1
			sEntryType="D"
		end if



		if sTaxMode="P" then
			IF Cstr(sInvType) = "Y" Then
				sQuery = "INSERT INTO Acc_T_CreatedVoucherTaxDet (CreatedTransNo, AccountHead, TransCrDrIndication,  "&_
						 "TaxEntryNo, TaxCategoryCode, TaxCode, InvoiceType, FormNumber, TaxPercentage, TaxAmount) "&_
						 "VALUES ("&iCrTransNo&", "&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,"&dTaxPer&","&dTaxAmount&")"

				Response.Write sQuery &"<br><br>"
				Con.Execute sQuery
			End IF
		else
			IF Cstr(sInvType) = "Y" Then
				sQuery = "INSERT INTO Acc_T_CreatedVoucherTaxDet (CreatedTransNo, AccountHead, TransCrDrIndication,  "&_
						 "TaxEntryNo, TaxCategoryCode, TaxCode, InvoiceType, FormNumber, TaxPercentage, TaxAmount) "&_
						 "VALUES ("&iCrTransNo&", "&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,NULL,"&dTaxAmount&")"

				Response.Write sQuery &"<br><br>"
				Con.Execute sQuery
			End IF
		end if
	    if Trim(sAccCode)>0 then
			if sTaxMode="P" then
				sQuery="INSERT INTO Acc_T_VoucherTaxDetails(TransactionNumber, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
					"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
					""&iTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,"&dTaxPer&","&dTaxAmount&")"
			else
				
				sQuery="INSERT INTO Acc_T_VoucherTaxDetails(TransactionNumber, AccountHead, TransCrDrIndication,TaxEntryNo,"&_
					"TaxCategoryCode,TaxCode,InvoiceType, FormNumber, TaxPercentage, TaxAmount) VALUES("&_
					""&iTransNo&","&sAccCode&",'"&sEntryType&"',"&iSno&","&iCatCode&","&iTaxCode&","&sSalType&",NULL,NULL,"&dTaxAmount&")"
			end if
		    Response.Write sQuery & "<br><br>"
		    con.execute(sQuery)
		 end if 'if Trim(sAccCode)>0 then
		 
		iSno=cint(iSno)+1
		Next

		'Response.Clear
		dim iAdvTranNo,dAdvAmount,iAdvNo,iAdvCrNo,iAdvTrCrNo
		dAdjAmtTotal=0
		'if  IsObject(oNodAdvRoot) then
			for each EntryNode in oNodAdvRoot.childNodes
				iAdvTranNo = EntryNode.Attributes.Item(0).nodeValue
				dAdvAmount = EntryNode.Attributes.Item(5).nodeValue
				iAdvTrCrNo = EntryNode.Attributes.Item(6).nodeValue
				iAdvNo = EntryNode.Attributes.Item(7).nodeValue
				dAdjAmtTotal = dAdjAmtTotal+CDbl(dAdvAmount)

				'IF CStr(sInvType) = "Y" Then
					sQuery = "Select CreatedAdvanceNo From Acc_T_AdvancePayments Where AdvanceNumber = "&iAdvNo&" "
					objRs.Open sQuery,Con
					IF Not objRs.EOF Then
						iAdvCrNo = objRs(0)
					Else
						iAdvCrNo = 0
					End IF
					objRs.Close

					sQuery = "UPDATE ACC_T_CREATEDADVANCES SET ADVANCEADJUSTED=ISNULL(ADVANCEADJUSTED,0)+"&dAdvAmount &_
							 " WHERE CREATEDTRANSNO = "&iAdvTrCrNo&" AND CREATEDADVANCENO = "&iAdvCrNo&" "

					'Response.Write sQuery &"<br><br>"
					Con.Execute sQuery

				'End IF

				sQuery="update Acc_T_AdvancePayments set AdvanceAdjusted="&dAdvAmount &_
				" where TransactionNumber="&iAdvTranNo

				Response.Write sQuery &"<br><br>"
				con.execute(sQuery)
			Next
		'end if
		dim iCrPayable,sPayNarration

		if CDbl(dTotal)>CDbl(dAdjAmtTotal) then

			sQuery="select isnull(max(ReceivableNumber),0)+1 from Acc_T_CreatedReceivables"
			objRs.open sQuery,con
				iPayableNo=objRs(0)
			objRs.Close

			'unblocked on feb 18,2010 - by kalaiselvi
			IF Cstr(sInvType) = "Y" Then
				sQuery = "INSERT INTO Acc_T_CreatedReceivables (ReceivableNumber, CreatedTransNo, OUDefinitionID,  "&_
						 "VoucherDate, PartyType, PartySubType, PartyCode, PartyInvoiceNumber, PartyInvoiceDate,  "&_
						 "AmountReceivable, AmountReceived,Narration) "&_
						 "VALUES ("&iPayableNo&", "&iCrTransNo&", '"&sOrgId&"', convert(datetime,'"&sVoucDate&"',103), "&_
						 "'"&sParType&"',"&sParSubType&","&sParCode&", '"&sSalInvNo&"', convert(datetime,'"&sSalInvDt&"',103),"&dTotal&","&dAdjAmtTotal&",'"&sNewNarr&"') "

				Response.Write sQuery &"<br><br>"
				con.execute(sQuery)
			End IF


			sQuery="select ReceivableNumber,Narration from Acc_T_CreatedReceivables where CreatedTransNo="&iCrTransNo
			objRs.open sQuery,con
			if not objRs.eof then
				iCrPayable=objRs(0)
				sPayNarration=objRs(1)
			end if
			objRs.Close


			sQuery="select isnull(max(ReceivableNumber),0)+1 from Acc_T_Receivables"
			objRs.open sQuery,con
			if not objRs.eof then
				iPayableNo=objRs(0)
			end if 
			objRs.Close

			sQuery="INSERT INTO Acc_T_Receivables(ReceivableNumber,TransactionNumber, OUDefinitionID, "&_
					"CreatedReceivable,VoucherDate, PartyType, PartySubType, PartyCode,PartyInvoiceNumber, "&_
					"PartyInvoiceDate, AmountReceivable, AmountReceived,Narration) values("&iPayableNo&","&iTransNo&",'"&sOrgId&"',"&_
					""&iCrPayable&",convert(datetime,'"&sVoucDate&"',103),'"&sParType&"',"&sParSubType&","&sParCode&",'"&sSalInvNo&"',"&_
					"convert(datetime,'"&sSalInvDt&"',103),"&dTotal&","&dAdjAmtTotal&",'"&sPayNarration&"')"

			Response.Write sQuery& "<BR>"
			con.execute(sQuery)

		end if

		'Response.Write "<BR><BR>"

	'========= This is been done at top itself ==================================================
		'dim iAgentCode,sCommType,dCommPer,dCommAmount,iCurrencyCode
		'if  IsObject(oNodAgent) then
		'	for each EntryNode in oNodAgent.childNodes
		'		iAgentCode=EntryNode.Attributes.Item(0).nodeValue
		'		sCommType=EntryNode.Attributes.Item(2).nodeValue
		'		dCommPer=EntryNode.Attributes.Item(3).nodeValue
		'		dCommAmount=EntryNode.Attributes.Item(4).nodeValue
		'		sQuery="INSERT INTO Sal_T_AdditionalAgents(AgentCode, CommissionType,CurrencyCode, "&_
		'				"AccTransactionNo, AgentCommission,CommissionToPay) values ("&iAgentCode&",'"&sCommType&"',1,"&_
		'				""&iTransNo&","&dCommAmount&",'1')"
		'	con.execute(sQuery)
		'	Next
		'end if

		IF CStr(sAgType) = "Y" Then
			sQuery = "INSERT INTO Sal_T_AdditionalAgents (AgentCode, CommissionType, AgentCommission, "&_
					 "CurrencyCode, AgentType, AgentSubType, AccTransactionNo) "&_
					 "Select  AgentCode, CommissionType, AgentCommission, CurrencyCode, "&_
					 "AgentType, AgentSubType, "&iTransNo&" From Sal_T_AdditionalAgents "&_
					 "Where SaleTransactionNo = "&iSaleno&" "

			Response.Write sQuery &"<br><br>"
			Con.Execute sQuery

		End IF



		sQuery="update Acc_T_CreatedVoucherHeader set BookNumber="&sBookNo&","&_
		" CreatedVouchStatus='"&sVouStatus&"' where CreatedTransNo="&iCrTransNo
		con.execute(sQuery)

		sQuery = "Update Acc_T_PartyTransactions Set VoucherNarration = '"&sNewItemDesc&"' Where  "&_
				 "TransactionNumber = "&iTransNo&" "
		con.execute(sQuery)

		sQuery = "Update Acc_T_GLTransactions Set VoucherNarration = '"&sNewItemDesc&"' Where  "&_
				 "TransactionNumber = "&iTransNo&" "
		con.execute(sQuery)


		'================== Checking for Debit and Credit Tally ===================================
		'Dim dDebTotal,dCreTotal
		'sQuery = "Select VoucherAmount,CrDrIndication From Acc_T_VoucherHeader  "&_
		'		 "Where TransactionNumber = "&iTransNo&" "
		'Objrs.Open sQuery,Con
		'IF Not Objrs.Eof Then
		'	IF Cstr(Objrs(1)) = "C" Then
		'		dCreTotal = Cdbl(Objrs(0))
		'	Else
		'		dDebTotal = Cdbl(Objrs(0))
		'	End IF
		'End IF
		'objrs.Close

		'sQuery = "Select Sum(Amount),TransCrDrIndication From Acc_T_VoucherDetails Where "&_
		'		 "TransactionNumber = "&iTransNo&" Group By TransCrDrIndication "

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

		'sQuery = "Select Sum(TaxAmount),TransCrDrIndication From Acc_T_VoucherTaxDetails Where "&_
		'		 "TransactionNumber = "&iTransNo&" Group By TransCrDrIndication "
		''Response.Write "<p>"&sQuery&"<BR>"
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

		'Response.Write "=============Differences=========<BR>"

		'IF Cdbl(dCreTotal) <> Cdbl(dDebTotal) Then
		'	'Response.Clear
		'	Response.Write "Total Credit and Debit Does'nt Match Transaction is been Deleted "
		'	Response.Write "<br>"
		'	Response.Write "Debit Total " & FormatNumber(dDebTotal,2,,,0) &" Credit Total " & FormatNumber(dCreTotal,2,,,0)
		'	Con.RollbackTrans
		'	Response.End
		'End IF

		'================ uPDATED bY mANOHAR
		Dim sTbGLVal,sTempVal,sDiffVal,dTBDRAmt,dTBCRAmt
		sTbGLVal = CheckTBGL(sOrgId)
		sTempVal = Split(sTbGLVal,":")
		dTBDRAmt = sTempVal(0)
		dTBCRAmt = sTempVal(1)

		dTBDRAmt = FormatNumber(dTBDRAmt,2,,,0)
		dTBCRAmt = FormatNumber(dTBCRAmt,2,,,0)

		IF CStr(Trim(sTempVal(2))) = "T" Then
			sDiffVal = "In Trial Balance "
		Else
			sDiffVal = "In Ledger  "
		End IF


		IF CDbl(dTBDRAmt) <> CDbl(dTBCRAmt) Then
			con.RollbackTrans


		'	Response.Clear
			Response.Write "<h3>Debit and Credit is Not Matching </h3><br>"
			Response.Write "<b>Debit Amount  :--> </b>" & dTBDRAmt &"<br>"
			Response.Write "<b>Credit Amount :--> </b>" & dTBCRAmt &"<br>"
			Response.Write "<b>Differences   :--> </b>" & sDiffVal &"<br>"
			Response.Write "<h3> Voucher Not Created " &"<br>"
			Response.End
		End IF



	'================== Checking for Debit and Credit Tally ===================================
	End IF
End IF 'If trim(sPara) = "Acc" or If trim(sPara) = "" then

 Response.Write "<P>"&sPara&"<BR>"
If trim(sPara) = "Edt" then
	sQuery="update Acc_T_CreatedVoucherHeader set BookNumber="&sBookNo&" where CreatedTransNo="&iCrTransNo
			Response.Write sQuery &"<br>"
			con.execute(sQuery)
ElseIF trim(sPara) = "App" then

set oNodRoot=oDOM.documentElement
sCrVouNo = oNodRoot.Attributes.Item(1).nodeValue
		for each oNodTemp in oNodRoot.childNodes
   			if oNodTemp.nodeName="Header" then
				for Each oNodEntry in  oNodTemp.childNodes
					if oNodEntry.nodeName="Party" then
						sParType=oNodEntry.Attributes.Item(0).nodeValue
						sParSubType=oNodEntry.Attributes.Item(1).nodeValue
						sParCode=oNodEntry.Attributes.Item(3).nodeValue
						sAgType = oNodEntry.Attributes.Item(4).nodeValue
						sParName=oNodEntry.text
					end if
					if oNodEntry.nodeName="SaleInvoice" then
						sSalInvNo=oNodEntry.Attributes.Item(0).nodeValue
						sSalInvDt=oNodEntry.Attributes.Item(1).nodeValue
						sRefernceNo=oNodEntry.Attributes.Item(2).nodeValue
					end if
				next
		    end if
		    if oNodTemp.nodename="Details" then
		        set oNodDeatils = oNodTemp 
		        sVoucDate=oNodDeatils.Attributes.Item(3).nodeValue
		    end if
		    if oNodTemp.nodeName="TaxDetails" then
		        set oNodTaxRoot = oNodTemp 
		        dTotal=oNodTemp.Attributes.Item(0).nodeValue
		    end if
			if oNodTemp.nodeName="AdvanceDetails" then
				set oNodAdvRoot=oNodTemp
			end if
		next

    if oNodAdvRoot.hasChildNodes() then
    	for each EntryNode in oNodAdvRoot.childNodes
			iAdvTranNo = EntryNode.Attributes.Item(0).nodeValue
			dAdvAmount = EntryNode.Attributes.Item(5).nodeValue
			iAdjRecNo = EntryNode.Attributes.Item(6).nodeValue
			sAdjType = EntryNode.Attributes.Item(8).nodeValue

			sExp = "//AdvanceDetails[@TransNo="&iAdvTranNo&" and @AdvNo]"
			Set CheckNode = oNodAdvRoot.selectNodes(sExp)

			IF CheckNode.length <> 0 Then
				iAdvNo = EntryNode.Attributes.Item(5).nodeValue
			Else
				iAdvNo = 0
			End IF

			IF CStr(iAdvNo) = "0" and CStr(sAdjType) = "I" Then
				sQuery = "Select AdvanceNumber From Acc_T_AdvancePayments Where TransactionNumber = "&iAdvTranNo&" "

				Response.Write sQuery &"<br><br>"
				objRs.Open sQuery,con
				IF Not objRs.EOF Then
					iAdvNo = objRs(0)
				End IF
				objRs.Close
			End IF
		'=====================================================================================

		if CDbl(dAdvAmount)>0 then
			dAdjAmtTotal=dAdjAmtTotal+CDbl(dAdvAmount)
		end if
		    Dim sAdjNarr
		    sAdjNarr = "Adjusted to Sales Invoice No "&sCrVouNo&" DT: "& sVoucDate &" Amt: "& dTotal
		
			IF CStr(sAdjType) = "I" and CDbl(dAdvAmount) <> 0 Then
				sQuery = "update Acc_T_AdvancePayments set AdvanceAdjusted="&dAdvAmount &" where AdvanceNumber = "&iAdvNo

				Response.Write sQuery &"<br><br>"
				con.execute(sQuery)

				sQuery = "Select CreatedAdvanceNo From Acc_T_AdvancePayments Where AdvanceNumber = "&iAdvNo
				Objrs.Open sQuery,Con
				IF Not Objrs.Eof Then
					iCrADvNo = Objrs(0)
				End IF
				Objrs.close

				sQuery = "Update Acc_T_CreatedAdvances Set AdvanceAdjusted = "&dAdvAmount&" Where CreatedAdvanceNo = "&iCrAdvNo
				    Response.Write "<p>"& sQuery
				con.execute(sQuery)
				
				sQuery = "Update Acc_T_AdvancePayments set AdvanceAdjusted = "& dAdvAmount &" where CreatedAdvanceNo = "& iCrAdvNo &" and CreatedTransNo = "& iCrTransNo
				    Response.Write "<p>"& sQuery
				con.execute(sQuery)
				
				sQuery = " Select CreatedAdvanceNo from ACC_T_CREATEDADVANCEADJ where CreatedAdvanceNo = "& iCrAdvNo &" and CreatedTransNo = "& iCrTransNo
				objrs.Open sQuery,con,3
				if not objRs.EOF then
				    sQuery =" Update ACC_T_CREATEDADVANCEADJ set AMOUNTADJUSTED = "& dAdvAmount &" where CreatedAdvanceNo = "& iCrAdvNo &" and CreatedTransNo = "& iCrTransNo
				    Response.Write "<p>"& sQuery
				    con.execute sQuery
				else
				    sQuery = "INSERT INTO ACC_T_CREATEDADVANCEADJ (CREATEDADVANCENO, CREATEDTRANSNO, ADJUSTEDON,  "&_
					 "AMOUNTADJUSTED, NARRATION) VALUES "&_
					 "("&iCrAdvNo&", "&iCrTransNo&", Convert(Datetime,'"&sVoucDate&"',103), "&dAdvAmount&", '"&sAdjNarr&"') "

			    Response.Write sQuery &"<br><br>"
			    con.execute(sQuery)
				end if
				objRs.Close 
				

			Elseif CStr(sAdjType) = "D" and CDbl(dAdvAmount) <> 0 Then

				sQuery = "insert into Acc_T_RcvblAdjustmentDetails(ReceivableNumber,RecdByTransactionNo,"&_
						 "ReceivedOn,AmountReceived) Values ("&iAdjRecNo&","&iTransNo&","&_
						 "getdate(),"&dAdvAmount&")"

				Response.Write sQuery& "<BR>"

				con.execute(sQuery)
			End IF
		Next
    end if 'if oNodAdvRoot.hasChildNodes() then
		sQuery="select isnull(ReceivableNumber,0) from Acc_T_CreatedReceivables where CreatedTransNo = "&iCrTransNo

		Response.Write sQuery &"<br><br>"
		objRs.open sQuery,con

		If Not objRs.EOF Then
			iCrPayableNo = objRs(0)
		Else
			iCrPayableNo = 0
		End If
		objRs.Close

		sPayNarration="Sal INV No:"&sSalInvNo &" Dt:"&sSalInvDt
		If cdbl(dAdjAmtTotal) > 0 then
			sQuery="insert into Acc_T_CreatedRcvbleAdjDet(ReceivableNumber,CreatedTransNo,"&_
				"ReceivedOn,AmountReceived) Values ("&iCrPayableNo&","&iCrTransNo&","&_
				"convert(datetime,'"&sVoucDate&"',103),"&dAdjAmtTotal&")"
	        Response.Write "<p>"&sQuery
	        con.execute(sQuery)
		end if







	sQuery="update Acc_T_CreatedVoucherHeader set BookNumber="&sBookNo&","&_
			" ApprovedBy = '"& getUserid() &"',ApprovedOn = Convert(Datetime,getDate(),103), CreatedVouchStatus='"&sVouStatus&"' where CreatedTransNo="&iCrTransNo
			Response.Write sQuery &"<br>"
			con.execute(sQuery)
			'Response.End
End IF

	if con.Errors.count <>0 then
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) &"<br>"
		next
		'Redirect to Error Handling System
	else
	'	Con.RollbackTrans
	'	Response.end

		Response.Clear
		con.CommitTrans

		'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCrTransNo&".xml")

		sMessage="Sales Voucher Accounted Sucessfuly "

	end if

%>
<HTML>
<head>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
</head>
<SCRIPT LANGUAGE=vbscript>
function Message(strr,strr1,sUnitID,sPara)
	IF trim(sPara) = "Acc" then
		MsgBox strr& vbCrLf & strr1
		' window.location.href = "AppOtherVoucherList.asp?selUnitId="&sUnitID&"&selApplication=3&selVoucher=05"
		window.location.href = "SALESVOUCHERS.ASP"
	ElseIF trim(sPara) = "App" or trim(sPara) = "Edt" then
		window.location.href = "SALESVOUCHERS.ASP"
	ElseIF  trim(sPara) = "" then
		MsgBox strr& vbCrLf & strr1
		window.location.href = "AppOtherVoucherList.asp?selUnitId="&sUnitID&"&selApplication=3&selVoucher=05"
	End IF

end function
</SCRIPT>
<BODY onLoad = "Message '<%=sMessage%>','<%=sMessage1%>', '<%=sOrgId%>','<%=sPara%>'"/>
</html>
