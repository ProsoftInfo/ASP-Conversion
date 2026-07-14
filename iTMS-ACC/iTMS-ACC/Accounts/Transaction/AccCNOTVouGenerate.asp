<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AccCNOTVouGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	April 16, 2003
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

Dim oDOM,nodHeader,Root,objRs,sQuery
DIM EntryNode,HeaderNode,nodANL,newElem
DIM sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
DIM sOrgId,sBookNo,sVouType,iVouNo,sVouName
DIM sVouCode,sApprove,sVoucDate,sAccUnit
DIM dTotal,sTransType,dCRAmt,dDRAmt
DIM sAccType,sAccCode,sEntryType,sEntryno,iTransNo
DIM sDocType,sVouStatus
DIM iSeriesNo,iSeriesCode,iCreatedVouNo,iCreatedTransNo,sAccHeadCode
Dim dTdsAmt,dTdsPer,sExp,CheckNode,iCrPayNo,iCrRecNo,sVouEntTy
Dim sSelVouTy,sFormVal,sRetVal

sSelVouTy = Request("voutype")
sFormVal = Request("hFormVal")

sVouName=Request("hVouName")
sVouCode=Request("hVouCode")
iCreatedTransNo=Request("hTransNo")

sVouStatus="010104" 'Voucher Accounted

SET objRs  = server.CreateObject("adodb.recordset")

' Create our DOM Document Objects
SET oDOM = Server.CreateObject("Microsoft.XMLDOM")

'oDOM.Load server.MapPath("../xmldata/Voucher/"&iCreatedTransNo&".xml")
sRetVal = GetVouchXML(iCreatedTransNo)
oDOM.Load server.MapPath(sRetVal)


'Response.Write iCreatedTransNo

sQuery = "Select CreatedTransNo From Acc_T_VoucherHeader Where CreatedTransNo = "&iCreatedTransNo&" "
objRs.Open sQuery,Con
IF Not objRs.EOF Then
	sVouEntTy = "Y" 'Already Accounted
Else
	sVouEntTy = "N" 'Can Be Accounted
End IF
objRs.Close



SET Root=oDOM.documentElement

sOrgId=Root.Attributes.Item(0).nodeValue
sBookNo =Root.Attributes.Item(2).nodeValue
sVouType=Root.Attributes.Item(4).nodeValue
sVoucDate=Root.Attributes.Item(5).nodeValue
sTemp= Split(Root.Attributes.Item(6).nodeValue,"?")
iCreatedVouNo=Root.Attributes.Item(9).nodeValue

IF CStr(sVouEntTy) = "N" Then

	FOR EACH EntryNode IN Root.childNodes
		IF EntryNode.nodeName="Entry" THEN
				dTotal=dTotal+CDbl(EntryNode.Attributes.Item(3).nodeValue)
		END IF
	NEXT

	sTransType=sVouName&"R"

	sQuery="select DrSeriesNo,DrSeriesCode from Acc_M_BookNumberSeries where "&_
			"OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo


	objRs.open sQuery,con
	if not objRs.EOF then
		iSeriesNo=objRs(0)
		iSeriesCode=objRs(1)
	end if	
	objRs.close()

	con.BeginTrans


	iVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)

	sQuery="select isnull(max(TransactionNumber),0)+1 from Acc_T_VoucherHeader"
	objRs.open sQuery,con
		iTransNo=objRs(0)
	objRs.Close


	sQuery=" insert into Acc_T_VoucherHeader (TransactionNumber,OUDefinitionID,BookCode,BookNumber,TransactionType,"&_
			"PartyType,PartySubType,PartyCode,AccountHead,VoucherNumber,CreatedVoucherNo,CreatedTransNo,VoucherDate,VoucherAmount,"&_
			"PayToRecdFrom,CrDrIndication,BankInstrumentType,CreatedBy,CreatedOn,ApprovedBy,AuditedBy,AccountedBy,BRSTransactionNo,VoucherStatus,AccountedOn) values"&_
			"("&iTransNo&",'"&sOrgId&"','"&sVouCode&"',"&sBookNo&",'"&sTransType&"',"&_
			"'"&trim(sTemp(0))&"',"&trim(sTemp(1))&","&trim(sTemp(3))&",NULL,'"&iVouNo&"','"&iCreatedVouNo&"',"&iCreatedTransNo&",convert(datetime,'"&sVoucDate&"',103),"&dTotal&_
			",'Others ','"&sVouType&"','OT',"&getUserid&",getdate(),NULL,NULL,"&getUserid&",NULL,'"&sVouStatus&"',getdate())"
	Response.Write sQuery& "<BR><BR><BR>"
	con.execute(sQuery)

	'-----------------------PROCESS ENTRY NODES-------------------------------


	FOR EACH EntryNode IN Root.childNodes
		IF EntryNode.nodeName="Entry" THEN
			sEntryno=EntryNode.Attributes.Item(0).nodeValue
			sAmount=EntryNode.Attributes.Item(3).nodeValue

			sExp = "//Entry[@No="&sEntryno&" and @TdsAmount]"
			Set CheckNode = Root.selectNodes(sExp)

			IF CheckNode.length <> 0 Then
				dTdsAmt = EntryNode.Attributes.Item(6).nodeValue
				dTdsPer = EntryNode.Attributes.Item(8).nodeValue
			Else
				dTdsAmt = 0
				dTdsPer = 0
			End IF

			sEntryType=EntryNode.Attributes.Item(1).nodeValue
			sAccUnit=EntryNode.Attributes.Item(4).nodeValue

		'---------PROCESS THE CHILD NODES OF ENTRIES FOR DETAIL TABLE UPDATION----
			FOR EACH HeaderNode IN EntryNode.childNodes
				IF HeaderNode.nodeName="AccHead" THEN
						sAccCode=HeaderNode.Attributes.Item(0).nodeValue
						sAccType=HeaderNode.Attributes.Item(4).nodeValue
				END IF 'End of Check for Account head Node
				IF 	HeaderNode.nodeName="Narration" THEN
						sNarration=HeaderNode.text
				END IF 'End of Check for Narration Node
			NEXT
		'-------------END OF PROCESSING CHILD NODES OF ENTRIES---------------------
		'----------------------------DETAIL TABLE UPDATION-------------------------
			IF StrComp(sAccType,"G")=0 THEN
				IF CStr(sAccUnit) = "" Then
					sAccUnit = sOrgId
				End IF
				sQuery="insert into Acc_T_VoucherDetails (TransactionNumber,AccountingUnit,"
				sQuery=sQuery&" VoucherEntryNumber,AccUnitAccountHead,AccUnitPartyType,AccUnitPartyCode,AccUnitPartySubType,"
				sQuery=sQuery&" VoucherNarration, Amount,TransCrDrIndication,TDSonAmount,TDSPercentage) values ("
				sQuery=sQuery& iTransNo&",'"&sAccUnit&"'"
				sQuery=sQuery& ","&sEntryno&","&sAccCode&",NULL,NULL,NULL,"
				sQuery=sQuery&" '"&sNarration&"',"&sAmount&",'"&sEntryType&"', "&dTdsAmt&", "&dTdsPer&")"
			END IF
		Response.Write sQuery& "<BR><BR><BR>"
		con.execute(sQuery)
		'-----------------------END OF DETAIL TABLE UPDATION----------------------
		DIM sCCGroup,sAddCode,sAddRatio,sAddAmount,dAddTotal
		dAddTotal=0
		'--------PROCESS CHILD NODES OF ENTRIES FOR ADDTIONAL DETAILS UPDATION----
			FOR EACH HeaderNode IN EntryNode.childNodes
		'----------------------PROCESS COST CENTER NODES -------------------------
				IF 	HeaderNode.nodeName="CostCenter" THEN
					FOR EACH  nodANL IN HeaderNode.childNodes
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
					NEXT
				END IF
		'-------------END OF PROCESSING COST CENTER NODES ------------------------
		'----------------------PROCESS ANALYTICAL NODES --------------------------
				Dim sAnalAccGp
				IF 	HeaderNode.nodeName="Analytical" THEN
					FOR EACH  nodANL IN HeaderNode.childNodes
						sAddCode=nodANL.Attributes.Item(0).nodeValue
						sAddRatio=nodANL.Attributes.Item(3).nodeValue
						sAddAmount=nodANL.Attributes.Item(4).nodeValue
						sAnalAccGp=nodANL.Attributes.Item(5).nodeValue
						sQuery="INSERT INTO Acc_T_AnalyticalVoucherDet(TransactionNumber, VoucherEntryNumber, AccountingUnit, "&_
							"AccUnitAccountHead, AccUnitAnalyticalCode,"&_
							"RatioPercentage, RatioAmount,AccAHGroupCode)"&_
							" VALUES("& iTransNo& ","& sEntryno& ",'"&sOrgId&"',"&sAccCode&","&_
							""&sAddCode&","&sAddRatio&","&sAddAmount&", '"&sAnalAccGp&"')"
							'Response.Write sQuery& "<BR>"
							con.execute(sQuery)
					NEXT
				END IF
		'-------------END OF PROCESSING ANALYTICAL NODES -------------------------
			NEXT
		'-END OF PROCESSING CHILD NODES OF ENTRIES FOR ADDTIONAL DETAILS UPDATION-
		END IF
	NEXT
	dim iPayableNo
	'------------------END OF PROCESSING ENTRY NODES--------------------------
	if sVouName="CN" then

		sQuery = "Select isNull(Max(PayablesNumber),0)+1 From Acc_T_Payables "
		objRs.open sQuery,con
			iPayableNo=objRs(0)
		objRs.Close

		sQuery = "Select PayablesNumber From Acc_T_CreatedPayables Where CreatedTransNo = "&iCreatedTransNo&" "
		objRs.open sQuery,con
			IF Not objRs.EOF Then
				iCrPayNo = objRs(0)
			End IF
		objRs.Close



		sQuery="INSERT INTO Acc_T_Payables(PayablesNumber,TransactionNumber, OUDefinitionID, "&_
				"VoucherDate, PartyType, PartySubType, PartyCode, "&_
				" AmountPayable, AmountPaid, CrCreatedPayable) values("&iPayableNo&","&iTransNo&",'"&sOrgId&"',"&_
				"convert(datetime,'"&sVoucDate&"',103),'"&trim(sTemp(0))&"',"&trim(sTemp(1))&","&trim(sTemp(3))&","&_
				""&dTotal&",0, "&iCrPayNo&")"

		Response.Write sQuery& "<BR>"
		con.execute(sQuery)
	else
		sQuery="Select isNull(Max(ReceivableNumber),0)+1 From Acc_T_Receivables  "
		objRs.open sQuery,con
			iPayableNo=objRs(0)
		objRs.Close

		sQuery = "Select ReceivableNumber From Acc_T_CreatedReceivables Where CreatedTransNo = "&iCreatedTransNo&" "
		objRs.open sQuery,con
			IF Not objRs.EOF Then
				iCrRecNo = objRs(0)
			End IF
		objRs.Close

		sQuery="INSERT INTO Acc_T_Receivables(ReceivableNumber,TransactionNumber, OUDefinitionID, "&_
				"VoucherDate, PartyType, PartySubType, PartyCode,"&_
				"AmountReceivable, AmountReceived, DrCreatedReceivable) values("&iPayableNo&","&iTransNo&",'"&sOrgId&"',"&_
				"convert(datetime,'"&sVoucDate&"',103),'"&trim(sTemp(0))&"',"&trim(sTemp(1))&","&trim(sTemp(3))&","&_
				""&dTotal&",0, "&iCrRecNo&")"

		'Response.Write sQuery& "<BR>"
		con.execute(sQuery)

	end if
	sQuery="update Acc_T_CreatedVoucherHeader set  CreatedVouchStatus='"&sVouStatus&"' where CreatedTransNo="&iCreatedTransNo
	con.execute(sQuery)

	if con.Errors.count <>0 THEN
		con.RollbackTrans
		FOR iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) &"<br>"
		NEXT
		'Redirect to Error Handling System
	ELSE
	   ' con.RollbackTrans
	    'Response.End 
		con.CommitTrans
		

		Root.Attributes.Item(9).nodeValue=iTransNo
		Root.Attributes.Item(10).nodeValue=iVouNo

		'Root.Attributes.Item(2).nodeValue=iTransNo
		'Root.Attributes.Item(3).nodeValue=iVouNo

		'Set newElem  = oDOM.createAttribute("CreatedTransNo")
		'newElem.value = iCreatedTransNo
		'Root.setAttributeNode(newElem)

		'Set newElem  = oDOM.createAttribute("CreatedVoucherNo")
		'newElem.value = iCreatedVouNo
		'Root.setAttributeNode(newElem)

		'oDOM.Save server.MapPath("../xmldata/Voucher/"&iCreatedTransNo&".xml")

		Response.Redirect ("CREDITVOUCHERS.ASP?hFormVal="&sSelVouTy&"&voutype="&sSelVouTy)

	END IF
Else
	Response.Redirect ("CREDITVOUCHERS.ASP?hFormVal="&sSelVouTy&"&voutype="&sSelVouTy)
End IF

%>

