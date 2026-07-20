<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"

	'Program Name				:	GetCDNVoucherxml.asp
	'Module Name				:	ACCOUNTS (Reports)
	'Author Name				:	Maheshwari S.
	'Created On					:	Oct 25, 2006
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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

Public Function GetXmlForCN(iVNo,sDispTy)
	Dim sBkCode
	Dim Root,objhttp,rs1,rs2,Advrs,rs3,rs4,rs5,Hnode,sName,sQry,OrgNode,oDOM,NewElem,fsobj,iCount,i
	Dim sOrgId,sBkNo,sBkName,sCrDr,sVDate,sBkAccNo,sTransNo,sVouNo,sCTNo,sCVNo,eCrDr
	Dim iENo,iEtCrDr,iPayToRecdFrom,iEAmt,iEAccUnit,sAccName,iETdsAmt,iETdsPer,SalTypNode
	Dim sAccHeadNo,sAccUnit,sAccPartyType,sAccPartySubType,sAccPartyCode,sPartyName,sSubTypeName
	Dim sTempNo,BookNode,PurTypNode,PurInvNode,PartyNode,AccHeadNode,TaxNode,sAccHeadName,sPassCode,sPType,sPName
	Dim sTempInvNo,sArr,sPurInvNo,sPurInvDate,sInvNo,sInvDate,sApprover,sFrmApp,iBkCode,iRoundOff
	Dim sParType,sParSubType,sParCode,sCostCnt,sAnlCnt,sAppTrNo,sBankType,sBKInsNo,sBKInsType
	Dim DetailNode,sBasVal,sVouDate,EntryNode,TaxDetNode,tVouAmt,NewTaxAmt,sSalName,SalInvNode
	Dim count,eNo,ePayTo,eAmt,eQty,eUOM,eRate,eActVal,eDisPer,eDisAmt,eItemCode,eClassCode
	Dim sCatCode,sTaxCode,sTaxMode,sTaxFormula,sTaxValue,sTaxAmount,sAccHead,sTaxShtName,sRoundOff
	Dim AdvNode,aAdvNo,aAmtRec,aAntToAdj,dCTransNo,aCTrNo,aTrNo,aVouNo,aVouDate,eTdsAmt,eTdsPer,sNarr 
	Dim sTdsElgi,sRetVal,sVouNarr,sCrInvType
	set oDOM=Server.CreateObject("Microsoft.XMLDOM")
	Set rs1 = server.CreateObject("ADODB.RecordSet")
	Set rs2 = server.CreateObject("ADODB.RecordSet")
	Set rs3 = server.CreateObject("ADODB.RecordSet")
	Set rs4 = server.CreateObject("ADODB.RecordSet")
	Set rs5 = server.CreateObject("ADODB.RecordSet")
	
	sPassCode = "07"
	
	

	sQry = "select BankInstrumentType from  Acc_T_CreatedVoucherHeader where CreatedTransNo = "&iVNo&" "
	with rs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQry
		.ActiveConnection = con 
		.Open
	End With
		
	IF not rs1.Eof then
		sBankType = rs1(0)
	End If
	rs1.Close
	
	If trim(sBankType) = "OT" then
		sQry = " Select A.CreatedTransNo,A.CreatedVoucherNo,isNull(V.TransactionNumber,0),isNull(V.VoucherNumber,'') from Acc_T_CreatedVoucherHeader A LEFT JOIN Acc_T_VoucherHeader V ON A.CreatedTransno = V.CreatedTransno where A.CreatedTransno = "& iVNo &" "
		
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
	
		IF not rs1.Eof then
			sCTNo = rs1(0)
			sCVNo = rs1(1)
			sTransNo = rs1(2)
			sVouNo = rs1(3)
		End IF
		rs1.Close
			
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = " select A.OUDefinitionID,D.OrgUnitShortDescription,A.BookNumber,A.CrDrIndication,Convert(VarChar,A.VoucherDate,103) from Acc_T_CreatedVoucherHeader A,"&_
					  " DCS_OrganizationUnitDefinitions D where createdTransNo = "&iVNo&" "&_
					  " and D.OUDefinitionID = A.OUDefinitionID "
			.ActiveConnection = con 
			.Open
		End With
	
		If not rs1.Eof then
			sOrgId = rs1(0)
			sName = rs1(1)
			sBkNo = rs1(2)
			sCrDr = rs1(3)	
			sVouDate = rs1(4)
		End If			
		rs1.close
									
		with rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source =" Select BookName from VwOrgBookNames V,Acc_T_CreatedVoucherHeader A "&_
					 " Where V. OUDefinitionID =  "&sOrgId&" and V.BookCode = "&sPassCode&" "&_
					 " and A.CreatedTransNo = "& iVNo &"  and V.BookNumber =  "&sBkNo&" "
			.ActiveConnection = con 
			.Open
		End With
		If not rs1.EOF then
			sBkName = rs1(0)
		End If 
		rs1.Close 
				
		sQry = " select A.PartyType,A.PartySubType,A.PartyCode,V.SubTypeName from Acc_T_CreatedVoucherHeader A, "&_
			   " VwOrgPartyType V where A.CreatedTransNo = "&iVNo&" and  V.OUDefinitionID = A.OUDefinitionID and "&_
			   " V.PartySubType = A.PartySubType and V.PartyType = A.PartyType "
			
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con
			.Open
		End with
		IF not rs1.EOF Then
			sParType  = rs1(0)
			sParSubType = rs1(1)
			sParCode = rs1(2)
			sSubTypeName = rs1(3)
		End IF
		rs1.Close
			
		with rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = " Select PartyName from App_M_PartyMaster where PartyCode = "&sParCode&" "
			.ActiveConnection = con
			.Open
		End with
		If not rs1.EOF then
			sPartyName	= rs1(0)
		End If	
		rs1.Close 
				
				
		IF CStr(sVouNo) = "" Then
			sVouNo = sCVNo
			sTransNo = sCTNo
		End IF
							
		set Root=oDOM.CreateElement("voucher")
		Root.setAttribute "UnitNo",sOrgId
		Root.setAttribute "UnitName",sName
		Root.setAttribute "BookNo",sBkNo
		Root.setAttribute "BookName",sBkName
		Root.setAttribute "CRDR",sCrDr
		Root.setAttribute "VouDate",sVouDate
		Root.setAttribute "PartyCode",sParType&"?"&sParSubType&"?"&sSubTypeName&"?"&sParCode
		Root.setAttribute "Approver","S"
		Root.setAttribute "PartyName",sPartyName
		Root.setAttribute "TransNo",sTransNo
		Root.setAttribute "VoucherNo",sVouNo
		Root.setAttribute "CreatedTransNo",sCTNo
		Root.setAttribute "CreatedVoucherNo",sCVNo
		oDOM.appendchild Root
'===================================creation of Entry Node ============================
		sQry = " Select VoucherEntryNumber,TransCrDrIndication,isNull(ItemDescription,0), "&_
			   " Amount,TdsOnAmount,TdsPercentage,AccUnitAccountHead,ItemCode,ClassificationCode from Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iVNo &" "
				
		with rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con
			.Open
		End With
				
		Do While not rs1.EOF
			eNo = rs1(0)
			eCrDr = rs1(1)
			ePayTo = rs1(2)
			eAmt = rs1(3)
			eTdsAmt = rs1(4)
			eTdsPer = rs1(5)
			sAccHeadNo = rs1(6)
			eItemCode = rs1(7)
			eClassCode = rs1(8)
							
			sQry = "Select Distinct EligibleForTds From VwOrgGLHeads Where  "&_
					 "AccountHead = "&sAccHeadNo&" and OUDefinitionID = '"&sOrgId&"' "
			rs2.Open sQry,Con
			IF Not rs2.EOF Then
				sTdsElgi = rs2(0)
			Else
				sTdsElgi = 0
			End IF
			rs2.Close
				
			Set EntryNode = oDOM.createElement("Entry")
			EntryNode.setAttribute "No",eNo
			EntryNode.setAttribute "CRDR",eCrDr
			EntryNode.setAttribute "Payto",ePayTo
			EntryNode.setAttribute "Amount",eAmt
			EntryNode.setAttribute "AccUnit",sOrgId
			EntryNode.setAttribute "AccName",sName
			EntryNode.setAttribute "TdsAmount",eTdsAmt
			EntryNode.setAttribute "TDSElgi",sTdsElgi
			EntryNode.setAttribute "TdsPercentage",eTdsPer
			EntryNode.setAttribute "PayRecAmount","0"
			EntryNode.setAttribute "ItemCode",eItemCode 
			EntryNode.setAttribute "ClassCode",eClassCode  
			Root.appendchild EntryNode
				
			With rs2
				.CursorLocation = 3
				.CursorType = 3
				.Source = "select Count(1) from Acc_T_CreatedVoucherCCDet where CreatedTransNo = "&iVNo&" "&_
						  " and AccUnitAccountHead = "&sAccHeadNo&" "
				.ActiveConnection = con
				.Open
			End With
					
			If not rs2.EOF	then
				sCostCnt = rs2(0)
			End If		
			rs2.Close									
																							
			With rs2
				.CursorLocation = 3
				.CursorType = 3
				.Source = "Select Count(1) from Acc_T_CretedVoucherAHDet where CreatedTransNo =	 "&iVNo&" "&_
						   " and AccUnitAccountHead = "&sAccHeadNo&" "
				.ActiveConnection = con
				.Open
			End With
															
			If not rs2.EOF	then
				sAnlCnt = rs2(0)
			End If		
			rs2.Close
																
			With rs2
				.CursorLocation = 3
				.CursorType = 3
				.Source = "select AccountDescription from Acc_M_GLAccountHead where AccountHead = "& sAccHeadNo &" "
				.ActiveConnection = con
				.Open
			End With
			If not rs2.EOF	then
				sAccHeadName = rs2(0)
			End If		
			rs2.Close
										
			Set AccHeadNode = oDOM.createElement("AccHead")
			AccHeadNode.setAttribute "No",sAccHeadNo
			AccHeadNode.setAttribute "CostCenter",sCostCnt
			AccHeadNode.setAttribute "Analytical",sAnlCnt
			AccHeadNode.setAttribute "Name",sAccHeadName
			AccHeadNode.setAttribute "Type","G"
			AccHeadNode.setAttribute "TransFlag","A"
			EntryNode.Appendchild AccHeadNode
									
			with rs2
				.CursorLocation = 3
				.CursorType = 3
				.Source = "select VoucherNarration from Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iVNo &" "&_
						  " and VoucherEntryNumber = "&eNo&" "
				.ActiveConnection = con
				.Open
			End with
					
			If not rs2.EOF then
				sNarr = rs2(0)
			End If 
			rs2.close
																			
			Set PartyNode = oDom.CreateElement("Narration")
			PartyNode.text = sNarr
			EntryNode.appendchild PartyNode
			rs1.Movenext
		loop
		rs1.close
	End IF 
	
'**********************************************************************************************
'********************************** For the Sales Commission Function Call********************* 
'**********************************************************************************************
	Dim iSalTrNo
	
	
	IF CStr(sBankType) = "SC" Then
		sQry =	"Select A.CreatedTransNo,A.CreatedVoucherNo,isNull(V.TransactionNumber,0),isNull(V.VoucherNumber,'') from Acc_T_CreatedVoucherHeader A LEFT JOIN Acc_T_VoucherHeader V ON A.CreatedTransno = V.CreatedTransno where A.CreatedTransno = "& iVNo &" "
			
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
		
		IF not rs1.Eof then
			sCTNo = rs1(0)
			sCVNo = rs1(1)
			sTransNo = rs1(2)
			sVouNo = rs1(3)
		End IF
		rs1.Close
				
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = " select A.OUDefinitionID,D.OrgUnitShortDescription,A.BookNumber,A.CrDrIndication, "&_
					  " Convert(Varchar,A.VoucherDate,103),isNull(BankInstrumentNo,'0') from Acc_T_CreatedVoucherHeader A, "&_
					  " DCS_OrganizationUnitDefinitions D where createdTransNo = "&iVNo&"  "&_
					  " and D.OUDefinitionID = A.OUDefinitionID  "
			.ActiveConnection = con 
			.Open
		End With
		
		If not rs1.Eof then
			sOrgId = rs1(0)
			sName = rs1(1)
			sBkNo = rs1(2)
			sCrDr = rs1(3)	
			sVouDate = rs1(4)
			iSalTrNo = rs1(5)
		End If			
			rs1.close
										
		with rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source =" Select BookName from VwOrgBookNames V,Acc_T_CreatedVoucherHeader A "&_
					 " Where V. OUDefinitionID =  "&sOrgId&" and V.BookCode = "&sPassCode&" "&_
					 " and A.CreatedTransNo = "& iVNo &"  and V.BookNumber =  "&sBkNo&" "
			.ActiveConnection = con 
			.Open
		End With
		If not rs1.EOF then
			sBkName = rs1(0)
		End If 
		rs1.Close 
					
		sQry = " select A.PartyType,A.PartySubType,A.PartyCode,V.SubTypeName from Acc_T_CreatedVoucherHeader A, "&_
			   " VwOrgPartyType V where A.CreatedTransNo = "&iVNo&" and  V.OUDefinitionID = A.OUDefinitionID and "&_
			   " V.PartySubType = A.PartySubType and V.PartyType = A.PartyType "
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con
			.Open
		End with
			
		IF Not rs1.EOF Then
			sParType  = rs1(0)
			sParSubType = rs1(1)
			sParCode = rs1(2)
			sSubTypeName = rs1(3)
		End IF
		rs1.Close
			
		with rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = " Select PartyName from App_M_PartyMaster where PartyCode = "&sParCode&" "
			.ActiveConnection = con
			.Open
		End with
		If not rs1.EOF then
			sPartyName	= rs1(0)
		End If	
		rs1.Close 
					
		IF CStr(sVouNo) = "" Then
			sVouNo = sCVNo
			sTransNo = sCTNo
		End IF
		
		Dim sSalInvNo,sSalInvDate,iSalCommVal
		
		with rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "Select VoucherNumber,Convert(Varchar,VoucherDate,103) From Acc_T_VoucherHeader Where TransactionNumber = "&iSalTrNo
			.ActiveConnection = con
			.Open
		End with
		If not rs1.EOF then
			sSalInvNo = rs1(0)
			sSalInvDate = rs1(1)
		End If	
		rs1.Close 
		
		with rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "Select CommissionValue From Sal_T_AdditionalAgents Where AccTransactionNo = "&iSalTrNo
			.ActiveConnection = con
			.Open
		End with
		If not rs1.EOF then
			iSalCommVal = rs1(0)
		End If	
		rs1.Close 
				
		Dim MainRoot
		Set MainRoot = oDOM.CreateElement("Root")
		oDOM.appendchild MainRoot
			
		
		set Root=oDOM.CreateElement("voucher")
		Root.setAttribute "UnitNo",sOrgId
		Root.setAttribute "UnitName",sName
		Root.setAttribute "BookNo",sBkNo
		Root.setAttribute "BookName",sBkName
		Root.setAttribute "VouDate",sVouDate
		Root.setAttribute "Approver","S"
		Root.setAttribute "SalTransNo",iSalTrNo
		Root.setAttribute "SalVouNo",sSalInvNo
		Root.setAttribute "SalVouDate",sSalInvDate
		Root.setAttribute "CrTransNo",sCTNo
		Root.setAttribute "CrVoucherNo",sCVNo
		Root.setAttribute "TransNo",sTransNo
		Root.setAttribute "VoucherNo",sVouNo
		Root.setAttribute "CommisionValue",iSalCommVal
					
		Set PartyNode = oDOM.CreateElement("Party")
		PartyNode.setAttribute "ParType",sParType
		PartyNode.setAttribute "ParSubType",sParSubType
		PartyNode.setAttribute "ParCode",sParCode
		PartyNode.text = sPartyName
		Root.appendChild PartyNode
		MainRoot.appendchild Root
			
		
'=======================================creation of Entry Node ============================
		sQry = "SELECT T.VoucherEntryNumber, T.AccUnitAccountHead, M.AccountDescription, "&_
			   "T.Amount, T.TransCrDrIndication, T.TDSonAmount, T.TDSPercentage,T.ItemCode,T.ClassificationCode "&_
			   "FROM Acc_M_GLAccountHead M INNER JOIN Acc_T_CreatedVoucherDetails T ON "&_
			   "M.AccountHead = T.AccUnitAccountHead Where T.CreatedTransNo = "&iVNo&" "
							
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con
			.Open
		End With
		Set rs1.ActiveConnection = Nothing
		Do While Not rs1.EOF 
			set EntryNode = oDOM.createElement("Entry")
			EntryNode.setAttribute "No",rs1("VoucherEntryNumber")
			EntryNode.setAttribute "PayTo",rs1("AccountDescription")
			EntryNode.setAttribute "Amount",rs1("Amount")
			EntryNode.setAttribute "CRDR",rs1("TransCrDrIndication")
			EntryNode.setAttribute "TdsAmount",rs1("TDSonAmount")
			EntryNode.setAttribute "TdsElgi","0"
			EntryNode.setAttribute "TdsPercentage",rs1("TDSPercentage")
			EntryNode.setAttribute "ItemCode",rs1("ItemCode")
			EntryNode.setAttribute "ClassCode",rs1("ClassificationCode")
			MainRoot.appendchild EntryNode
						
			set AccHeadNode = oDOM.createElement("AccHead")
			AccHeadNode.setAttribute "No",rs1("AccUnitAccountHead")
			AccHeadNode.setAttribute "CostCenter","0"
			AccHeadNode.setAttribute "Analytical","0"
			AccHeadNode.setAttribute "Name",rs1("AccountDescription")
			AccHeadNode.setAttribute "Type","G"
			AccHeadNode.setAttribute "TransFlag","A"
			EntryNode.Appendchild AccHeadNode
						
			with rs2
				.CursorLocation = 3
				.CursorType = 3
				.Source = "select VoucherNarration from Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iVNo &" "&_
						  " and VoucherEntryNumber = "&rs1(0)&" "
				.ActiveConnection = con
				.Open
			End with
			If not rs2.EOF then
				sNarr = rs2(0)
			End If 
			rs2.close
																				
			Set PartyNode = oDom.CreateElement("Narration")
			PartyNode.text = sNarr
			EntryNode.appendchild PartyNode
			rs1.Movenext
		loop
		rs1.close
	End IF
'**********************************************************************************************
'********************************** For the Sales Returns Function Call********************* 
'**********************************************************************************************
	If trim(sBankType) = "SR" then
		sQry =  " Select A.CreatedTransNo,A.CreatedVoucherNo,isNull(V.TransactionNumber,0),isNull(V.VoucherNumber,'') from Acc_T_CreatedVoucherHeader A LEFT JOIN Acc_T_VoucherHeader V ON A.CreatedTransno = V.CreatedTransno where A.CreatedTransno = "& iVNo &" "
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
	
		IF not rs1.Eof then
			sCTNo = rs1(0)
			sCVNo = rs1(1)
			sTransNo = rs1(2)
			sVouNo = rs1(3)
		End IF
		rs1.Close
		set Root=oDOM.CreateElement("Voucher")
		Root.setAttribute "CreatedTransNo", sCTNo
		Root.setAttribute "CreatedVouNo",sCVNo 
		Root.setAttribute "TransNo", sTransNo 
		Root.setAttribute "VouNo", sVouNo 		
		oDOM.appendchild Root
'=================================creation of Header node=======================	
		Set Hnode = oDOM.createElement("Header")
		Root.appendchild Hnode
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = " select OUDefinitionID,isNull(BankInstrumentType,0),VoucherAmount from Acc_T_CreatedVoucherHeader where CreatedTransno = "&iVNo&" "
			.ActiveConnection = con 
			.Open
		End With
		
		IF not rs1.Eof Then
			sOrgId = rs1(0)
			sPType = rs1(1)
			tVouAmt = rs1(2)
		End IF
		rs1.Close
					
		with rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source =" Select D.OrgUnitShortDescription,A.BookNumber from DCS_OrganizationUnitDefinitions D,Acc_T_CreatedVoucherHeader A "&_
					 " where A.CreatedTransNo = "& iVNo &" and D.OUDefinitionID =  "&sOrgId&" "
			.ActiveConnection = con 
			.Open
		End With
		
		Response.Write rs1.Source &"<br>"
		If not rs1.EOF then
			sName = rs1(0)
			sBkNo = rs1(1)
		End If 
		rs1.Close 
		
		set OrgNode = oDOM.createElement("Organization")
		OrgNode.setAttribute "OrgId",sOrgId
		OrgNode.setAttribute "AccUnit",""
		OrgNode.text = sName
		Hnode.appendchild OrgNode
			
		with rs1
				.CursorLocation = 3
				.CursorType = 3
				.Source =" Select BookName from VwOrgBookNames V,Acc_T_CreatedVoucherHeader A "&_
						 " Where V. OUDefinitionID =  "&sOrgId&" and V.BookCode = "&sPassCode&" "&_
						 " and A.CreatedTransNo = "& iVNo &"  and V.BookNumber =  "&sBkNo&" "
				.ActiveConnection = con 
				.Open
		End With
		Response.Write rs1.Source &"<br>"
		If not rs1.EOF then
			sBkName = rs1(0)
		End If 
		rs1.Close 
			
		Set BookNode = oDOM.createElement("Book")
		BookNode.setAttribute "BookId",sBkNo
		BookNode.setAttribute "BKAccHead","0"
		BookNode.setAttribute "BKOtherUnits","1"
		BookNode.text = sBkName
		Hnode.appendchild BookNode
			
		sQry = "select isNull(BankInstrumentNo,0) from Acc_T_CreatedVoucherHeader where createdTransNo = "&iVNo&" "
		'Response.Write sQry &"<br>"
		with rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
		If not rs1.EOF then
			sBKInsNo = rs1(0)
		End If 
		rs1.Close 
			
		'sQry = "Select isNull(BankInstrumentType,0) From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&sBKInsNo&" "
		sQry = " Select TypeOfSale from Sal_T_SalesReturnHeader S, ACC_T_CreatedVoucherCRDRNotes A where A.RefNumber = s.SalesReturnNo"&_
		"  and A.CreatedTransNo = "& iVNo
		
		with rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
			
		If not rs1.EOF then
			sBKInsType = rs1(0)
		Else
			sBKInsType = 0
		End If 
		rs1.Close 
			
		sQry = " Select S.InvoiceTypeName from Sal_M_InvoiceTypes S "&_
			   " where  S.InvoiceType = "&sBKInsType&"  "

			
		with rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
		If not rs1.EOF then
			sSalName = rs1(0)
		End If 
		rs1.Close 
							
		Set SalTypNode = oDOM.createElement("SalesType")
		SalTypNode.setAttribute "SalType",sBKInsType
		SalTypNode.text = sSalName
		Hnode.appendchild SalTypNode
				
		'sQry = " select PayToRecdFrom,Convert(Varchar,VoucherDate,103),CreatedBy from Acc_T_CreatedVoucherHeader  where CreatedTransNo =  "& iVNo &" "
		sQry = "Select InvoiceNumber,Convert(varchar,InvoiceDate,103),CreatedBy from Sal_T_SalesReturnHeader S, ACC_T_CreatedVoucherCRDRNotes A where A.RefNumber = s.SalesReturnNo "&_
		" and A.CreatedTransNo = "& iVNo
		with rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
		
		If not rs2.EOF then
			sInvNo = rs2(0)						
			sInvDate = rs2(1)
			sApprover = rs2(2)
		End If 'If not rs2.EOF then
		rs2.Close 
			
		Set SalInvNode = oDOM.createElement("SaleInvoice")
		SalInvNode.setAttribute "InvNo",sInvNo
		SalInvNode.setAttribute "InvDate",sInvDate
		SalInvNode.setAttribute "RefNo",sInvNo
		SalInvNode.setAttribute "Approval","Y"
		SalInvNode.setAttribute "Approver",sApprover
		SalInvNode.setAttribute "SalTrNo",sBKInsNo
		Hnode.appendchild SalInvNode
					
		sQry =  " select A.PartyType,A.PartySubType,A.PartyCode,V.SubTypeName from Acc_T_CreatedVoucherHeader A, "&_
				" VwOrgPartyType V where A.CreatedTransNo = "&iVNo&" and  V.OUDefinitionID = A.OUDefinitionID and "&_
				" V.PartySubType = A.PartySubType and V.PartyType = A.PartyType "
					
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con
			.Open
		End with
		IF not rs1.EOF Then
			sParType  = rs1(0)
			sParSubType = rs1(1)
			sParCode = rs1(2)
			sSubTypeName = rs1(3)
		End IF
		rs1.Close
		
		with rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = " Select PartyName from App_M_PartyMaster where PartyCode = "&sParCode&" "
			.ActiveConnection = con
			.Open
		End with
		If not rs1.EOF then
			sPartyName	= rs1(0)
		End If	
		rs1.Close 
						
		Set PartyNode = oDOM.createElement("Party")
		PartyNode.setAttribute "ParType",sParType
		PartyNode.setAttribute "ParSubType",sParSubType
		PartyNode.setAttribute "ParSubTypeName",sSubTypeName
		PartyNode.setAttribute "ParCode",sParCode
		PartyNode.text = sPartyName
		Hnode.appendchild PartyNode
'====================================creation of Detail node============================	
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = " select Sum(amount) from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iVNo&" "
			.ActiveConnection = con
			.Open
		End with
		If not rs1.EOF then
			sBasVal = rs1(0)
		End If
		rs1.Close
		
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = " select Convert(Varchar,VoucherDate,103) from Acc_T_CreatedVoucherHeader where CreatedTransNo =  "&iVNo&" "
			.ActiveConnection = con
			.Open
		End with
		If not rs1.EOF then
			sVouDate = rs1(0)
		End If
		rs1.Close
					
		Set DetailNode = oDOM.createElement("Details")
		DetailNode.setAttribute "BasicValue",sBasVal
		DetailNode.setAttribute "Discount","0"
		DetailNode.setAttribute "ActualValue",sBasVal
		DetailNode.setAttribute "VouDate",sVouDate
		Root.appendchild DetailNode
		
		sQry = " Select VoucherEntryNumber,ItemDescription,Amount,InvoicedQuantity,InvoicedUOM,InvoicedRate,BasicAmount, "&_
				" DiscountPercent,DiscountAmount,ItemCode,ClassificationCode from Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iVNo &" "
		with rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con
			.Open
		End With
		If not rs1.EOF then
		    do while not rs1.EOF 
			eNo = rs1(0)
			ePayTo = rs1(1)
			eAmt = rs1(2)
			eQty = rs1(3)
			eUOM = rs1(4)
			eRate = rs1(5)
			eActVal = rs1(6)
			eDisPer = rs1(7)
			eDisAmt = rs1(8)
			eItemCode = rs1(9)
			eClassCode = rs1(10)
			
			Set EntryNode = oDOM.createElement("Entry")
		        EntryNode.setAttribute "No",eNo
		        EntryNode.setAttribute "PayTo",ePayTo
		        EntryNode.setAttribute "Amount",eAmt
		        EntryNode.setAttribute "Qty",eQty
		        EntryNode.setAttribute "UOM",eUOM
		        EntryNode.setAttribute "UOMValue",eUOM
		        EntryNode.setAttribute "Rate",eRate
		        EntryNode.setAttribute "ActValue",eActVal
		        EntryNode.setAttribute "DisPer",eDisPer
		        EntryNode.setAttribute "DisAmount",eDisAmt
		        EntryNode.setAttribute "ItemCode",eItemCode
		        EntryNode.setAttribute "ClassCode",eClassCode 
		        DetailNode.appendchild EntryNode
		        
		        
		            With rs2
			            .CursorLocation = 3
			            .CursorType = 3
			            .Source =" Select isNull(AccUnitAccountHead,0)from Acc_T_CreatedVoucherDetails where CreatedtransNo = "&iVNo&" "&_
					             " and VoucherEntryNumber = "&eNo&" "
			            .ActiveConnection = con
			            .Open
		            End with	
		            Do while not rs2.EOF 
			            sAccHeadNo = rs2(0)
			            With rs3
				            .CursorLocation = 3
				            .CursorType = 3
				            .Source = "select Count(1) from Acc_T_CreatedVoucherCCDet where CreatedTransNo = "&iVNo&" "&_
						              " and AccUnitAccountHead = "&sAccHeadNo&" "
				            .ActiveConnection = con
				            .Open
			            End With
			            If not rs3.EOF	then
				            sCostCnt = rs3(0)
			            End If		
			            rs3.Close									
            																						
			            With rs3
				            .CursorLocation = 3
				            .CursorType = 3
				            .Source = "Select Count(1) from Acc_T_CretedVoucherAHDet where CreatedTransNo =	 "&iVNo&" "&_
						               " and AccUnitAccountHead = "&sAccHeadNo&" "
				            .ActiveConnection = con
				            .Open
			            End With
			            If not rs3.EOF	then
				            sAnlCnt = rs3(0)
			            End If		
			            rs3.Close
            															
			            With rs3
				            .CursorLocation = 3
				            .CursorType = 3
				            .Source = "select AccountDescription from Acc_M_GLAccountHead where AccountHead = "& sAccHeadNo &" "
				            .ActiveConnection = con
				            .Open
			            End With
			            If not rs3.EOF	then
				            sAccHeadName = rs3(0)
			            End If		
			            rs3.Close
            									
			            Set AccHeadNode = oDOM.createElement("AccHead")
			            AccHeadNode.setAttribute "No",sAccHeadNo
			            AccHeadNode.setAttribute "CostCenter",sCostCnt
			            AccHeadNode.setAttribute "Analytical",sAnlCnt
			            AccHeadNode.setAttribute "Name",sAccHeadName
			            AccHeadNode.setAttribute "Type","G"
			            AccHeadNode.setAttribute "Group",""
			            EntryNode.Appendchild AccHeadNode
			            rs2.Movenext
		            loop
		            rs2.close
		    rs1.MoveNext 
		    loop
		End If
		rs1.Close 
		
		
		Set TaxDetNode = oDOM.createElement("TaxDetails")
		TaxDetNode.setAttribute "InvoiceVlaue",tVouAmt
		TaxDetNode.setAttribute "BasicValue",sBasVal
		TaxDetNode.setAttribute "NettValue",tVouAmt
		TaxDetNode.setAttribute "RoundOffValue",iRoundOff
		Root.appendchild TaxDetNode
				
		sQry = " Select V.TaxShortName,V.TaxCategoryCode,V.TaxCode,V.ComputationMode,isnull(V.SumOfFields,'') Formula,"&_
				" V.AccountHead,isNull(V.RoundOff,0) RoundOff ,T.TaxAMount,isnull(V.FlatAmount,0) TaxValue from "&_ 
				" VwSalesTaxDetails V,Acc_T_CreatedVoucherTaxDet T where V.ComputationMode is not null and  "&_
				" V.OUDefinitionID= "&sOrgId&" and V.InvoiceType=T.InvoiceType and V.TaxCategoryCode = T.TaxCategoryCode "&_
				" and V.TaxCode = T.TaxCode and T.TaxCode <> 0 and T.TaxCategoryCode <> 0 and T.CreatedTransNo = "&iVNo&" order by V.TaxHierarchy "
	   ' Response.Write "<textarea>"& sQry &"</textarea>"
	   ' Response.End 
		with rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con
			.Open
		End With
		'IF Not rs1.EOF Then
			Do while  not rs1.Eof 
				sTaxShtName = rs1(0)
				sCatCode = rs1(1)
				sTaxCode = rs1(2)
				sTaxMode = rs1(3)
				sTaxFormula = rs1(4)
				sTaxValue = rs1(8)
				sTaxAmount = rs1(7)
				sAccHead = rs1(5)
				sRoundOff = rs1(6)
									
				Set TaxNode = oDOM.createElement("Tax")
				TaxNode.setAttribute "CatCode",sCatCode
				TaxNode.setAttribute "TaxCode",sTaxCode
				TaxNode.setAttribute "TaxMode",sTaxMode
				TaxNode.setAttribute "TaxFormula",sTaxFormula
				TaxNode.setAttribute "TaxValue",sTaxValue
				TaxNode.setAttribute "TaxAmount",sTaxAmount
				TaxNode.setAttribute "AccHead",sAccHead
				TaxNode.setAttribute "ItemValue","0"
				TaxNode.setAttribute "RoundOff",sRoundOff
				TaxNode.text = sTaxShtName					
				TaxDetNode.appendchild TaxNode
				
				rs1.MoveNext
			loop
			rs1.Close
		'Else
		NewTaxAmt = 0
		sQry = "Select TaxAmount,TransCrDrIndication,AccountHead from Acc_T_CreatedVoucherTaxDet where "&_
				" TaxCategoryCode = 0 and TaxCode = 0 and CreatedTransNo = "&iVNo&" "
		with rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con
			.Open
		End With
		If not rs2.EOF then
		    ''blocked and added by ragav on Jan 11,2012 to avoid the balance mistake
		    'IF CStr(rs2(1)) = "D" Then
			IF CStr(rs2(1)) = "C" Then
				NewTaxAmt = rs2(0)
				NewTaxAmt = CDbl(NewTaxAmt) * - 1
				sAccHead = rs2(2)
			Else
				NewTaxAmt = rs2(0)
				sAccHead = rs2(2)
			End IF
		End If		
		rs2.close
		set TaxNode = oDOM.createElement("Tax")
		TaxNode.setAttribute "CatCode","0"
		TaxNode.setAttribute "TaxCode","0"
		TaxNode.setAttribute "TaxMode","0"
		TaxNode.setAttribute "TaxFormula","0"
		TaxNode.setAttribute "TaxValue","0"
		TaxNode.setAttribute "TaxAmount",NewTaxAmt
		TaxNode.setAttribute "AccHead",sAccHead
		TaxNode.text = "ROUND OFF"					
		TaxDetNode.appendchild TaxNode
		'End IF	
		
		
		sQry =  " SELECT A.CreatedAdvanceNo,V.CreatedVoucherNo, V.VoucherDate, V.VoucherAmount, "&_
				" P.AmountReceived, V.CreatedTransNo FROM Acc_T_CreatedVoucherHeader V INNER JOIN "&_
				" Acc_T_CreatedAdvances A ON V.CreatedTransNo = A.CreatedTransNo INNER JOIN "&_
				" Acc_T_CreatedRcvbleAdjDet P ON A.CreatedAdvanceNo = P. ReceivableNumber WHERE "&_
				" P.AdjustType IS NOT NULL AND V.BookCode IN ('01', '02') and "&_
				" P.CreatedTransNo = "&iVNo&" "
		Set AdvNode = oDOM.createElement("AdvanceDetails")
		Root.appendchild AdvNode
		with rs2
			.CursorLocation = 3
			.CursorType = 3 
			.Source = sQry
			.ActiveConnection = con
			.Open
		End with
		
		Do while not rs2.EOF 
			aAdvNo = rs2(0)
			aAmtRec = rs2(3)
			aAntToAdj = rs2(4)
			dCTransNo = rs2(5)
			aCTrNo = rs2(6)
								
			sQry = " Select TransactionNumber,VoucherNumber,Convert(Varchar,VoucherDate,103) from Acc_T_VoucherHeader where createdTransNo = "&dCTransNo&" "
			with rs1
				.CursorLocation = 3
				.CursorType = 3 
				.Source = sQry
				.ActiveConnection = con
				.Open
			End with
			Do while not rs1.EOF 
				aTrNo = rs1(0)
				aVouNo = rs1(1)
				aVouDate = rs1(2)
											
				Set AdvNode = oDOM.createElement("Advance")
				AdvNode.setAttribute "TransNo",aTrNo
				AdvNode.setAttribute "VoucherNo",aVouNo
				AdvNode.setAttribute "VoucherDate",aVouDate
				AdvNode.setAttribute "AmountRec",aAmtRec
				AdvNode.setAttribute "AmountAdj","0"
				AdvNode.setAttribute "AmountToAdj",aAntToAdj
				AdvNode.setAttribute "CreatedTransNo",aCTrNo
				AdvNode.setAttribute "AdvNo",aAdvNo
				AdvNode.setAttribute "AdjType","I"
				AdvNode.setAttribute "ToAccount","0"
				AdvNode.appendchild AdvNode
				rs1.MoveNext 
			loop
			rs1.Close
			rs2.MoveNext 
		loop
		rs2.Close
	End If 'If trim(sBankType) = "SR" then		
'**********************************************************************************************
'********************************** For the Sales Others Function Call********************* 
'**********************************************************************************************
	
	If trim(sBankType) = "OS" then
		sQry = " Select A.CreatedTransNo,A.CreatedVoucherNo,isNull(V.TransactionNumber,0),isNull(V.VoucherNumber,'') from Acc_T_CreatedVoucherHeader A LEFT JOIN Acc_T_VoucherHeader V ON A.CreatedTransno = V.CreatedTransno where A.CreatedTransno = "& iVNo &" "
		
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
	
		IF not rs1.Eof then
			sCTNo = rs1(0)
			sCVNo = rs1(1)
			sTransNo = rs1(2)
			sVouNo = rs1(3)
		End IF
		rs1.Close
		
		Set Root=oDOM.CreateElement("Voucher")
		Root.setAttribute "CreatedTransNo", sCTNo
		Root.setAttribute "CreatedVouNo",sCVNo 
		Root.setAttribute "TransNo", sTransNo 
		Root.setAttribute "VouNo", sVouNo 		
		oDOM.appendchild Root
'=================================creation of Header node=======================	
		Set Hnode = oDOM.createElement("Header")
		Root.appendchild Hnode
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = " select OUDefinitionID,isNull(BankInstrumentType,0) from Acc_T_CreatedVoucherHeader where CreatedTransno = "&iVNo&" "
			.ActiveConnection = con 
			.Open
		End With
		
		Do while not rs1.Eof 
			sOrgId = rs1(0)
			sPType = rs1(1)
					
			with rs2
				.CursorLocation = 3
				.CursorType = 3
				.Source =" Select D.OrgUnitShortDescription,A.BookNumber from DCS_OrganizationUnitDefinitions D,Acc_T_CreatedVoucherHeader A "&_
						 " where A.CreatedTransNo = "& iVNo &" and D.OUDefinitionID =  "&sOrgId&" "
				.ActiveConnection = con 
				.Open
			End With
			
			If not rs2.EOF then
				sName = rs2(0)
				sBkNo = rs2(1)
			End If 
			rs2.Close 
			rs1.movenext
		loop
		rs1.close
		
		
								
		Set OrgNode = oDOM.createElement("Organization")
		OrgNode.setAttribute "OrgId",sOrgId
		OrgNode.setAttribute "AccUnit",""
		OrgNode.text = sName
		Hnode.appendchild OrgNode
			
		with rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source =" Select BookName from VwOrgBookNames V,Acc_T_CreatedVoucherHeader A "&_
					 " Where V. OUDefinitionID =  "&sOrgId&" and V.BookCode = "&sPassCode&" "&_
					 " and A.CreatedTransNo = "& iVNo &"  and V.BookNumber =  "&sBkNo&" "
			.ActiveConnection = con 
			.Open
		End With
				
		If not rs2.EOF then
			sBkName = rs2(0)
		End If 
		rs2.Close 
			
		Set BookNode = oDOM.createElement("Book")
		BookNode.setAttribute "BookId",sBkNo
		BookNode.setAttribute "BKAccHead","0"
		BookNode.setAttribute "BKOtherUnits","1"
		BookNode.text = sBkName
		Hnode.appendchild BookNode
			
		sQry = "select isNull(BankInstrumentNo,0),VoucherAmount from Acc_T_CreatedVoucherHeader where createdTransNo = "&iVNo&" "
		with rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
		
		If not rs2.EOF then
			sBKInsNo = rs2(0)
			tVouAmt = rs2(1)
		End If 
		rs2.Close 
		
		sQry = "Select isNull(BankInstrumentType,0) From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&sBKInsNo&" "
		with rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
		If not rs2.EOF then
			sBKInsType = rs2(0)
		End If 
		rs2.Close 
			
		sQry = " Select S.InvoiceTypeName from Sal_M_InvoiceTypes S,Acc_T_CreatedVoucherHeader A "&_
			   " where  S.InvoiceType = A.BankInstrumentType  and A.CreatedTransNo = "& iVNo &" "
		with rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
		If not rs2.EOF then
			sSalName = rs2(0)
		End If 
		rs2.Close 
		
		Set SalTypNode = oDOM.createElement("SalesType")
		SalTypNode.setAttribute "SalType",sBKInsType
		SalTypNode.text = sSalName
		Hnode.appendchild SalTypNode
				
		sQry = " select PayToRecdFrom,Convert(Varchar,VoucherDate,103),CreatedBy from Acc_T_CreatedVoucherHeader  where CreatedTransNo =  "& iVNo &" "
		with rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
		If not rs2.EOF then
			sInvNo = rs2(0)						
			sInvDate = rs2(1)
			sApprover = rs2(2)
		End If 'If not rs2.EOF then
		rs2.Close 
		
		set SalInvNode = oDOM.createElement("SaleInvoice")
		SalInvNode.setAttribute "InvNo",sInvNo
		SalInvNode.setAttribute "InvDate",sInvDate
		SalInvNode.setAttribute "RefNo",sInvNo
		SalInvNode.setAttribute "Approval","Y"
		SalInvNode.setAttribute "Approver",sApprover
		SalInvNode.setAttribute "SalTrNo",sBKInsNo
		Hnode.appendchild SalInvNode
					
		sQry = " select A.PartyType,A.PartySubType,A.PartyCode,V.SubTypeName from Acc_T_CreatedVoucherHeader A, "&_
			   " VwOrgPartyType V where A.CreatedTransNo = "&iVNo&" and  V.OUDefinitionID = A.OUDefinitionID and "&_
			   " V.PartySubType = A.PartySubType and V.PartyType = A.PartyType "
		With rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con
			.Open
		End with
		IF not rs2.EOF Then
			sParType  = rs2(0)
			sParSubType = rs2(1)
			sParCode = rs2(2)
			sSubTypeName = rs2(3)
		End IF 
		rs2.close
		
		with rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = " Select PartyName from App_M_PartyMaster where PartyCode = "&sParCode&" "
			.ActiveConnection = con
			.Open
		End with
		If not rs2.EOF then
			sPartyName	= rs2(0)
		End If	
		rs2.Close 
		
		Set PartyNode = oDOM.createElement("Party")
		PartyNode.setAttribute "ParType",sParType
		PartyNode.setAttribute "ParSubType",sParSubType
		PartyNode.setAttribute "ParSubTypeName",sSubTypeName
		PartyNode.setAttribute "ParCode",sParCode
		PartyNode.text = sPartyName
		Hnode.appendchild PartyNode
'====================================creation of Detail node============================	
		With rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = " select Sum(amount) from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iVNo&" "
			.ActiveConnection = con
			.Open
		End with
		If not rs2.EOF then
			sBasVal = rs2(0)
		End If
		rs2.Close
		
		With rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = " select Convert(Varchar,VoucherDate,103) from Acc_T_CreatedVoucherHeader where CreatedTransNo =  "&iVNo&" "
			.ActiveConnection = con
			.Open
		End with
		If not rs2.EOF then
			sVouDate = rs2(0)
		End If
		rs2.Close
					
		Set DetailNode = oDOM.createElement("Details")
		DetailNode.setAttribute "BasicValue",sBasVal
		DetailNode.setAttribute "Discount","0"
		DetailNode.setAttribute "ActualValue",sBasVal
		DetailNode.setAttribute "VouDate",sVouDate
		Root.appendchild DetailNode
		
		sQry = " Select VoucherEntryNumber,ItemDescription,Amount,InvoicedQuantity,InvoicedUOM,InvoicedRate,BasicAmount, "&_
				" DiscountPercent,DiscountAmount,isNull(VoucherNarration,''),ItemCode,ClassificationCode from Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iVNo &" "
				Response.Write "<p>"& sQry 
		with rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con
			.Open
		End With
		If not rs2.EOF then
		    do while not rs2.EOF 
			eNo = rs2(0)
			ePayTo = rs2(1)
			eAmt = rs2(2)
			eQty = rs2(3)
			eUOM = rs2(4)
			eRate = rs2(5)
			eActVal = rs2(6)
			eDisPer = rs2(7)
			eDisAmt = rs2(8)
			sVouNarr = rs2(9)
			eItemCode = rs2(10)
			eClassCode = rs2(11)
			    
		    Set EntryNode = oDOM.createElement("Entry")
		    EntryNode.setAttribute "No",eNo
		    EntryNode.setAttribute "PayTo",ePayTo
		    EntryNode.setAttribute "Amount",eAmt
		    EntryNode.setAttribute "Qty",eQty
		    EntryNode.setAttribute "UOM",eUOM
		    EntryNode.setAttribute "UOMValue",eUOM
		    EntryNode.setAttribute "Rate",eRate
		    EntryNode.setAttribute "ActValue",eActVal
		    EntryNode.setAttribute "DisPer",eDisPer
		    EntryNode.setAttribute "DisAmount",eDisAmt
		    EntryNode.setAttribute "ItemCode",eItemCode
		    EntryNode.setAttribute "ClassCode",eClassCode 
		    DetailNode.appendchild EntryNode
							
		With rs4
			.CursorLocation = 3
			.CursorType = 3
			.Source =" Select isNull(AccUnitAccountHead,0)from Acc_T_CreatedVoucherDetails where CreatedtransNo = "&iVNo&" "&_
					 " and VoucherEntryNumber = "&eNo&" "
			.ActiveConnection = con
			.Open
		End with	
		Do while not rs4.EOF 
			sAccHeadNo = rs4(0)
			With rs3
				.CursorLocation = 3
				.CursorType = 3
				.Source = "select Count(1) from Acc_T_CreatedVoucherCCDet where CreatedTransNo = "&iVNo&" "&_
						  " and AccUnitAccountHead = "&sAccHeadNo&" "
				.ActiveConnection = con
				.Open
			End With
			If not rs3.EOF	then
				sCostCnt = rs3(0)
			End If		
			rs3.Close									
			
			With rs3
				.CursorLocation = 3
				.CursorType = 3
				.Source = "Select Count(1) from Acc_T_CretedVoucherAHDet where CreatedTransNo =	 "&iVNo&" "&_
						   " and AccUnitAccountHead = "&sAccHeadNo&" "
				.ActiveConnection = con
				.Open
			End With
			If not rs3.EOF	then
				sAnlCnt = rs3(0)
			End If		
			rs3.Close
			
			With rs3
				.CursorLocation = 3
				.CursorType = 3
				.Source = "select AccountDescription from Acc_M_GLAccountHead where AccountHead = "& sAccHeadNo &" "
				.ActiveConnection = con
				.Open
			End With
			If not rs3.EOF	then
				sAccHeadName = rs3(0)
			End If		
			rs3.Close
									
			Set AccHeadNode = oDOM.createElement("AccHead")
			AccHeadNode.setAttribute "No",sAccHeadNo
			AccHeadNode.setAttribute "CostCenter",sCostCnt
			AccHeadNode.setAttribute "Analytical",sAnlCnt
			AccHeadNode.setAttribute "Name",sAccHeadName
			AccHeadNode.setAttribute "Type","G"
			AccHeadNode.setAttribute "Group",""
			EntryNode.Appendchild AccHeadNode
			rs4.Movenext
		loop
		rs4.close
	rs2.MoveNext 
	loop
End If
rs2.Close 
		
		Set TaxDetNode = oDOM.createElement("TaxDetails")
		TaxDetNode.setAttribute "InvoiceVlaue",tVouAmt
		TaxDetNode.setAttribute "BasicValue",sBasVal
		TaxDetNode.setAttribute "NettValue",tVouAmt
		TaxDetNode.setAttribute "RoundOffValue",iRoundOff
		Root.appendchild TaxDetNode
		
		sQry =  " Select V.TaxShortName,V.TaxCategoryCode,V.TaxCode,V.ComputationMode,isnull(V.SumOfFields,'') Formula,"&_
				" V.AccountHead,isNull(V.RoundOff,0) RoundOff ,T.TaxAMount,isnull(V.FlatAmount,0) TaxValue,TransCrDrIndication from "&_ 
				" VwSalesTaxDetails V,Acc_T_CreatedVoucherTaxDet T where V.ComputationMode is not null and  "&_
				" V.OUDefinitionID= "&sOrgId&" and V.InvoiceType=T.InvoiceType and V.TaxCategoryCode = T.TaxCategoryCode "&_
				" and V.TaxCode = T.TaxCode and T.CreatedTransNo = "&iVNo&" order by V.TaxHierarchy "
					
		'Response.Write sqry 
		'Response.End
		
		with rs4
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con
			.Open
		End With
		
		Do while  not rs4.Eof 
			sTaxShtName = rs4(0)
			sCatCode = rs4(1)
			sTaxCode = rs4(2)
			sTaxMode = rs4(3)
			sTaxFormula = rs4(4)
			sTaxValue = rs4(8)
			sTaxAmount = rs4(7)
			sAccHead = rs4(5)
			sRoundOff = rs4(6)
			
			IF Cstr(rs4("TransCrDrIndication")) = "C" Then
				sTaxAmount = CDbl(sTaxAmount) * - 1
			End IF
									
			set TaxNode = oDOM.createElement("Tax")
			TaxNode.setAttribute "CatCode",sCatCode
			TaxNode.setAttribute "TaxCode",sTaxCode
			TaxNode.setAttribute "TaxMode",sTaxMode
			TaxNode.setAttribute "TaxFormula",sTaxFormula
			TaxNode.setAttribute "TaxValue",sTaxValue
			TaxNode.setAttribute "TaxAmount",sTaxAmount
			TaxNode.setAttribute "AccHead",sAccHead
			TaxNode.setAttribute "ItemValue","0"
			TaxNode.setAttribute "RoundOff",sRoundOff
			TaxNode.text = sTaxShtName					
			TaxDetNode.appendchild TaxNode
			rs4.MoveNext
		loop
		rs4.Close
			
			NewTaxAmt = 0
			sQry = "Select TaxAmount,TransCrDrIndication,AccountHead from Acc_T_CreatedVoucherTaxDet where "&_
					" TaxCategoryCode = 0 and TaxCode = 0 and CreatedTransNo = "&iVNo&" "
			with rs4
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQry
				.ActiveConnection = con
				.Open
			End With
			If not rs4.EOF then
				NewTaxAmt = rs4(0)
				sAccHead = rs4(2)
					
				IF CStr(rs4("TransCrDrIndication")) = "C" Then
			    	NewTaxAmt = CDbl(NewTaxAmt) * - 1
				End IF
			End If		
			rs4.close
			set TaxNode = oDOM.createElement("Tax")
			TaxNode.setAttribute "CatCode","0"
			TaxNode.setAttribute "TaxCode","0"
			TaxNode.setAttribute "TaxMode","0"
			TaxNode.setAttribute "TaxFormula","0"
			TaxNode.setAttribute "TaxValue","0"
			TaxNode.setAttribute "TaxAmount",NewTaxAmt
			TaxNode.setAttribute "AccHead",sAccHead
			TaxNode.text = "ROUND OFF"					
			TaxDetNode.appendchild TaxNode
			
								
		sQry =  " SELECT A.CreatedAdvanceNo,V.CreatedVoucherNo, V.VoucherDate, V.VoucherAmount, "&_
				" P.AmountReceived, V.CreatedTransNo FROM Acc_T_CreatedVoucherHeader V INNER JOIN "&_
				" Acc_T_CreatedAdvances A ON V.CreatedTransNo = A.CreatedTransNo INNER JOIN "&_
				" Acc_T_CreatedRcvbleAdjDet P ON A.CreatedAdvanceNo = P. ReceivableNumber WHERE "&_
				" P.AdjustType IS NOT NULL AND V.BookCode IN ('01', '02') and "&_
				" P.CreatedTransNo = "&iVNo&" "
								
		set AdvNode = oDOM.createElement("AdvanceDetails")
		Root.appendchild AdvNode
		
		with rs3
			.CursorLocation = 3
			.CursorType = 3 
			.Source = sQry
			.ActiveConnection = con
			.Open
		End with
		Do while not rs3.EOF 
			aAdvNo = rs3(0)
			aAmtRec = rs3(3)
			aAntToAdj = rs3(4)
			dCTransNo = rs3(5)
			aCTrNo = rs3(6)
								
			sQry = " Select TransactionNumber,VoucherNumber,Convert(Varchar,VoucherDate,103) from Acc_T_VoucherHeader where createdTransNo = "&dCTransNo&" "
			with rs1
				.CursorLocation = 3
				.CursorType = 3 
				.Source = sQry
				.ActiveConnection = con
				.Open
			End with
			Do while not rs1.EOF 
				aTrNo = rs1(0)
				aVouNo = rs1(1)
				aVouDate = rs1(2)
											
				set AdvNode = oDOM.createElement("Advance")
				AdvNode.setAttribute "TransNo",aTrNo
				AdvNode.setAttribute "VoucherNo",aVouNo
				AdvNode.setAttribute "VoucherDate",aVouDate
				AdvNode.setAttribute "AmountRec",aAmtRec
				AdvNode.setAttribute "AmountAdj","0"
				AdvNode.setAttribute "AmountToAdj",aAntToAdj
				AdvNode.setAttribute "CreatedTransNo",aCTrNo
				AdvNode.setAttribute "AdvNo",aAdvNo
				AdvNode.setAttribute "AdjType","I"
				AdvNode.setAttribute "ToAccount","0"
				AdvNode.appendchild AdvNode
											
				rs1.MoveNext 
			loop
			rs1.Close
			rs3.MoveNext 
		loop
		rs3.Close
		Set AdvNode = oDOM.createElement("Narration")
		AdvNode.Text = sVouNarr
		Root.appendchild AdvNode
		
		
	End If 'If trim(sBankType) = "OS" then
'**********************************************************************************************
'********************************** For the Purchase Vouchers Function Call********************* 
'**********************************************************************************************
	
	
	If trim(sBankType) = "PA" then
		sQry = " Select A.CreatedTransNo,A.CreatedVoucherNo,isNull(V.TransactionNumber,0),isNull(V.VoucherNumber,'') from Acc_T_CreatedVoucherHeader A LEFT JOIN Acc_T_VoucherHeader V ON A.CreatedTransno = V.CreatedTransno where A.CreatedTransno = "& iVNo &" "
		
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
	
		IF not rs1.Eof then
			sCTNo = rs1(0)
			sCVNo = rs1(1)
			sTransNo = rs1(2)
			sVouNo = rs1(3)
		End IF
		rs1.Close
	
		set Root=oDOM.CreateElement("Voucher")
		Root.setAttribute "CreatedTransNo", sCTNo
		Root.setAttribute "CreatedVouNo",sCVNo 
		Root.setAttribute "TransNo", sTransNo 
		Root.setAttribute "VouNo", sVouNo 		
		oDOM.appendchild Root
'=================================creation of Header node=======================	
		Set Hnode = oDOM.createElement("Header")
		Root.appendchild Hnode
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = " select OUDefinitionID,isNull(BankInstrumentType,0),VoucherAmount from Acc_T_CreatedVoucherHeader where CreatedTransno = "&iVNo&" "
			.ActiveConnection = con 
			.Open
		End With
		Do while not rs1.Eof 
			sOrgId = rs1(0)
			sPType = rs1(1)
			tVouAmt = rs1(2)
			
			with rs2
				.CursorLocation = 3
				.CursorType = 3
				.Source =" Select D.OrgUnitShortDescription,A.BookNumber from DCS_OrganizationUnitDefinitions D,Acc_T_CreatedVoucherHeader A "&_
						 " where A.CreatedTransNo = "& iVNo &" and D.OUDefinitionID =  "&sOrgId&" "
				.ActiveConnection = con 
				.Open
			End With
			If not rs2.EOF then
				sName = rs2(0)
				sBkNo = rs2(1)
			End If 
			rs2.Close 
		rs1.movenext
		loop
		rs1.close
								
		set OrgNode = oDOM.createElement("Organization")
		OrgNode.setAttribute "OrgId",sOrgId
		OrgNode.setAttribute "AccUnit",""
		OrgNode.text = sName
		Hnode.appendchild OrgNode
			
		with rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source =" Select BookName from VwOrgBookNames V,Acc_T_CreatedVoucherHeader A "&_
					 " Where V. OUDefinitionID =  "&sOrgId&" and V.BookCode = "&sPassCode&" "&_
					 " and A.CreatedTransNo = "& iVNo &"  and V.BookNumber =  "&sBkNo&" "
			.ActiveConnection = con 
			.Open
		End With
		
		If not rs2.EOF then
			sBkName = rs2(0)
		End If 
		rs2.Close 
			
		set BookNode = oDOM.createElement("Book")
		BookNode.setAttribute "BookId",sBkNo
		BookNode.setAttribute "BKAccHead","0"
		BookNode.setAttribute "BKOtherUnits","1"
		BookNode.text = sBkName
		Hnode.appendchild BookNode
			
		sQry = "select BankInstrumentNo,isNull(PurchaseBillType,'') from Acc_T_CreatedVoucherHeader where createdTransNo = "&iVNo&" "
		with rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
			
			If not rs2.EOF then
				sBKInsNo = rs2(0)
				sCrInvType = rs2(1)
			End If 
			rs2.Close 
		sQry = "Select BankInstrumentType From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&sBKInsNo&" "
		with rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
			
		If not rs2.EOF then
			sBKInsType = rs2(0)
		End If 
		rs2.Close 
			
		sQry = " select P.PurchaseTypeName from APP_M_PurchaseTypes P,Acc_T_CreatedVoucherHeader A where  P.PurchaseType = "&sBKInsType&" "&_
		   	  " and A.CreatedTransNo = "& iVNo &" "
			
		with rs2
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQry
				.ActiveConnection = con 
				.Open
		End With
			
		If not rs2.EOF then
			sPName = rs2(0)
		End If 
		rs2.Close 
						
		set PurTypNode = oDOM.createElement("PurchaseType")
		PurTypNode.setAttribute "PurTypeId",sBKInsType
		PurTypNode.text = sPName
		Hnode.appendchild PurTypNode
		
		sQry = "Select PayToRecdFrom,CreatedBy,isNull(FromApplication,0),isNull(OtherApplnTransNo,0) From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&sBKInsNo&" "
		Response.Write "qry="& sQry
		with rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con 
			.Open
		End With
		If not rs2.EOF then
			sFrmApp = rs2(2)
										
			If trim(sFrmApp) = "0"  then
				sTempInvNo = rs2(0)
				sArr = split(sTempInvNo,"-")
				sPurInvNo = sArr(0)
				IF UBound(sArr) > 0 Then
					sPurInvDate = Trim(sArr(1))
				Else
					sPurInvDate = ""
				End IF
				sApprover = rs2(1)
				'Response.Write "sPurInvNo="& sPurInvNo & "<BR><BR>"
			else
				sAppTrNo = rs2(3)
				'Response.Write "sAppTrNo="& sAppTrNo
				with rs3
					.CursorLocation = 3
					.CursorType = 3
					.Source = "Select SuppInvoiceNo,SuppInvoiceDate From RCV_T_INVOICEHEADER Where InvoiceNumber = "&sAppTrNo&" "
					.ActiveConnection = con
					.Open
				End with
				If not rs3.EOF then
					sPurInvNo = rs3(0)
					sPurInvDate = rs3(1)
					'Response.Write "sPurInvNo="& sPurInvNo & "<BR><BR>"
				End If
				rs3.Close
			End If	'If trim(sFrmApp) = "Null" then
		End If 'If not rs2.EOF then
		rs2.Close 
			
		set PurInvNode = oDOM.createElement("PurInvoice")
		PurInvNode.setAttribute "PurInvNo",sPurInvNo
		PurInvNode.setAttribute "PurInvDate",sPurInvDate
		PurInvNode.setAttribute "Approval","Y"
		PurInvNode.setAttribute "Approver",sApprover
		PurInvNode.setAttribute "PurTransNo",sBKInsNo
		PurInvNode.setAttribute "CRNoteType",sCrInvType
		Hnode.appendchild PurInvNode
								
		sQry = " select A.PartyType,A.PartySubType,A.PartyCode,V.SubTypeName from Acc_T_CreatedVoucherHeader A, "&_
			   " VwOrgPartyType V where A.CreatedTransNo = "&iVNo&" and  V.OUDefinitionID = A.OUDefinitionID and "&_
			   " V.PartySubType = A.PartySubType and V.PartyType = A.PartyType "
		With rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQry
			.ActiveConnection = con
			.Open
		End with
		Do while not rs2.EOF 
			sParType  = rs2(0)
			sParSubType = rs2(1)
			sParCode = rs2(2)
			sSubTypeName = rs2(3)
			with rs3
				.CursorLocation = 3
				.CursorType = 3
				.Source = " Select PartyName from App_M_PartyMaster where PartyCode = "&sParCode&" "
				.ActiveConnection = con
				.Open
			End with
			If not rs3.EOF then
				sPartyName	= rs3(0)
			End If	
			rs3.Close 
			rs2.MoveNext 
		loop
		rs2.Close
				
		Set PartyNode = oDOM.createElement("Party")
		PartyNode.setAttribute "ParType",sParType
		PartyNode.setAttribute "ParSubType",sParSubType
		PartyNode.setAttribute "ParSubTypeName",sSubTypeName
		PartyNode.setAttribute "ParCode",sParCode
		PartyNode.text = sPartyName
		Hnode.appendchild PartyNode
'====================================creation of Detail node============================	
		With rs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = " select Sum(amount) from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iVNo&" "
			.ActiveConnection = con
			.Open
		End with
		
		If not rs2.EOF then
			sBasVal = rs2(0)
		End If
		rs2.Close
		With rs3
			.CursorLocation = 3
			.CursorType = 3
			.Source = " select Convert(Varchar,VoucherDate,103),VoucherAmount from Acc_T_CreatedVoucherHeader where CreatedTransNo =  "&iVNo&" "
			.ActiveConnection = con
			.Open
		End with
		If not rs3.EOF then
			sVouDate = rs3(0)
			tVouAmt = rs3(1)
		End If
		rs3.Close
				
		set DetailNode = oDOM.createElement("Details")
		DetailNode.setAttribute "BasicValue",sBasVal
		DetailNode.setAttribute "Discount","0"
		DetailNode.setAttribute "ActualValue",sBasVal
		DetailNode.setAttribute "VouDate",sVouDate
		Root.appendchild DetailNode
'=======================creaino of Entry Node ============================
		sQry = " Select VoucherEntryNumber,ItemDescription,Amount,InvoicedQuantity,InvoicedUOM,InvoicedRate,BasicAmount, "&_
					" DiscountPercent,DiscountAmount,ItemCode,ClassificationCode from Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iVNo &" "
			with rs3
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQry
				.ActiveConnection = con
				.Open
			End With
					
			If not rs3.EOF then
			    do while not rs3.EOF 
				eNo = rs3(0)
				ePayTo = rs3(1)
				eAmt = rs3(2)
				eQty = rs3(3)
				eUOM = rs3(4)
				eRate = rs3(5)
				eActVal = rs3(6)
				eDisPer = rs3(7)
				eDisAmt = rs3(8)
				eItemCode = rs3(9)
				eClassCode = rs3(10)
				    
											
			set EntryNode = oDOM.createElement("Entry")
			EntryNode.setAttribute "No",eNo
			EntryNode.setAttribute "PayTo",ePayTo
			EntryNode.setAttribute "Amount",eAmt
			EntryNode.setAttribute "Qty",eQty
			EntryNode.setAttribute "UOM",eUOM
			EntryNode.setAttribute "UOMValue",eUOM
			EntryNode.setAttribute "Rate",eRate
			EntryNode.setAttribute "ActValue",eActVal
			EntryNode.setAttribute "DisPer",eDisPer
			EntryNode.setAttribute "DisAmount",eDisAmt
			EntryNode.setAttribute "ItemCode",eItemCode
			EntryNode.setAttribute "ClassCode",eClassCode
			DetailNode.appendchild EntryNode
						
			With rs5
				.CursorLocation = 3
				.CursorType = 3
				.Source =" Select isNull(AccUnitAccountHead,0)from Acc_T_CreatedVoucherDetails where CreatedtransNo = "&iVNo&" "&_
						 " and VoucherEntryNumber = "&eNo&" "
				.ActiveConnection = con
				.Open
			End with	
			Do while not rs5.EOF 
				sAccHeadNo = rs5(0)
				With rs4
					.CursorLocation = 3
					.CursorType = 3
					.Source = "select Count(1) from Acc_T_CreatedVoucherCCDet where CreatedTransNo = "&iVNo&" "&_
							  " and AccUnitAccountHead = "&sAccHeadNo&" "
					.ActiveConnection = con
					.Open
				End With
				
				If not rs4.EOF	then
					sCostCnt = rs4(0)
				End If		
				rs4.Close									
																					
				With rs4
					.CursorLocation = 3
					.CursorType = 3
					.Source = "Select Count(1) from Acc_T_CretedVoucherAHDet where CreatedTransNo =	 "&iVNo&" "&_
							   " and AccUnitAccountHead = "&sAccHeadNo&" "
					.ActiveConnection = con
					.Open
				End With
													
				If not rs4.EOF	then
					sAnlCnt = rs4(0)
				End If		
				rs4.Close
														
				With rs4
					.CursorLocation = 3
					.CursorType = 3
					.Source = "select AccountDescription from Acc_M_GLAccountHead where AccountHead = "& sAccHeadNo &" "
					.ActiveConnection = con
					.Open
				End With
				If not rs4.EOF	then
					sAccHeadName = rs4(0)
				End If		
				rs4.Close
				set AccHeadNode = oDOM.createElement("AccHead")
				AccHeadNode.setAttribute "No",sAccHeadNo
				AccHeadNode.setAttribute "CostCenter",sCostCnt
				AccHeadNode.setAttribute "Analytical",sAnlCnt
				AccHeadNode.setAttribute "Name",sAccHeadName
				AccHeadNode.setAttribute "Type","G"
				AccHeadNode.setAttribute "TransFlag","A"
				EntryNode.Appendchild AccHeadNode
				rs5.Movenext
			loop
			rs5.close
			
			rs3.MoveNext 
				loop 
			End If
			rs3.Close 
			
'=========================creation of TaxDetails node ==================
			tVouAmt = 0
			With rs3
				.CursorLocation = 3
				.CursorType = 3
				.Source = " Select VoucherAmount from Acc_T_CreatedVoucherHeader where  CreatedTransNo = "&iVNo&" "
				.ActiveConnection = con
				.Open
			End With
			If not rs3.EOF	then
				tVouAmt = rs3(0)
			End If 
			rs3.Close
			
			With rs3
				.CursorLocation = 3
				.CursorType = 3
				.Source =  "Select TaxAmount,TransCrDrIndication from Acc_T_CreatedVoucherTaxDet where "&_
							" TaxCategoryCode = 0 and TaxCode = 0 and CreatedTransNo = "&iVNo&" "
					
				.ActiveConnection = con
				.Open
			End With
					
			If not rs3.EOF	then
			    iRoundOff = rs3(0)
			    Response.Write "<p>"& iRoundOff 
				IF CStr(rs3("TransCrDrIndication")) = "C" Then
					iRoundOff = cdbl(iRoundOff) * - 1
				End IF
			End If 
			rs3.Close
					
			Set TaxDetNode = oDOM.createElement("TaxDetails")
			TaxDetNode.setAttribute "InvoiceVlaue",tVouAmt
			TaxDetNode.setAttribute "BasicValue",sBasVal
			TaxDetNode.setAttribute "NettValue",tVouAmt
			TaxDetNode.setAttribute "RoundOffValue",iRoundOff
			Root.appendchild TaxDetNode
					
			
'==========================Tax node============================
			sQry = " select V.TaxShortName,V.TaxCategoryCode,V.TaxCode,V.ComputationMode,isnull(V.SumOfFields,'') Formula, "&_
					" V.AccountHead,isNull(V.RoundOff,0) RoundOff ,T.TaxAMount,isnull(V.FlatAmount,0) TaxValue "&_
					" from VwPurchaseTaxDetails V,Acc_T_CreatedVoucherTaxDet T where V.ComputationMode is not null "&_
					" and  V.OUDefinitionID="&sOrgId&" and V.PurchaseType=T.InvoiceType and V.TaxCategoryCode = T.TaxCategoryCode "&_
					" and V.TaxCode = T.TaxCode and T.CreatedTransNo = "&iVNo&" order by V.TaxHierarchy "
			with rs4
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQry
				.ActiveConnection = con
				.Open
			End With
			Do while  not rs4.Eof 
				sTaxShtName = rs4(0)
				sCatCode = rs4(1)
				sTaxCode = rs4(2)
				sTaxMode = rs4(3)
				sTaxFormula = rs4(4)
				sTaxValue = rs4(8)
				sTaxAmount = rs4(7)
				sAccHead = rs4(5)
				sRoundOff = rs4(6)
									
				set TaxNode = oDOM.createElement("Tax")
				TaxNode.setAttribute "CatCode",sCatCode
				TaxNode.setAttribute "TaxCode",sTaxCode
				TaxNode.setAttribute "TaxMode",sTaxMode
				TaxNode.setAttribute "TaxFormula",sTaxFormula
				TaxNode.setAttribute "TaxValue",sTaxValue
				TaxNode.setAttribute "TaxAmount",sTaxAmount
				TaxNode.setAttribute "AccHead",sAccHead
				TaxNode.setAttribute "ItemValue","0"
				TaxNode.setAttribute "RoundOff",sRoundOff
				TaxNode.text = sTaxShtName					
				TaxDetNode.appendchild TaxNode
										
				If trim(sCatCode) <> "0" and trim(sTaxCode) <> "0" then

					sQry = "Select TaxAmount,TransCrDrIndication from Acc_T_CreatedVoucherTaxDet where "&_
							" TaxCategoryCode = 0 and TaxCode = 0 and CreatedTransNo = "&iVNo&" "
							
					'Response.Write sQry
					'Response.End
					with rs3
						.CursorLocation = 3
						.CursorType = 3
						.Source = sQry
						.ActiveConnection = con
						.Open
					End With
					If not rs3.EOF then
						NewTaxAmt = rs3(0)
						IF CStr(rs3("TransCrDrIndication")) = "C" Then
							NewTaxAmt = CDbl(NewTaxAmt) * -1
						End IF
					End If		
					rs3.close
				End If
				rs4.MoveNext
			loop
			rs4.Close
			
			IF CStr(NewTaxAmt) = "" Then
				NewTaxAmt = 0
			End IF
			set TaxNode = oDOM.createElement("Tax")
			TaxNode.setAttribute "CatCode","0"
			TaxNode.setAttribute "TaxCode","0"
			TaxNode.setAttribute "TaxMode","0"
			TaxNode.setAttribute "TaxFormula","0"
			TaxNode.setAttribute "TaxValue","0"
			TaxNode.setAttribute "TaxAmount",NewTaxAmt
			TaxNode.setAttribute "AccHead","0"
			TaxNode.text = "ROUND OFF"					
			TaxDetNode.appendchild TaxNode
	End If 'If trim(sBankType) = "PA" then

	
	Response.Clear
	'Response.Write sBankType
	IF CStr(sDispTy) <> "S" Then
		Response.ContentType = "text/XML"
		Response.Write oDOM.xml
	Else
		Response.ContentType = "text/HTML"
		oDOM.save Server.MapPath("../temp/transaction/"&sCTNo&".xml")
		GetXmlForCN = "../temp/transaction/"&sCTNo&".xml"	
	End IF	
End Function 
%>
	
