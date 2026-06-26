<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
	'Program Name				:	GetDNVoucherxml.asp
	'Module Name				:	ACCOUNTS (Reports)
	'Author Name				:	Maheshwari S.
	'Created On					:	Oct 27, 2006
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

	'Dim iVNo,rs1,sBkCode
	
	'Set rs1 = Server.CreateObject("ADODB.Recordset")
	
	'iVNo = Request("Para")
	'Response.Write iVNo
	
	'With rs1
	'	.CursorLocation = 3
	'	.CursorType = 3
	'	.Source =  "Select BookCode From Acc_T_CreatedVoucherHeader Where CreatedTransNo = " & iVNo & " "
	'	.ActiveConnection = con 
	'	.Open
	'End With
	'If not rs1.EOF then	
	'	sBkCode = rs1(0)
	'End If
	'rs1.Close 
	'Response.Write "BookCode = "& sBkCode & "<BR><BR>"
	'GetXml(sBkCode)

Function GetXmlForDN(iVno,sDispTy)
Dim Root,objhttp,rs1,rs2,Advrs,rs3,rs4,rs5,Hnode,sName,sQry,OrgNode,oDOM,NewElem,fsobj,iCount,i
Dim sOrgId,sBkNo,sBkName,sCrDr,sVDate,sBkAccNo,sTransNo,sVouNo,sCTNo,sCVNo,eCrDr
Dim iENo,iEtCrDr,iPayToRecdFrom,iEAmt,iEAccUnit,sAccName,iETdsAmt,iETdsPer,SalTypNode
Dim sAccHeadNo,sAccUnit,sAccPartyType,sAccPartySubType,sAccPartyCode,sPartyName,sSubTypeName
Dim sTempNo,BookNode,PurTypNode,PurInvNode,PartyNode,AccHeadNode,TaxNode,sAccHeadName,sPassCode,sPType,sPName
Dim sTempInvNo,sArr,sPurInvNo,sPurInvDate,sInvNo,sInvDate,sApprover,sFrmApp,iBkCode,iRoundOff
Dim sParType,sParSubType,sParCode,sCostCnt,sAnlCnt,sAppTrNo,sBankType,sBKInsNo,sBKInsType,sRefNo
Dim DetailNode,sBasVal,sVouDate,EntryNode,TaxDetNode,tVouAmt,NewTaxAmt,sSalName,SalInvNode
Dim count,eNo,ePayTo,eAmt,eQty,eUOM,eRate,eActVal,eDisPer,eDisAmt,eItemCode,eClassCode,sCount,AgentDetNode,AgentNode,sAgent
Dim sCatCode,sTaxCode,sTaxMode,sTaxFormula,sTaxValue,sTaxAmount,sAccHead,sTaxShtName,sRoundOff
Dim AdvNode,aAdvNo,aAmtRec,aAntToAdj,dCTransNo,aCTrNo,aTrNo,aVouNo,aVouDate,eTdsAmt,eTdsPer,sNarr

	Set rs1 = server.CreateObject("ADODB.RecordSet")
	Set rs2 = server.CreateObject("ADODB.RecordSet")
	Set rs3 = server.CreateObject("ADODB.RecordSet")
	Set rs4 = server.CreateObject("ADODB.RecordSet")
	Set rs5 = server.CreateObject("ADODB.RecordSet")
	
	Set Advrs = server.CreateObject("ADODB.RecordSet")

	set oDOM=Server.CreateObject("Microsoft.XMLDOM")
	set fsobj=Server.CreateObject("Scripting.FileSystemObject")
	set Root=oDOM.DocumentElement
			
	sPassCode = "06"
	
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
		Response.Write sBankType
		
		
	If trim(sBankType) = "OT" then
		sQry = " Select A.CreatedTransNo,A.CreatedVoucherNo,isNull(V.TransactionNumber,0),isNull(V.VoucherNumber,'') from Acc_T_VoucherHeader V,Acc_T_CreatedVoucherHeader A where A.CreatedTransno = "& iVNo &" "&_
			   " and A.CreatedTransno *= V.CreatedTransno "
		'Response.Write sQry & "<BR><BR>"
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
		
		IF CStr(sTransNo) = "0" Then
			sTransNo = sCTNo 
			sVouNo = sCVNo
		End IF
			
		With rs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = " select A.OUDefinitionID,D.OrgUnitShortDescription,A.BookNumber,A.CrDrIndication,Convert(Varchar,A.VoucherDate,103) from Acc_T_CreatedVoucherHeader A,"&_
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
				
				sTempNo = sParType & "?" & sParSubType & "?" & sSubTypeName & "?" & sParCode
							
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
							
			set Root=oDOM.CreateElement("voucher")
			Root.setAttribute "UnitNo",sOrgId
			Root.setAttribute "UnitName",sName
			Root.setAttribute "BookNo",sBkNo
			Root.setAttribute "BookName",sBkName
			Root.setAttribute "CRDR",sCrDr
			Root.setAttribute "VouDate",sVouDate
			Root.setAttribute "PartyCode",sTempNo
			Root.setAttribute "PartyName",sPartyName
			Root.setAttribute "Approver","S"
			Root.setAttribute "TransNo",sTransNo
			Root.setAttribute "VoucherNo",sVouNo
			Root.setAttribute "CreatedTransNo",sCTNo
			Root.setAttribute "CreatedVoucherNo",sCVNo
			oDOM.appendchild Root
			'=======================creation of Entry Node ============================
			
							sQry = " Select VoucherEntryNumber,TransCrDrIndication,isNull(ItemDescription,0), "&_
									" Amount,ItemCode,ClassificationCode from Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iVNo &" "
							with rs3
								.CursorLocation = 3
								.CursorType = 3
								.Source = sQry
								.ActiveConnection = con
								.Open
							End With
							
							Do WHile Not rs3.EOF
								eNo = rs3(0)
								eCrDr = rs3(1)
								ePayTo = rs3(2)
								eAmt = rs3(3)
								eItemCode = rs3(4)
								eClassCode = rs3(5)
													
								set EntryNode = oDOM.createElement("Entry")
								EntryNode.setAttribute "No",eNo
								EntryNode.setAttribute "CRDR",eCrDr
								EntryNode.setAttribute "Payto",ePayTo
								EntryNode.setAttribute "Amount",eAmt
								EntryNode.setAttribute "AccUnit",sOrgId
								EntryNode.setAttribute "AccName",sName
								EntryNode.setAttribute "ItemCode",eItemCode
								EntryNode.setAttribute "ClassCode",eClassCode
								Root.appendchild EntryNode
										
								With rs4
									.CursorLocation = 3
									.CursorType = 3
									.Source =" Select isNull(AccUnitAccountHead,0)from Acc_T_CreatedVoucherDetails where CreatedtransNo = "&iVNo&" "&_
											 " and VoucherEntryNumber = "&eNo&" "
									.ActiveConnection = con
									.Open
								End with	
								IF not rs4.EOF Then
								    do while not rs4.EOF
									sAccHeadNo = rs4(0)
									With rs5
										.CursorLocation = 3
										.CursorType = 3
										.Source = "select Count(1) from Acc_T_CreatedVoucherCCDet where CreatedTransNo = "&iVNo&" "&_
												  " and AccUnitAccountHead = "&sAccHeadNo&" "
										.ActiveConnection = con
										.Open
									End With
									If not rs5.EOF	then
										sCostCnt = rs5(0)
									End If		
									rs5.Close									
																							
									With rs5
										.CursorLocation = 3
										.CursorType = 3
										.Source = "Select Count(1) from Acc_T_CretedVoucherAHDet where CreatedTransNo =	 "&iVNo&" "&_
												   " and AccUnitAccountHead = "&sAccHeadNo&" "
										.ActiveConnection = con
										.Open
									End With
															
									If not rs5.EOF	then
										sAnlCnt = rs5(0)
									End If		
									rs5.Close
																
									With rs5
										.CursorLocation = 3
										.CursorType = 3
										.Source = "select AccountDescription from Acc_M_GLAccountHead where AccountHead = "& sAccHeadNo &" "
										.ActiveConnection = con
										.Open
									End With
									If not rs5.EOF	then
										sAccHeadName = rs5(0)
									End If		
									rs5.Close
										
									set AccHeadNode = oDOM.createElement("AccHead")
									AccHeadNode.setAttribute "No",sAccHeadNo
									AccHeadNode.setAttribute "CostCenter",sCostCnt
									AccHeadNode.setAttribute "Analytical",sAnlCnt
									AccHeadNode.setAttribute "Name",sAccHeadName
									AccHeadNode.setAttribute "Type","G"
									AccHeadNode.setAttribute "TransFlag","A"
									EntryNode.Appendchild AccHeadNode
									    rs4.MoveNext 
									loop
								End IF
								rs4.Close
									
									with rs4
										.CursorLocation = 3
										.CursorType = 3
										.Source = "select VoucherNarration from Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iVNo &" "&_
												  " and VoucherEntryNumber = "&eNo&" "
										.ActiveConnection = con
										.Open
									End with
									If not rs4.EOF then
										sNarr = rs4(0)
									End If 
									rs4.close
																			
									Set PartyNode = oDom.CreateElement("Narration")
									PartyNode.text = sNarr
									EntryNode.appendchild PartyNode
									rs3.Movenext
							loop
							rs3.close
											
						
	End IF 'If trim(sBankType) = "OT" then
	
	If trim(sBankType) = "OP" or trim(sBankType) = "PR" then
			
		sQry = " Select A.CreatedTransNo,A.CreatedVoucherNo,isNull(V.TransactionNumber,0),isNull(V.VoucherNumber,'')  from Acc_T_VoucherHeader V,Acc_T_CreatedVoucherHeader A where A.CreatedTransno = "& iVNo &" "&_
			   " and A.CreatedTransno *= V.CreatedTransno "
		'Response.Write sQry & "<BR><BR>"

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
		
		IF CStr(sTransNo) = "0" Then
			sTransNo = sCTNo 
			sVouNo = sCVNo
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
				
				sQry = "select isNull(BankInstrumentNo,0) from Acc_T_CreatedVoucherHeader where createdTransNo = "&iVNo&" "
				with rs2
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQry
					.ActiveConnection = con 
					.Open
				End With				
				If not rs2.EOF then
					sBKInsNo = rs2(0)
				Else
					sBKInsNo = 0
				End If 
				rs2.Close 
				
				sQry = "Select isNull(BankInstrumentType,'') From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&sBKInsNo&" "
				
				with rs2
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQry
					.ActiveConnection = con 
					.Open
				End With
				
					If not rs2.EOF then
						sBKInsType = rs2(0)
					Else
						sBKInsType = "0"
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
			
				sQry = "Select PayToRecdFrom,CreatedBy,isNull(FromApplication,0),OtherApplnTransNo From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&sBKInsNo&" "
				'Response.Write "qry="& sQry
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
							sPurInvDate = Trim(sArr(1))
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
					PurInvNode.setAttribute "CrTransNo",sBKInsNo
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
					
					set PartyNode = oDOM.createElement("Party")
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
						.Source = " select Sum(Amount) from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iVNo&" "
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
						.Source = " select Convert(Varchar,VoucherDate,103) from Acc_T_CreatedVoucherHeader where CreatedTransNo =  "&iVNo&" "
						.ActiveConnection = con
						.Open
					End with
					If not rs3.EOF then
						sVouDate = rs3(0)
					End If
					rs3.Close
					
					set DetailNode = oDOM.createElement("Details")
					DetailNode.setAttribute "BasicValue",sBasVal
					DetailNode.setAttribute "Discount","0"
					DetailNode.setAttribute "ActualValue",sBasVal
					DetailNode.setAttribute "VouDate",sVouDate
					Root.appendchild DetailNode
				'=======================creaino of Entry Node ============================
					'sQry = " Select Count(1) from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iVNo&" "
					
					'With rs2
					'	.CursorLocation = 3
					'	.CursorType = 3
					'	.Source = sQry
					'	.ActiveConnection = con
					'	.Open
					'End With
					'If not rs2.Eof then
					'	Count = rs2(0)
					'End If 
					'rs2.Close
					
					'For i = 1 to Count 
						sQry = " Select VoucherEntryNumber,ItemDescription,Amount,InvoicedQuantity,InvoicedUOM,InvoicedRate,BasicAmount, "&_
								" DiscountPercent,DiscountAmount,ItemCode,ClassificationCode from Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iVNo &" "
						with rs3
							.CursorLocation = 3
							.CursorType = 3
							.Source = sQry
							.ActiveConnection = con
							.Open
						End With
						
						'If not rs3.EOF then
						Do While Not rs3.EOF
							eNo = rs3(0)
							ePayTo = rs3(1)
							eAmt = rs3(2)
							eQty = rs3(3)
							eUOM = rs3(4)
							eRate = rs3(5)
							'eActVal = rs3(6)
							eDisPer = rs3(7)
							eDisAmt = rs3(8)
							eItemCode = rs3(9)
							eClassCode = rs3(10)
						'End If
						'rs3.Close 
						
						'sQry = "select Amount  from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&sBKInsNo&" "
						'	with rs3
						'		.CursorLocation = 3
						'		.CursorType = 3
						'		.Source = sQry
						'		.ActiveConnection = con
						'		.Open
						'	End With
							
						'	If not rs3.EOF then
						'		eActVal = rs3(0)					
						'	End If
						'	rs3.Close
							
												
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
							
							With rs4
								.CursorLocation = 3
								.CursorType = 3
								.Source =" Select isNull(AccUnitAccountHead,0)from Acc_T_CreatedVoucherDetails where CreatedtransNo = "&iVNo&" "&_
										 " and VoucherEntryNumber = "&eNo&" "
								.ActiveConnection = con
								.Open
							End with	
							IF Not rs4.EOF Then
							'Do while not rs4.EOF 
								sAccHeadNo = rs4(0)
							End IF
							rs4.Close
							
							
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
								
							rs3.Movenext
						loop
						rs3.close
										
					'next
					
					With rs3
						.CursorLocation = 3
						.CursorType = 3
						.Source = "Select VoucherNarration From Acc_T_CreatedVoucherDetails Where CreatedTransNo = "&iVNo&" "
						.ActiveConnection = con
						.Open
					End With
					
					If not rs3.EOF	then
						Set EntryNode = oDOM.createElement("Narration")
						EntryNode.text = rs3(0)
						Root.appendChild EntryNode
					ENd IF
					rs3.Close
					'=========================creation of TaxDetails node ==================
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
						Response.Write rs3.Source
						If not rs3.EOF	then
							IF CStr(rs3(1)) = "D" Then
								iRoundOff = rs3(0)
								iRoundOff = Cdbl(iRoundOff) * - 1
							Else
								iRoundOff = rs3(0)
							End IF
						End If 
						rs3.Close
														
								Set TaxDetNode = oDOM.createElement("TaxDetails")
								TaxDetNode.setAttribute "InvoiceVlaue",tVouAmt
								TaxDetNode.setAttribute "BasicValue",sBasVal
								TaxDetNode.setAttribute "NettValue",sBasVal
								TaxDetNode.setAttribute "RoundOffValue",iRoundOff
								Root.appendchild TaxDetNode
									
								'==========================Tax node============================
									sQry = " select V.TaxShortName,V.TaxCategoryCode,V.TaxCode,V.ComputationMode,isnull(V.SumOfFields,'') Formula, "&_
											" T.AccountHead,isNull(V.RoundOff,0) RoundOff ,T.TaxAMount,isnull(V.FlatAmount,0) TaxValue "&_
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
									Do While not rs4.Eof 
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
										
										rs4.MoveNext
									loop
									rs4.Close
											
									If trim(sCatCode) <> "0" and trim(sTaxCode) <> "0" then

										sQry = " Select AccountHead,TaxAmount,TransCrDrIndication from Acc_T_CreatedVoucherTaxDet where "&_
												" TaxCategoryCode = 0 and TaxCode = 0 and CreatedTransNo = "&iVNo&" "
										with rs3
											.CursorLocation = 3
											.CursorType = 3
											.Source = sQry
											.ActiveConnection = con
											.Open
										End With
										If not rs3.EOF then
											NewTaxAmt = rs3(0)
											sTaxAmount = rs3(1)
											IF CStr(rs3(2)) = "D" Then
												sTaxAmount = Cdbl(sTaxAmount) * - 1
											End IF
										End If		
										rs3.close
									End If
								
									set TaxNode = oDOM.createElement("Tax")
									TaxNode.setAttribute "CatCode","0"
									TaxNode.setAttribute "TaxCode","0"
									TaxNode.setAttribute "TaxMode","0"
									TaxNode.setAttribute "TaxFormula","0"
									TaxNode.setAttribute "TaxValue","0"
									TaxNode.setAttribute "TaxAmount",sTaxAmount
									TaxNode.setAttribute "AccHead",NewTaxAmt
									TaxNode.text = "ROUND OFF"					
									TaxDetNode.appendchild TaxNode
			
		End If 'If trim(sBankType) = "OP" or trim(sBankType) = "PR"then
		If trim(sBankType) = "SI" then
			sQry = " Select A.CreatedTransNo,A.CreatedVoucherNo,isNull(V.TransactionNumber,0),isNull(V.VoucherNumber,'')  from Acc_T_VoucherHeader V,Acc_T_CreatedVoucherHeader A where A.CreatedTransno = "& iVNo &" "&_
		   " and A.CreatedTransno *= V.CreatedTransno "
		'Response.Write sQry & "<BR><BR>"

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
			
			IF CStr(sTransNo) = "0" Then
				sTransNo = sCTNo 
				sVouNo = sCVNo
			End IF
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
			
				sQry = "select isNull(BankInstrumentNo,0) from Acc_T_CreatedVoucherHeader where createdTransNo = "&iVNo&" "
				with rs2
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQry
					.ActiveConnection = con 
					.Open
				End With
			
				If not rs2.EOF then
					sBKInsNo = rs2(0)
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
					   " where  S.InvoiceType = "&sBKInsType&"  and A.CreatedTransNo = "& iVNo &" "
					'Response.Write sQry
			
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
								
					set SalTypNode = oDOM.createElement("SalesType")
					SalTypNode.setAttribute "SalType",sBKInsType
					SalTypNode.text = sSalName
					Hnode.appendchild SalTypNode
					'Response.Write sPassCode 
					'If sPasscode = "04" then
					'sQry = " select CreatedVoucherNo,VoucherDate,CreatedBy from Acc_T_CreatedVoucherHeader  where CreatedTransNo = "&sBKInsNo&" "&_
					'		" and BookCode = "& sPassCode &" "
					'End If
					sQry = " select createdVoucherNo,Convert(Varchar,VoucherDate,103),CreatedBy,PayToRecdFrom From Acc_T_CreatedVoucherHeader Where CreatedTransNo = "&sBKInsNo&" "
						'Response.Write "qry="& sQry
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
							sRefNo = rs2(3)
						End If 'If not rs2.EOF then
						rs2.Close 
			
						set SalInvNode = oDOM.createElement("SaleInvoice")
						SalInvNode.setAttribute "InvNo",sInvNo
						SalInvNode.setAttribute "InvDate",sInvDate
						SalInvNode.setAttribute "RefNo",sRefNo
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
						with rs2
							.CursorLocation = 3
							.CursorType = 3
							.Source = " Select  count(1) from Acc_T_VoucherHeader A,Sal_T_AdditionalAgents S "&_
									  " where A.CreatedTransNo = "& iVNo &" and S.AccTransactionNo = A.TransactionNumber "
							.ActiveConnection = con
							.Open
						End with
							If not rs2.EOF then
								sCount	= rs2(0)
							End If	
							rs2.Close 
						If trim(sCount) <> "0" then
							sAgent = "Y"
						else
							sAgent = "N"
						End If
				
						
						set PartyNode = oDOM.createElement("Party")
						PartyNode.setAttribute "ParType",sParType
						PartyNode.setAttribute "ParSubType",sParSubType
						PartyNode.setAttribute "ParSubTypeName",sSubTypeName
						PartyNode.setAttribute "ParCode",sParCode
						PartyNode.setAttribute "Agent", sAgent
						PartyNode.text = sPartyName
						Hnode.appendchild PartyNode
						
						If trim(sAgent) = "Y" then
							sQry= "Select T.AgentCode,P.PartyName,T.CommissionType,T.CommissionToPay,T.AgentCommission "&_
								  " From Sal_T_AdditionalAgents T,App_M_PartyMaster P Where P.PartyCode = T.AgentCode "&_
								  " and T.AccTransactionNo = "&iVNo&" "
							with rs3
								.CursorLocation = 3
								.CursorType = 3
								.Source = sQry
								.ActiveConnection = con
								.Open
							End with
							
							If not rs3.EOF then
								sAgnCode = rs3(0)
								sAgnname = rs3(1)
								sCommtype = rs3(2)
								sCommToPay = rs3(3)
								sCommVal = rs3(4)
							End If	
							rs3.Close 
								set AgentDetNode = oDOM.CreateElement("AgentDetails")
								Root.Appendchild AgentDetNode
								
								set AgentNode = oDOM.CreateElement("Agent")
								AgentNode.SetAttribute "Agentcode",sAgnCode
								AgentNode.SetAttribute "Agentname",sAgnName
								AgentNode.SetAttribute "Commisiontype",sCommType
								AgentNode.SetAttribute "Commision",sCommToPay
								AgentNode.SetAttribute "CommValue",sCommVal
								AgentDetNode.Appendchild AgentNode
						End If	'If trim(sAgent) = "Y" then
				
				'====================================creation of Detail node============================	
							
						With rs2
							.CursorLocation = 3
							.CursorType = 3
							.Source = " select Sum(Amount) from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iVNo&" "
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
							.Source = " select Convert(Varchar,VoucherDate,103) from Acc_T_CreatedVoucherHeader where CreatedTransNo =  "&iVNo&" "
							.ActiveConnection = con
							.Open
						End with
						If not rs3.EOF then
							sVouDate = rs3(0)
						End If
						rs3.Close
					
						set DetailNode = oDOM.createElement("Details")
						DetailNode.setAttribute "BasicValue",sBasVal
						DetailNode.setAttribute "Discount","0"
						DetailNode.setAttribute "ActualValue",sBasVal
						DetailNode.setAttribute "VouDate",sVouDate
						Root.appendchild DetailNode
					
						sQry = " Select Count(1) from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iVNo&" "
				
						With rs2
							.CursorLocation = 3
							.CursorType = 3
							.Source = sQry
							.ActiveConnection = con
							.Open
						End With
						If not rs2.Eof then
							Count = rs2(0)
						End If 
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
								    
								        sQry = "select Amount  from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&sBKInsNo&" "
							            with rs4
								            .CursorLocation = 3
								            .CursorType = 3
								            .Source = sQry
								            .ActiveConnection = con
								            .Open
							            End With
            							
							            If not rs4.EOF then
								            eActVal = rs4(0)					
							            End If
							            rs4.Close
							            
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
        								
									        With rs5
										        .CursorLocation = 3
										        .CursorType = 3
										        .Source = "select Count(1) from Acc_T_CreatedVoucherCCDet where CreatedTransNo = "&iVNo&" "&_
												          " and AccUnitAccountHead = "&sAccHeadNo&" "
										        .ActiveConnection = con
										        .Open
									        End With
									        If not rs5.EOF	then
										        sCostCnt = rs5(0)
									        End If		
									        rs5.Close									
        																							
									        With rs5
										        .CursorLocation = 3
										        .CursorType = 3
										        .Source = "Select Count(1) from Acc_T_CretedVoucherAHDet where CreatedTransNo =	 "&iVNo&" "&_
												           " and AccUnitAccountHead = "&sAccHeadNo&" "
										        .ActiveConnection = con
										        .Open
									        End With
        															
									        If not rs5.EOF	then
										        sAnlCnt = rs5(0)
									        End If		
									        rs5.Close
        																
									        With rs5
										        .CursorLocation = 3
										        .CursorType = 3
										        .Source = "select AccountDescription from Acc_M_GLAccountHead where AccountHead = "& sAccHeadNo &" "
										        .ActiveConnection = con
										        .Open
									        End With
									        If not rs5.EOF	then
										        sAccHeadName = rs5(0)
									        End If		
									        rs5.Close
        										
									        set AccHeadNode = oDOM.createElement("AccHead")
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
								    rs3.MoveNext 
								loop
							End If
							rs3.Close 
						
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
						.Source =  "Select TaxAmount from Acc_T_CreatedVoucherTaxDet where "&_
									" isNull(TaxCategoryCode,0) = 0 and isNull(TaxCode,0) = 0 and CreatedTransNo =  "&iVno&" "
					
						.ActiveConnection = con
						.Open
					End With
					If not rs3.EOF	then
						iRoundOff = rs3(0)
					End If 
					rs3.Close
													
							Set TaxDetNode = oDOM.createElement("TaxDetails")
							TaxDetNode.setAttribute "InvoiceVlaue",tVouAmt
							TaxDetNode.setAttribute "BasicValue",sBasVal
							TaxDetNode.setAttribute "NettValue",tVouAmt
							TaxDetNode.setAttribute "RoundOffValue",iRoundOff
							Root.appendchild TaxDetNode
								
					
					'--------------------------------------------------- Creation of Tax node -------------------------------------------					
						sQry = " Select V.TaxShortName,V.TaxCategoryCode,V.TaxCode,V.ComputationMode,isnull(V.SumOfFields,'') Formula,"&_
								" V.AccountHead,isNull(V.RoundOff,0) RoundOff ,T.TaxAMount,isnull(V.FlatAmount,0) TaxValue from "&_ 
								" VwSalesTaxDetails V,Acc_T_CreatedVoucherTaxDet T where V.ComputationMode is not null and  "&_
								" V.OUDefinitionID= "&sOrgId&" and V.InvoiceType=T.InvoiceType and V.TaxCategoryCode = T.TaxCategoryCode "&_
								" and V.TaxCode = T.TaxCode and T.CreatedTransNo = "&iVNo&" order by V.TaxHierarchy "
							'Response.Write sQry
						with rs4
							.CursorLocation = 3
							.CursorType = 3
							.Source = sQry
							.ActiveConnection = con
							.Open
						End With
							Do While not rs4.Eof 
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
								
								rs4.MoveNext 
							loop
							rs4.Close
											
									If trim(sCatCode) <> "0" and trim(sTaxCode) <> "0" then

										sQry = "Select AccountHead,isNull(TaxAmount,0),TransCrDrIndication from Acc_T_CreatedVoucherTaxDet where "&_
												" isNull(TaxCategoryCode,0) = 0 and isNull(TaxCode,0) = 0 and CreatedTransNo = "&iVno&" "
										'Response.Write sQry
										with rs3
											.CursorLocation = 3
											.CursorType = 3
											.Source = sQry
											.ActiveConnection = con
											.Open
										End With
										If not rs3.EOF then
											NewTaxAmt = rs3(0)
											sTaxAmount = rs3(1)
											IF CStr(rs3(2)) = "D" Then
												sTaxAmount = Cdbl(sTaxAmount) * - 1
											End IF
										End If		
										rs3.close
										set TaxNode = oDOM.createElement("Tax")
										TaxNode.setAttribute "CatCode","0"
										TaxNode.setAttribute "TaxCode","0"
										TaxNode.setAttribute "TaxMode","0"
										TaxNode.setAttribute "TaxFormula","0"
										TaxNode.setAttribute "TaxValue","0"
										TaxNode.setAttribute "TaxAmount",sTaxAmount 
										TaxNode.setAttribute "AccHead",NewTaxAmt
										TaxNode.text = "ROUND OFF"					
										TaxDetNode.appendchild TaxNode
					
									End If
							
													
				'--------------------------creation of Advance node --------------------------				
								sQry = " SELECT A.CreatedAdvanceNo,V.CreatedVoucherNo, V.VoucherDate, V.VoucherAmount, "&_
									" P.AmountReceived, V.CreatedTransNo FROM Acc_T_CreatedVoucherHeader V INNER JOIN "&_
									" Acc_T_CreatedAdvances A ON V.CreatedTransNo = A.CreatedTransNo INNER JOIN "&_
									" Acc_T_CreatedRcvbleAdjDet P ON A.CreatedAdvanceNo = P. ReceivableNumber WHERE "&_
									" P.AdjustType IS NOT NULL AND V.BookCode IN ('01', '02') and "&_
									" P.CreatedTransNo = "&iVNo&" "
									
									set AdvNode = oDOM.createElement("AdvanceDetails")
									Root.appendchild AdvNode
									
									'Response.Write sQry
										with Advrs
											.CursorLocation = 3
											.CursorType = 3 
											.Source = sQry
											.ActiveConnection = con
											.Open
										End with
										Do while not Advrs.EOF 
											aAdvNo = Advrs(0)
											aAmtRec = Advrs(3)
											aAntToAdj = Advrs(4)
											dCTransNo = Advrs(5)
											aCTrNo = Advrs(6)
									
											sQry = " Select TransactionNumber,VoucherNumber,Convert(Varchar,VoucherDate,103) from Acc_T_VoucherHeader where createdTransNo = "&dCTransNo&" "
											with rs3
												.CursorLocation = 3
												.CursorType = 3 
												.Source = sQry
												.ActiveConnection = con
												.Open
											End with
											Do while not rs3.EOF 
												aTrNo = rs3(0)
												aVouNo = rs3(1)
												aVouDate = rs3(2)
												
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
												
												rs3.MoveNext 
											loop
											rs3.Close
										Advrs.MoveNext 
										
									loop
									Advrs.Close
		End If 'If trim(sBankType) = "SI" then			
		
				
		Response.Clear
		IF CStr(sDispTy) <> "S" Then
			Response.ContentType = "text/XML"
			Response.Write oDOM.xml
		Else
			Response.ContentType = "text/HTML"
			oDOM.save Server.MapPath("../temp/transaction/DebitNt_Voucher_xml_"&Session.SessionID&".xml")
			GetXmlForDN = "../temp/transaction/DebitNt_Voucher_xml_"&Session.SessionID&".xml"	
		End IF	


		
		
	'End IF

End Function 
%>





	
	
