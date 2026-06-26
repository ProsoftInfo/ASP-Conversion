<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GetPSVoucherxml.asp
	'Module Name				:	ACCOUNTS (Reports)
	'Author Name				:	Maheshwari S.
	'Created On					:	Oct 23, 2006
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
%>

<%

'GetXmlForSal "13","T"










'This Function will call the Purchase Invoice XML Creation
Function GetXmlForPur(iVNo,sDispTy)

	Dim Root,objhttp,rs1,rs2,Advrs,rs3,rs4,Hnode,sName,sQry,Elem1,oDOM,NewElem,fsobj,iCount,i
	Dim sOrgId,sBkNo,sBkName,sCrDr,sVDate,sBkAccNo,sTransNo,sVouNo,sCTNo,sCVNo
	Dim iENo,iEtCrDr,iPayToRecdFrom,iEAmt,iEAccUnit,sAccName,iETdsAmt,iETdsPer
	Dim sAccHeadNo,sAccUnit,sAccPartyType,sAccPartySubType,sAccPartyCode,sPartyName,sSubTypeName
	Dim sTempNo,Elem2,Elem3,Elem4,Elem5,Elem6,Elem7,sAccHeadName,sPassCode,sPType,sPName
	Dim sTempInvNo,sArr,sPurInvNo,sPurInvDate,sApprover,sFrmApp,iBkCode,iRoundOff
	Dim sParType,sParSubType,sParCode,sCostCnt,sAnlCnt,sAppTrNo
	Dim DetailNode,sBasVal,sVouDate,EntryNode,TaxNode,tVouAmt,NewTaxAmt
	Dim count,eNo,ePayTo,eAmt,eQty,eUOM,eRate,eActVal,eDisPer,eDisAmt,eItemCode,eClassCode
	Dim sCatCode,sTaxCode,sTaxMode,sTaxFormula,sTaxValue,sTaxAmount,sAccHead,sTaxShtName,sRoundOff
	Dim AdvNode,aAdvNo,aAmtRec,aAntToAdj,dCTransNo,aCTrNo,aTrNo,aVouNo,aVouDate,Elem8
	Dim sBkCode,sCrAgain,sPurCatTy,sPurCatName,sVatElgi

	Set rs1 = server.CreateObject("ADODB.RecordSet")
	Set rs2 = server.CreateObject("ADODB.RecordSet")
	Set rs3 = server.CreateObject("ADODB.RecordSet")
	Set rs4 = server.CreateObject("ADODB.RecordSet")
	Set Advrs = server.CreateObject("ADODB.RecordSet")

	With rs1
		.CursorLocation = 3
		.CursorType = 3
		.Source =  "Select BookCode,isNull(InvCategoryCode,0) From Acc_T_CreatedVoucherHeader Where CreatedTransNo = " & iVNo & " "
		.ActiveConnection = con
		.Open
	End With
	If not rs1.EOF then
		sBkCode = rs1(0)
		sPurCatTy = rs1(1)
	End If
	rs1.Close
	sPassCode = sBkCode

	sPurCatName = ""
	Select Case CStr(sPurCatTy)
		Case "C" sPurCatName = "CAPITAL GOODS"
		Case "E" sPurCatName = "EXEMPTED"
		Case "I" sPurCatName = "IMPORT"
		Case "B" sPurCatName = "INDUSTRIAL INPUT"
		Case "O" sPurCatName = "INTER STATE SALES"
		Case "R" sPurCatName = "LOCAL PURCHASE INPUT(FIRST SCHEDULE)"
		Case "A" sPurCatName = "PURCHASE AFFECTED THROUGH AGENTS/BRANCHES"
		Case "S" sPurCatName = "STOCK RECEIPTS FROM HEAD OFFICE/BRANCHES/PRINCIPALS OUT SIDE THE STATE"
	End Select

'	sQry = " Select A.CreatedTransNo,A.CreatedVoucherNo,isNull(V.TransactionNumber,0),isNull(V.VoucherNumber,'') from Acc_T_VoucherHeader V,Acc_T_CreatedVoucherHeader A where A.CreatedTransno = "& iVNo &" "&_
'		   " and A.CreatedTransno *= V.CreatedTransno "
		   
	sQry = "  SELECT A.CreatedTransNo,A.CreatedVoucherNo,isNull(V.TransactionNumber,0),isNull(V.VoucherNumber,'') FROM Acc_T_CreatedVoucherHeader A LEFT OUTER JOIN "&_
		   " Acc_T_VoucherHeader V ON A.CreatedTransNo = V.CreatedTransNo WHERE A.CreatedTransNo = " & iVNo

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
	'Response.Write "sVDate =" & sVDate & "<BR><BR>"

	'------------------------------------

		set oDOM=Server.CreateObject("Microsoft.XMLDOM")
		set fsobj=Server.CreateObject("Scripting.FileSystemObject")
		'creation of Voucher node
			set Root=oDOM.DocumentElement
			set Root=oDOM.CreateElement("Voucher")
			Root.setAttribute "CreatedTransNo", sCTNo
			Root.setAttribute "CreatedVouNo",sCVNo
			Root.setAttribute "TransNo", sCTNo
			Root.setAttribute "VouNo", sCVNo
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
					.Source =" Select D.OrgUnitShortDescription,isNull(A.BookNumber,0) from DCS_OrganizationUnitDefinitions D,Acc_T_CreatedVoucherHeader A "&_
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

			set Elem1 = oDOM.createElement("Organization")
			Elem1.setAttribute "OrgId",sOrgId
			Elem1.setAttribute "AccUnit",""
			Elem1.text = sName
			Hnode.appendchild Elem1




			with rs2
					.CursorLocation = 3
					.CursorType = 3
					.Source =" Select BookName,isNull(BookAccountHead,0),OtherUnitTransaction from VwOrgBookNames V,Acc_T_CreatedVoucherHeader A "&_
							 " Where V. OUDefinitionID =  "&sOrgId&" and V.BookCode = "&sPassCode&" "&_
							 " and A.CreatedTransNo = "& iVNo &"  and V.BookNumber =  "&sBkNo&" "
					.ActiveConnection = con
					.Open
			End With
				If not rs2.EOF then
						sBkName = rs2(0)
						    ''Add by ragav on Jan 03,2012 for avoid the empty value of selected acc head
						set Elem2 = oDOM.createElement("Book")
				            Elem2.setAttribute "BookId",sBkNo
				            Elem2.setAttribute "BKAccHead",rs2(1)
				            Elem2.setAttribute "BKOtherUnits",rs2(2)
				            Elem2.text = sBkName
				            Hnode.appendchild Elem2
				         ''end
				    else
				        ''Add by ragav on Jan 03,2012 for avoid the empty value of selected acc head
				        set Elem2 = oDOM.createElement("Book")
				            Elem2.setAttribute "BookId",sBkNo
				            Elem2.setAttribute "BKAccHead","0"
				            Elem2.setAttribute "BKOtherUnits","1"
				            Elem2.text = sBkName
				            Hnode.appendchild Elem2
				         ''end
			    	End If
					rs2.Close
					
					''blocked by ragav on Jan 03,2012 for avoid the empty value of selected acc head
		'		set Elem2 = oDOM.createElement("Book")
		'		Elem2.setAttribute "BookId",sBkNo
		'		Elem2.setAttribute "BKAccHead","0"
		'		Elem2.setAttribute "BKOtherUnits","1"
		'		Elem2.text = sBkName
		'		Hnode.appendchild Elem2
		''end 

			sQry = " select P.PurchaseTypeName from APP_M_PurchaseTypes P,Acc_T_CreatedVoucherHeader A where  P.PurchaseType = A.BankInstrumentType "&_
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

			set Elem3 = oDOM.createElement("PurchaseType")
			Elem3.setAttribute "PurTypeId",sPType
			Elem3.text = sPName
			Hnode.appendchild Elem3

			sQry = " Select PayToRecdFrom,CreatedBy,isNull(FromApplication,0),isNull(OtherApplnTransno,0),BookCode,isNull(PayableAt,'0') from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& iVNo &" "&_
					" and BookCode = "& sPassCode &" "
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
					iBkCode = rs2(4)
					sCrAgain = rs2(5)
					'Response.Write "frmapp="& iBkCode & "<BR><BR>"
					If trim(sFrmApp) = "0" and iBkCode = "04" then
						sTempInvNo = rs2(0)
						sArr = split(sTempInvNo,"-")
						sPurInvNo = sArr(0)
						sPurInvDate = Right(sTempInvNo,10)
						sApprover = rs2(1)
						'Response.Write "sPurInvNo="& sPurInvNo & "<BR><BR>"
					else
						sAppTrNo = rs2(3)
						'Response.Write "sAppTrNo="& sAppTrNo
						with rs3
							.CursorLocation = 3
							.CursorType = 3
							.Source = "Select SuppInvoiceNo,convert(Varchar(10),SuppInvoiceDate,103) From RCV_T_INVOICEHEADER Where InvoiceNumber = "&sAppTrNo&" "
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

				set Elem4 = oDOM.createElement("PurInvoice")
				Elem4.setAttribute "PurInvNo",sPurInvNo
				Elem4.setAttribute "PurInvDate",sPurInvDate
				Elem4.setAttribute "Approval","Y"
				Elem4.setAttribute "Approver",sApprover
				Elem4.setAttribute "CrTransNo",sTransNo

				Hnode.appendchild Elem4
				'VwOrgPartyType
				sQry = " select A.PartyType,A.PartySubType,A.PartyCode,V.SubTypeName from Acc_T_CreatedVoucherHeader A, "&_
					   " VwOrgParty V where A.CreatedTransNo = "&iVNo&" and  V.OUDefinitionID = A.OUDefinitionID and "&_
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

				set Elem5 = oDOM.createElement("Party")
				Elem5.setAttribute "ParType",sParType
				Elem5.setAttribute "ParSubType",sParSubType
				Elem5.setAttribute "ParSubTypeName",sSubTypeName
				Elem5.setAttribute "ParCode",sParCode
				Elem5.text = sPartyName
				Hnode.appendchild Elem5

				Set Elem5 = oDOM.createElement("PurCategory")
				Elem5.setAttribute "Code",sPurCatTy
				Elem5.text = sPurCatName
				Hnode.appendchild Elem5

			'====================================creation of Detail node============================

				With rs2
					.CursorLocation = 3
					.CursorType = 3
					.Source = " select Sum(Basicamount) from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iVNo&" "
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

				'For i = 1 to Count
					sQry = " Select VoucherEntryNumber,isNull(ItemDescription,''),Amount,InvoicedQuantity,isNull(InvoicedUOM,''),InvoicedRate,BasicAmount, "&_
							" DiscountPercent,DiscountAmount,ItemCode,ClassificationCode,isNull(InvoiceType,0),isNull(VATEligibility,'N'),TransCrDrIndication from Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iVNo &" "
					'	Response.Write sQry
					with rs3
						.CursorLocation = 3
						.CursorType = 3
						.Source = sQry
						.ActiveConnection = con
						.Open
					End With

					'If not rs3.EOF then
					Do While Not rs3.Eof
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
						sPType = rs3(11)
						sVatElgi = rs3(12)
						sCrDr  = rs3(13)
					'End If
					'rs3.Close

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
					EntryNode.setAttribute "PurType",sPType
					EntryNode.setAttribute "VATElg",sVatElgi
					EntryNode.setAttribute "CatCode",""
					EntryNode.setAttribute "CRDR",sCrDr
					DetailNode.appendchild EntryNode

						With rs4
							.CursorLocation = 3
							.CursorType = 3
							.Source =" Select isNull(AccUnitAccountHead,0)from Acc_T_CreatedVoucherDetails where CreatedtransNo = "&iVNo&" "&_
									 " and VoucherEntryNumber = "&eNo&" "
							.ActiveConnection = con
							.Open
						End with
						'Do while not rs3.EOF
						IF Not rs4.eof Then
							sAccHeadNo = rs4(0)
						End IF
						rs4.close

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

						set Elem6 = oDOM.createElement("AccHead")
						Elem6.setAttribute "No",sAccHeadNo
						Elem6.setAttribute "CostCenter",sCostCnt
						Elem6.setAttribute "Analytical",sAnlCnt
						Elem6.setAttribute "Name",sAccHeadName
						Elem6.setAttribute "Type","G"
						Elem6.setAttribute "TransFlag","A"
						EntryNode.Appendchild Elem6
					rs3.Movenext
				loop
				rs3.close

				'next
				'=========================creation of TaxDetails node ==================
					Dim sPurInvTy
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

					Dim sRndCrDr

					With rs3
						.CursorLocation = 3
						.CursorType = 3
						.Source =  "Select TaxAmount,InvoiceType,TransCrDrIndication from Acc_T_CreatedVoucherTaxDet where "&_
									" isNull(TaxCategoryCode,0) = 0 and isNull(TaxCode,0) = 0 and CreatedTransNo = "&iVNo&" "

						'Response.Write rs3.Source
						.ActiveConnection = con
						.Open
					End With
					If not rs3.EOF	then
						iRoundOff = rs3(0)
						sPurInvTy = rs3(1)
						sRndCrDr = rs3("TransCrDrIndication")
					End If
					rs3.Close
					IF Cstr(sRndCrDr) = "C" Then
						iRoundOff = Cdbl(iRoundOff) * - 1
					End IF
							'Response.Write "iRoundOff="&iRoundOff&"<br>"
							Set TaxNode = oDOM.createElement("TaxDetails")
							TaxNode.setAttribute "InvoiceVlaue",tVouAmt
							TaxNode.setAttribute "BasicValue",sBasVal
							TaxNode.setAttribute "NettValue",tVouAmt
							TaxNode.setAttribute "RoundOffValue",iRoundOff
							TaxNode.setAttribute "PurchaseType",sPurInvTy
							Root.appendchild TaxNode

							'==========================Tax node============================
								sQry = " Select V.TaxShortName,V.TaxCategoryCode,V.TaxCode,V.ComputationMode,isnull(V.SumOfFields,'') Formula, "&_
									   " V.AccountHead,isNull(V.RoundOff,0) RoundOff ,T.TaxAMount,isnull(V.FlatAmount,0) TaxValue,T.InvoiceType,T.TransCrDrIndication from "&_
									   " VwPurchaseTaxDetails V,Acc_T_CreatedVoucherTaxDet T where V.ComputationMode is not null and  "&_
									   " V.OUDefinitionID= "&sOrgId&" and V.PurchaseType = T.InvoiceType and V.TaxCategoryCode = T.TaxCategoryCode "&_
									   " and V.TaxCode = T.TaxCode and T.CreatedTransNo = "&iVNo&" order by V.TaxHierarchy "
								'Response.Write sQry
								'Response.End

								with rs4
									.CursorLocation = 3
									.CursorType = 3
									.Source = sQry
									.ActiveConnection = con
									.Open
								End With

								IF Not rs4.eof Then
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
										sPurInvTy = rs4("InvoiceType")
										sRndCrDr = rs4("TransCrDrIndication")

										'Response.Write sRndCrDr &"<br>"

										IF Cstr(sRndCrDr) = "C" Then
											sTaxAmount = Cdbl(sTaxAmount) * - 1
										End IF

											If trim(sAccHead) = "" then sAccHead = 0
											set Elem7 = oDOM.createElement("Tax")
											Elem7.setAttribute "CatCode",sCatCode
											Elem7.setAttribute "TaxCode",sTaxCode
											Elem7.setAttribute "TaxMode",sTaxMode
											Elem7.setAttribute "TaxFormula",sTaxFormula
											Elem7.setAttribute "TaxValue",sTaxValue
											If cint(sCatCode) = 0 and cint(sTaxCode) = 0 then
												Elem7.setAttribute "TaxAmount",iRoundOff
											Else
												Elem7.setAttribute "TaxAmount",sTaxAmount
											End If
											Elem7.setAttribute "AccHead",sAccHead
											Elem7.setAttribute "ItemValue","0"
											Elem7.setAttribute "RoundOff",sRoundOff
											Elem7.setAttribute "PurchaseType",sPurInvTy
											Elem7.text = sTaxShtName
											TaxNode.appendchild Elem7

											If trim(sCatCode) <> "0" and trim(sTaxCode) <> "0" then

												sQry = "Select TaxAmount,TransCrDrIndication from Acc_T_CreatedVoucherTaxDet where "&_
														" isNull(TaxCategoryCode,0) = 0 and isNull(TaxCode,0) = 0  and CreatedTransNo = "&iVNo&" "
												with rs3
													.CursorLocation = 3
													.CursorType = 3
													.Source = sQry
													.ActiveConnection = con
													.Open
												End With
												If not rs3.EOF then
													NewTaxAmt = rs3(0)
													sRndCrDr = rs3(1)
												End If
												rs3.close
											End If

										rs4.MoveNext
									loop
								Else
									Dim iRndAccHead
									sQry = "Select TaxAmount,TransCrDrIndication,isNull(AccountHead,0) from Acc_T_CreatedVoucherTaxDet where "&_
											" isNull(TaxCategoryCode,0) = 0 and isNull(TaxCode,0) = 0  and CreatedTransNo = "&iVNo&" "
									with rs3
										.CursorLocation = 3
										.CursorType = 3
										.Source = sQry
										.ActiveConnection = con
										.Open
									End With
									If not rs3.EOF then
										NewTaxAmt = rs3(0)
										sRndCrDr = rs3(1)
										iRndAccHead = rs3(2)
									End If
									rs3.close
								End IF
								rs4.Close

								IF Cstr(sRndCrDr) = "C" Then
									NewTaxAmt = Cdbl(NewTaxAmt) * - 1
								End IF
								'Response.Write "NewTaxAmt="&NewTaxAmt
								NewTaxAmt = "0.00"
										If trim(iRndAccHead) = "" then iRndAccHead = 0
										set Elem8 = oDOM.createElement("Tax")
										Elem8.setAttribute "CatCode","0"
										Elem8.setAttribute "TaxCode","0"
										Elem8.setAttribute "TaxMode","0"
										Elem8.setAttribute "TaxFormula","0"
										Elem8.setAttribute "TaxValue","0"
										'Elem8.setAttribute "TaxAmount",NewTaxAmt
										Elem8.setAttribute "TaxAmount",iRoundOff
										Elem8.setAttribute "AccHead",iRndAccHead
										Elem8.text = "ROUND OFF"
										TaxNode.appendchild Elem8


				'===============================creation of Advance detail node ================
								set AdvNode = oDOM.createElement("AdvanceDetails")
								Root.appendchild AdvNode

								sQry = "SELECT A.CreatedAdvanceNo,V.CreatedVoucherNo, V.VoucherDate, V.VoucherAmount,P.AmountAdjusted, "&_
										" V.CreatedTransNo FROM Acc_T_CreatedVoucherHeader V INNER JOIN Acc_T_CreatedAdvances A "&_
										" ON V.CreatedTransNo = A.CreatedTransNo INNER JOIN ACC_T_CREATEDADVANCEADJ P ON "&_
										" A.CreatedAdvanceNo = P.CreatedAdvanceNo WHERE  P.CreatedTransNo = "&iVNo

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
										'aCTrNo = Advrs(6)

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

											set Elem8 = oDOM.createElement("Advance")
											Elem8.setAttribute "TransNo",aTrNo
											Elem8.setAttribute "VoucherNo",aVouNo
											Elem8.setAttribute "VoucherDate",aVouDate
											Elem8.setAttribute "AmountRec",aAmtRec
											Elem8.setAttribute "AmountAdj","0"
											Elem8.setAttribute "AmountToAdj",aAntToAdj
											Elem8.setAttribute "CreatedTransNo",dCTransNo
											Elem8.setAttribute "AdvNo",aAdvNo
											Elem8.setAttribute "AdjType","I"
											Elem8.setAttribute "ToAccount","0"
											AdvNode.appendchild Elem8

											rs3.MoveNext
										loop
										rs3.Close
									Advrs.MoveNext

								loop
								Advrs.Close

				oDOM.save Server.MapPath("../temp/transaction/Pur_Voucher_xml_"&Session.SessionID&".xml")

	Response.Clear
	IF CStr(sDispTy) <> "S" Then
		Response.ContentType = "text/XML"
		Response.Write oDOM.xml
	Else
		Response.ContentType = "text/HTML"
		oDOM.save Server.MapPath("../temp/transaction/Pur_Voucher_xml_"&Session.SessionID&".xml")
		GetXmlForPur = "../temp/transaction/Pur_Voucher_xml_"&Session.SessionID&".xml"
	End IF

End Function
%>

<%
	'This Function will call the Sales Invoice XML Creation
	Function GetXmlForSal(iVNo,sDispTy)
		Dim sBkCode,sCode
		Dim Root,objhttp,rs1,rs2,Advrs,rs3,rs4,rs5,Hnode,sName,sQry,Elem1,oDOM,NewElem,fsobj,iCount,i
		Dim sOrgId,sBkNo,sBkName,sCrDr,sVDate,sBkAccNo,sTransNo,sVouNo,sCTNo,sCVNo
		Dim iENo,iEtCrDr,iPayToRecdFrom,iEAmt,iEAccUnit,sAccName,iETdsAmt,iETdsPer
		Dim sAccHeadNo,sAccUnit,sAccPartyType,sAccPartySubType,sAccPartyCode,sPartyName,sSubTypeName
		Dim sTempNo,Elem2,Elem3,Elem4,Elem5,Elem6,Elem7,sAccHeadName,sPassCode,sSalType,sSalName
		Dim sInvNo,sInvDate,sApprover,iRoundOff,ePackCode,eNoOfPack,eRatePer
		Dim sParType,sParSubType,sParCode,sCostCnt,sAnlCnt,sAppTrNo,sCount,sAgent
		Dim DetailNode,sBasVal,sVouDate,EntryNode,TaxNode,tVouAmt,NewTaxAmt
		Dim count,eNo,ePayTo,eAmt,eQty,eUOM,eRate,eActVal,eDisPer,eDisAmt,eItemCode,eClassCode
		Dim sCatCode,sTaxCode,sTaxMode,sTaxFormula,sTaxValue,sTaxAmount,sAccHead,sTaxShtName,sRoundOff
		Dim AdvNode,aAdvNo,aAmtRec,aAntToAdj,dCTransNo,aCTrNo,aTrNo,aVouNo,aVouDate,Elem8
		Dim sAgnCode,sAgnname,sCommtype,sCommToPay,sCommVal,AgentDetNode,AgentNode
		Dim sAgTy,sAgSubTy,seRndVal
		Set rs1 = server.CreateObject("ADODB.RecordSet")
		Set rs2 = server.CreateObject("ADODB.RecordSet")
		Set rs3 = server.CreateObject("ADODB.RecordSet")
		Set rs4 = server.CreateObject("ADODB.RecordSet")
		Set rs5 = server.CreateObject("ADODB.RecordSet")

		Set Advrs = server.CreateObject("ADODB.RecordSet")
		sCode = "05"
		sPassCode = sCode
		'Response.Write "pass="&sPassCode
		'creation of Voucher node
		
		sQry = "  SELECT  A.CreatedTransNo,A.CreatedVoucherNo,isNull(V.TransactionNumber,0), isNull(V.VoucherNumber,'') FROM Acc_T_CreatedVoucherHeader A LEFT OUTER JOIN "&_
			   " Acc_T_VoucherHeader V ON A.CreatedTransNo = V.CreatedTransNo WHERE A.CreatedTransNo = " & iVNo
       
	'	sQry = " Select A.CreatedTransNo,A.CreatedVoucherNo,isNull(V.TransactionNumber,0),isNull(V.VoucherNumber,'') from Acc_T_VoucherHeader V,Acc_T_CreatedVoucherHeader A where A.CreatedTransno = "& iVNo &" "&_
	'	   " and A.CreatedTransno *= V.CreatedTransno "
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
		'Response.Write "sVDate =" & sVDate & "<BR><BR>"

		'------------------------------------

			set oDOM=Server.CreateObject("Microsoft.XMLDOM")
			set fsobj=Server.CreateObject("Scripting.FileSystemObject")
			'creation of Voucher node
				set Root=oDOM.DocumentElement
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
				 sSalType = rs1(1)

					with rs2
						.CursorLocation = 3
						.CursorType = 3
						.Source =" Select D.OrgUnitShortDescription,isNull(A.BookNumber,0) from DCS_OrganizationUnitDefinitions D,Acc_T_CreatedVoucherHeader A "&_
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

				set Elem1 = oDOM.createElement("Organization")
				Elem1.setAttribute "OrgId",sOrgId
				Elem1.setAttribute "AccUnit",""
				Elem1.text = sName
				Hnode.appendchild Elem1

				with rs2
						.CursorLocation = 3
						.CursorType = 3
						.Source =" Select BookName,BookAccountHead,OtherUnitTransaction from VwOrgBookNames V,Acc_T_CreatedVoucherHeader A "&_
								 " Where V. OUDefinitionID =  "&sOrgId&" and V.BookCode = "&sPassCode&" "&_
								 " and A.CreatedTransNo = "& iVNo &"  and V.BookNumber =  "&sBkNo&" "
						.ActiveConnection = con
						.Open
				End With
					If not rs2.EOF then
						sBkName = rs2(0)
						    ''Add by ragav on Jan 03,2012 for avoid the empty value of selected acc head
						set Elem2 = oDOM.createElement("Book")
				            Elem2.setAttribute "BookId",sBkNo
				            Elem2.setAttribute "BKAccHead",rs2(1)
				            Elem2.setAttribute "BKOtherUnits",rs2(2)
				            Elem2.text = sBkName
				            Hnode.appendchild Elem2
				         ''end
				    else
				        ''Add by ragav on Jan 03,2012 for avoid the empty value of selected acc head
				        set Elem2 = oDOM.createElement("Book")
				            Elem2.setAttribute "BookId",sBkNo
				            Elem2.setAttribute "BKAccHead","0"
				            Elem2.setAttribute "BKOtherUnits","1"
				            Elem2.text = sBkName
				            Hnode.appendchild Elem2
				         ''end
			    	End If
					rs2.Close
					
					''blocked by ragav on Jan 03,2012 for avoid the empty value of selected acc head
		'		set Elem2 = oDOM.createElement("Book")
		'		Elem2.setAttribute "BookId",sBkNo
		'		Elem2.setAttribute "BKAccHead","0"
		'		Elem2.setAttribute "BKOtherUnits","1"
		'		Elem2.text = sBkName
		'		Hnode.appendchild Elem2
		''end 

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

				set Elem3 = oDOM.createElement("SalesType")
				Elem3.setAttribute "SalType",sSalType
				Elem3.text = sSalName
				Hnode.appendchild Elem3

				sQry = " select CreatedVoucherNo,Convert(Varchar,VoucherDate,103),CreatedBy from Acc_T_CreatedVoucherHeader  where CreatedTransNo =  "& iVNo &" "
						'" and BookCode = "& sPassCode &" "
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
					End If 'If not rs2.EOF then
					rs2.Close

					set Elem4 = oDOM.createElement("SaleInvoice")
					Elem4.setAttribute "InvNo",sInvNo
					Elem4.setAttribute "InvDate",sInvDate
					Elem4.setAttribute "RefNo",sInvNo
					Elem4.setAttribute "Approval","Y"
					Elem4.setAttribute "Approver",sApprover
					Elem4.setAttribute "CrTransNo",iVNo
					Hnode.appendchild Elem4

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
					with rs3
						.CursorLocation = 3
						.CursorType = 3
						.Source = " Select  count(1) from Sal_T_AdditionalAgents  "&_
								  " where AccTransactionNo = "& iVNo &" "
						.ActiveConnection = con
						.Open
					End with

					'Response.Write rs3.Source

						If not rs3.EOF then
							sCount	= rs3(0)
						End If
						rs3.Close
					'Response.Write "count="& sCount
					If trim(sCount) <> "0" then
						sAgent = "Y"
					else
						sAgent = "N"
					End If

					set Elem5 = oDOM.createElement("Party")
					Elem5.setAttribute "ParType",sParType
					Elem5.setAttribute "ParSubType",sParSubType
					Elem5.setAttribute "ParSubTypeName",sSubTypeName
					Elem5.setAttribute "ParCode",sParCode
					Elem5.setAttribute "Agent", sAgent
					Elem5.text = sPartyName
					Hnode.appendchild Elem5

					If trim(sAgent) = "Y" then
						sQry= "Select T.AgentCode,P.PartyName,T.CommissionType,T.CommissionToPay,T.AgentCommission, "&_
							  " T.AgentType ,T.AgentSubType From Sal_T_AdditionalAgents T,App_M_PartyMaster P Where P.PartyCode = T.AgentCode "&_
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
							sAgTy = rs3(5)
							sAgSubTy = rs3(6)
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
							AgentNode.SetAttribute "AgentType",sAgTy
							AgentNode.SetAttribute "AgentSubType",sAgSubTy
							AgentDetNode.Appendchild AgentNode
					End If	'If trim(sAgent) = "Y" then


				'====================================Creation of Detail node============================

					With rs2
						.CursorLocation = 3
						.CursorType = 3
						.Source = " select Sum(Basicamount) from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iVNo&" "
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
				'=======================creation of Entry Node ============================
					'sQry = " Select Count(1) from Acc_T_CreatedVoucherDetails where CreatedTransNo = "&iVNo&" "
					'Response.Write	sQry &"<BR><BR>"
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

					'For i = 1 to Count
					i = 0
						sQry = " Select VoucherEntryNumber,isNull(ItemDescription,0),Amount,InvoicedQuantity,isNull(InvoicedUOM,''),InvoicedRate,BasicAmount, "&_
								" DiscountPercent,DiscountAmount,ItemCode,ClassificationCode,RatePer,isNull(NoOfPack,0),isNull(PackCode,0),isNull(RoundOffValue,0),TransCrDrIndication from "&_
								" Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iVNo
								Response.Write "2="&sQry
						with rs3
							.CursorLocation = 3
							.CursorType = 3
							.Source = sQry
							.ActiveConnection = con
							.Open
						End With

						If not rs3.EOF then
						    do while not rs3.eof 
							i = Cint(i) + 1
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
							eRatePer = rs3(11)
							eNoOfPack = rs3(12)
							ePackCode = rs3(13)
							seRndVal = rs3(14)
							sCrDr = rs3(15)
							
							Response.Write "<P>"& ePayTo
							set EntryNode = oDOM.createElement("Entry")
							'EntryNode.setAttribute "No",eNo
							EntryNode.setAttribute "No",i
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
							EntryNode.setAttribute "TransBasicamt",eActVal
							EntryNode.setAttribute "TransRate",eRate
							EntryNode.setAttribute "TransDisAmt",eDisPer
							EntryNode.setAttribute "TransInvAmt",eDisAmt
							EntryNode.setAttribute "RatePer",eRatePer
							EntryNode.setAttribute "NoofPack",eNoOfPack
							EntryNode.setAttribute "PackType",ePackCode
							EntryNode.setAttribute "RndOff",seRndVal
							EntryNode.setAttribute "CRDR",sCrDr

							DetailNode.appendchild EntryNode

							IF CStr(eNo) = "" Then
								eNo = 1
							End IF

							With rs5
								.CursorLocation = 3
								.CursorType = 3
								.Source =" Select isNull(AccUnitAccountHead,0)from Acc_T_CreatedVoucherDetails where CreatedtransNo = "&iVNo

										 '" and VoucherEntryNumber = "&eNo&" "

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

								set Elem6 = oDOM.createElement("AccHead")
								Elem6.setAttribute "No",sAccHeadNo
								Elem6.setAttribute "CostCenter",sCostCnt
								Elem6.setAttribute "Analytical",sAnlCnt
								Elem6.setAttribute "Name",sAccHeadName
								Elem6.setAttribute "Type","G"
								Elem6.setAttribute "TransFlag","A"
								EntryNode.Appendchild Elem6
								rs5.movenext
							loop
						rs5.close
						rs3.Movenext
					loop
					end if'If not rs3.EOF then
					rs3.close

					'next
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
						Dim sTaxCrDr
						With rs3
							.CursorLocation = 3
							.CursorType = 3
							.Source =  "Select TaxAmount,TransCrDrIndication from Acc_T_CreatedVoucherTaxDet where "&_
										" isNull(TaxCategoryCode,0) = 0 and isNull(TaxCode,0) = 0  and CreatedTransNo = "&iVNo&" "

							.ActiveConnection = con
							.Open
						End With



						If not rs3.EOF	then
							iRoundOff = rs3(0)
							sTaxCrDr = rs3(1)
						End If
						rs3.Close

						IF Cstr(sTaxCrDr) = "D" Then
							iRoundOff = Cdbl(iRoundOff) * -1
						End IF

								Set TaxNode = oDOM.createElement("TaxDetails")
								TaxNode.setAttribute "InvoiceVlaue",tVouAmt
								TaxNode.setAttribute "BasicValue",sBasVal
								TaxNode.setAttribute "NettValue",tVouAmt
								TaxNode.setAttribute "RoundOffValue",iRoundOff
								Root.appendchild TaxNode

								'==========================Tax node============================
									Dim iTaxAmount,sTaxCrDrIndi
									sQry = "select V.TaxShortName,V.TaxCategoryCode,V.TaxCode,V.ComputationMode,isnull(V.SumOfFields,'') Formula, "&_
											" V.AccountHead,isNull(V.RoundOff,0) RoundOff ,T.TaxAMount,isnull(V.FlatAmount,0) TaxValue,TransCrDrIndication from "&_
											" VwSalesTaxDetails V,Acc_T_CreatedVoucherTaxDet T where V.ComputationMode is not null and  V.OUDefinitionID= "&sOrgId&" "&_
											" and V.InvoiceType=T.InvoiceType and V.TaxCategoryCode = T.TaxCategoryCode and V.TaxCode = T.TaxCode "&_
											" and T.CreatedTransNo = "&iVNo&" and T.TaxCode <> 0 and T.TaxCategoryCode <> 0 order by V.TaxHierarchy "
									'Response.Write sQry



									with rs4
										.CursorLocation = 3
										.CursorType = 3
										.Source = sQry
										.ActiveConnection = con
										.Open
									End With
									'IF Not rs4.Eof Then
										Do while  not rs4.Eof
											sTaxShtName = rs4(0)
											sCatCode = rs4(1)
											sTaxCode = rs4(2)
											sTaxMode = rs4(3)
											sTaxFormula = rs4(4)
											sTaxValue = rs4(8)
											iTaxAmount = rs4(7)
											sAccHead = rs4(5)
											sRoundOff = rs4(6)
											sTaxCrDrIndi = rs4("TransCrDrIndication")

											IF Cstr(sTaxCrDrIndi) = "D" Then
												iTaxAmount = Cdbl(iTaxAmount) * - 1
											End IF

											'Response.write iTaxAmount &"<br>"

											If trim(sAccHead) = "" then sAccHead = 0

												set Elem7 = oDOM.createElement("Tax")
												Elem7.setAttribute "CatCode",sCatCode
												Elem7.setAttribute "TaxCode",sTaxCode
												Elem7.setAttribute "TaxMode",sTaxMode
												Elem7.setAttribute "TaxFormula",sTaxFormula
												Elem7.setAttribute "TaxValue",sTaxValue
												If cint(sCatCode) = 0 and cint(sTaxCode) = 0 then
													Elem7.setAttribute "TaxAmount",iRoundOff
												Else
													Elem7.setAttribute "TaxAmount",iTaxAmount
												End If
												Elem7.setAttribute "AccHead",sAccHead
												Elem7.setAttribute "ItemValue","0"
												Elem7.setAttribute "RoundOff",sRoundOff
												Elem7.text = sTaxShtName
												TaxNode.appendchild Elem7
											rs4.MoveNext
										loop
										rs4.Close
									'Else
										Dim iRndAccHead
										sQry = "Select TaxAmount,isNull(AccountHead,0) from Acc_T_CreatedVoucherTaxDet where "&_
												" isNull(TaxCategoryCode,0) = 0 and isNull(TaxCode,0) = 0  and CreatedTransNo = "&iVNo&" "
										'		Response.Write sQry
										with rs3
											.CursorLocation = 3
											.CursorType = 3
											.Source = sQry
											.ActiveConnection = con
											.Open
										End With
										If not rs3.EOF then
											NewTaxAmt = rs3(0)
											iRndAccHead = rs3(1)
										Else
											NewTaxAmt = 0
											iRndAccHead = 0
										End If
										rs3.close
									'End IF


									If trim(iRndAccHead) = "" then iRndAccHead = 0

									set Elem8 = oDOM.createElement("Tax")
									Elem8.setAttribute "CatCode","0"
									Elem8.setAttribute "TaxCode","0"
									Elem8.setAttribute "TaxMode","0"
									Elem8.setAttribute "TaxFormula","0"
									Elem8.setAttribute "TaxValue","0"

									'Elem8.setAttribute "TaxAmount",NewTaxAmt
									Elem8.setAttribute "TaxAmount",iRoundOff
									Elem8.setAttribute "AccHead",iRndAccHead
									Elem8.text = "ROUND OFF"
									TaxNode.appendchild Elem8


					'===============================creation of Advance detail node ================
									set AdvNode = oDOM.createElement("AdvanceDetails")
									Root.appendchild AdvNode
					'The above Query for if any advance is been adjusted ===========================
									sQry = "SELECT A.CreatedAdvanceNo,V.CreatedVoucherNo, V.VoucherDate, V.VoucherAmount,P.AmountAdjusted, "&_
										" V.CreatedTransNo FROM Acc_T_CreatedVoucherHeader V INNER JOIN Acc_T_CreatedAdvances A "&_
										" ON V.CreatedTransNo = A.CreatedTransNo INNER JOIN ACC_T_CREATEDADVANCEADJ P ON "&_
										" A.CreatedAdvanceNo = P.CreatedAdvanceNo WHERE  P.CreatedTransNo = "&iVNo

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
											'aCTrNo = Advrs(6)

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

												set Elem8 = oDOM.createElement("Advance")
												Elem8.setAttribute "TransNo",aTrNo
												Elem8.setAttribute "VoucherNo",aVouNo
												Elem8.setAttribute "VoucherDate",aVouDate
												Elem8.setAttribute "AmountRec",aAmtRec
												Elem8.setAttribute "AmountAdj","0"
												Elem8.setAttribute "AmountToAdj",aAntToAdj
												Elem8.setAttribute "CreatedTransNo",dCTransNo
												Elem8.setAttribute "AdvNo",aAdvNo
												Elem8.setAttribute "AdjType","I"
												Elem8.setAttribute "ToAccount","0"
												AdvNode.appendchild Elem8

												rs3.MoveNext
											loop
											rs3.Close
										Advrs.MoveNext

									loop
									Advrs.Close
							'The above Query for if any advance is been adjusted ===========================
									sQry =	" SELECT P.TransactionNumber, P.AmountPayable, C.AmountPaid, C.CreatedTransNo,  "&_
											" C.CreatedTransNo, H.VoucherNumber, Convert(Varchar,H.VoucherDate,103) VoucherDate, P.PayablesNumber FROM Acc_T_CreatedPybleAdjDet C  "&_
											" INNER JOIN Acc_T_Payables P ON C.PayablesNumber = P.CRCreatedPayable INNER JOIN "&_
											" Acc_T_VoucherHeader H ON P.TransactionNumber = H.TransactionNumber "&_
											" WHERE C.CreatedTransNo = "&iVNo&" and H.BankInstrumentType = 'SC' "

									'Response.Write sQry
										with Advrs
											.CursorLocation = 3
											.CursorType = 3
											.Source = sQry
											.ActiveConnection = con
											.Open
										End with
										Do while not Advrs.EOF
											set Elem8 = oDOM.createElement("Advance")
											Elem8.setAttribute "TransNo",Advrs(0)
											Elem8.setAttribute "VoucherNo",Advrs("VoucherNumber")
											Elem8.setAttribute "VoucherDate",Advrs("VoucherDate")
											Elem8.setAttribute "AmountRec",Advrs("AmountPayable")
											Elem8.setAttribute "AmountAdj","0"
											Elem8.setAttribute "AmountToAdj",Advrs("AmountPaid")
											Elem8.setAttribute "CreatedTransNo",Advrs("PayablesNumber")
											Elem8.setAttribute "AdvNo",0
											Elem8.setAttribute "AdjType","SC"
											Elem8.setAttribute "ToAccount","0"
											AdvNode.appendchild Elem8
											Advrs.MoveNext
										loop
										Advrs.Close

							'The above Query for if any advance is been adjusted ===========================
									sQry =	" SELECT P.TransactionNumber, P.AmountPayable, C.AmountPaid, P.CreatedPayablesNumber,  "&_
											" C.CreatedTransNo, H.VoucherNumber, Convert(Varchar,H.VoucherDate,103) VoucherDate, P.PayablesNumber FROM Acc_T_CreatedPybleAdjDet C  "&_
											" INNER JOIN Acc_T_Payables P ON C.PayablesNumber = P.CRCreatedPayable INNER JOIN "&_
											" Acc_T_VoucherHeader H ON P.TransactionNumber = H.TransactionNumber "&_
											" WHERE C.CreatedTransNo = "&iVNo&" and H.BankInstrumentType = 'SR' "

									'Response.Write sQry
										with Advrs
											.CursorLocation = 3
											.CursorType = 3
											.Source = sQry
											.ActiveConnection = con
											.Open
										End with
										Do while not Advrs.EOF
											set Elem8 = oDOM.createElement("Advance")
											Elem8.setAttribute "TransNo",Advrs(0)
											Elem8.setAttribute "VoucherNo",Advrs("VoucherNumber")
											Elem8.setAttribute "VoucherDate",Advrs("VoucherDate")
											Elem8.setAttribute "AmountRec",Advrs("AmountPayable")
											Elem8.setAttribute "AmountAdj","0"
											Elem8.setAttribute "AmountToAdj",Advrs("AmountPaid")
											Elem8.setAttribute "CreatedTransNo",Advrs("PayablesNumber")
											Elem8.setAttribute "AdvNo",0
											Elem8.setAttribute "AdjType","SR"
											Elem8.setAttribute "ToAccount","0"
											AdvNode.appendchild Elem8
											Advrs.MoveNext
										loop
										Advrs.Close

					oDOM.save Server.MapPath("../temp/transaction/Sal_Voucher_xml_"&Session.SessionID&".xml")

		Response.Clear
		IF CStr(sDispTy) <> "S" Then
			Response.ContentType = "text/XML"
			Response.Write oDOM.xml
		Else
			Response.ContentType = "text/HTML"
			oDOM.save Server.MapPath("../temp/transaction/Sal_Voucher_xml_"&Session.SessionID&".xml")
			GetXmlForSal = "../temp/transaction/Sal_Voucher_xml_"&Session.SessionID&".xml"
		End IF

End Function

%>







