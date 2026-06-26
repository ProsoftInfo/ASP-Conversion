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
	'This Function will call the Sales Invoice XML Creation
	Function GetXmlForSalesReturn(iVNo)
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
		Dim sAgTy,sAgSubTy,seRndVal,iTotalTaxAmount,iTaxPercentage,iCrTransNoForSalesInvoice
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
		'sQry = " Select A.CreatedTransNo,A.CreatedVoucherNo,isNull(V.TransactionNumber,0),isNull(V.VoucherNumber,'') from Acc_T_VoucherHeader V,Acc_T_CreatedVoucherHeader A where A.CreatedTransno = "& iVNo &" "&_
		'   " and A.CreatedTransno *= V.CreatedTransno "
        sQry = " Select SalesReturnNo,I.InvoiceNumber,R.SaleTransactionNo,I.InvoiceNumber,R.TypeOfSale,R.TypeOfInvoice from Sal_T_SalesReturnHeader R,Sal_T_InvoiceHeader I where R.SaleTransactionNo = I.SaleTransactionNo and SalesReturnNo = "& iVNo 
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
			sSalType = rs1(4)
		End IF
		rs1.Close
		'Response.Write "sVDate =" & sVDate & "<BR><BR>"
		sOrgId = Session("organizationcode")
		sName = Session("OrgShortName")
		
	sQry = "Select CreatedTransNo from Acc_T_CreatedVoucherHeader where FromApplication=3 and  OtherApplnTransNo = "& sTransNo 
	rs1.Open sQry,con
	if not rs1.EOF then
	    iCrTransNoForSalesInvoice = rs1(0)
	end if
	rs1.Close 

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

				set Elem1 = oDOM.createElement("Organization")
				Elem1.setAttribute "OrgId",sOrgId
				Elem1.setAttribute "AccUnit",""
				Elem1.text = sName
				Hnode.appendchild Elem1

				        set Elem2 = oDOM.createElement("Book")
				            Elem2.setAttribute "BookId","0"
				            Elem2.setAttribute "BKAccHead","0"
				            Elem2.setAttribute "BKOtherUnits","1"
				            Elem2.text = sBkName
				            Hnode.appendchild Elem2
					
		
				sQry = " Select InvoiceTypeName from Sal_M_InvoiceTypes where InvoiceType ="& sSalType 

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

					sQry = " select V.PartyType,V.PartySubType,V.PartyCode,V.SubTypeName from Sal_T_SalesReturnHeader A, "&_
						   " VwOrgParty V where A.SalesReturnNo = "&iVNo&" and  V.OUDefinitionID = A.InvoicedForUnit and "&_
						   " V.PartySubType = A.PartySubType and V.PartyType = A.PartyType and V.PartyCode = A.PartyCode "
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
						.Source = " Select isNull(SUM(BasicAmount),0) from Sal_T_SalesReturnDetail where SalesReturnNo = "& iVNo
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
						.Source = "Select Convert(varchar,InvoiceDate,103) from Sal_T_SalesReturnHeader where SalesReturnNo = "&iVNo&" "
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
					'	sQry = " Select VoucherEntryNumber,isNull(ItemDescription,0),Amount,InvoicedQuantity,isNull(InvoicedUOM,''),InvoicedRate,BasicAmount, "&_
					'			" DiscountPercent,DiscountAmount,ItemCode,ClassificationCode,RatePer,isNull(NoOfPack,0),isNull(PackCode,0),isNull(RoundOffValue,0),TransCrDrIndication from "&_
					'			" Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iVNo
					
					    sQry = "Select 0,isNull(ItemDescription,0),BasicAmount,InvoicedQuantity,isNull(InvoicedUom,''),InvoicedRate,BasicAmount, "&_
					           " DiscountPercent,DiscountAmount,D.ItemCode,D.ClassificationCode,RatePer,0,0,"&_
					           " 0,'D' from Sal_T_SalesReturnDetail D,VWItem V where D.ItemCode = V.ItemCode "&_
					           " and D.ClassificationCode = V.ClassificationCode and SalesReturnNo = " & iVNo
								
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
								.Source =" Select isNull(AccUnitAccountHead,0)from Acc_T_CreatedVoucherDetails where CreatedtransNo = "&iCrTransNoForSalesInvoice &" and ItemCode = "& eItemCode
								.ActiveConnection = con
								.Open
							End with
							Do while not rs5.EOF
								sAccHeadNo = rs5(0)


								With rs4
									.CursorLocation = 3
									.CursorType = 3
									.Source = "select Count(1) from Acc_T_CreatedVoucherCCDet where CreatedTransNo = "&iCrTransNoForSalesInvoice&" "&_
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
									.Source = "Select Count(1) from Acc_T_CretedVoucherAHDet where CreatedTransNo =	 "&iCrTransNoForSalesInvoice&" "&_
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
							'.Source = " Select VoucherAmount from Acc_T_CreatedVoucherHeader where  CreatedTransNo = "&iVNo&" "
							.Source = " Select isNull(SUM(BasicAmount),0) from Sal_T_SalesReturnDetail where SalesReturnNo = "& iVNo 
							.ActiveConnection = con
							.Open
						End With
						If not rs3.EOF	then
							tVouAmt = rs3(0)
						End If
						rs3.Close
						Dim sTaxCrDr
						
								Set TaxNode = oDOM.createElement("TaxDetails")
								TaxNode.setAttribute "InvoiceVlaue",tVouAmt
								TaxNode.setAttribute "BasicValue",sBasVal
								TaxNode.setAttribute "NettValue",tVouAmt
								TaxNode.setAttribute "RoundOffValue",iRoundOff
								Root.appendchild TaxNode

								'==========================Tax node============================
									Dim iTaxAmount,sTaxCrDrIndi
								'	sQry = "select V.TaxShortName,V.TaxCategoryCode,V.TaxCode,V.ComputationMode,isnull(V.SumOfFields,'') Formula, "&_
								'			" V.AccountHead,isNull(V.RoundOff,0) RoundOff ,T.TaxAMount,isnull(V.FlatAmount,0) TaxValue,TransCrDrIndication from "&_
								'			" VwSalesTaxDetails V,Acc_T_CreatedVoucherTaxDet T where V.ComputationMode is not null and  V.OUDefinitionID= "&sOrgId&" "&_
								'			" and V.InvoiceType=T.InvoiceType and V.TaxCategoryCode = T.TaxCategoryCode and V.TaxCode = T.TaxCode "&_
								'			" and T.CreatedTransNo = "&iVNo&" and T.TaxCode <> 0 and T.TaxCategoryCode <> 0 order by V.TaxHierarchy "
								'	'Response.Write sQry
								
								    sQry = "Select V.TaxShortName,V.TaxCategoryCode,V.TaxCode,V.ComputationMode,isnull(V.SumOfFields,'') Formula, "&_
								           "  V.AccountHead,isNull(V.RoundOff,0) RoundOff,T.TaxAMount,isnull(V.FlatAmount,0) TaxValue,T.TaxPercentage from "&_
								           "  VwSalesTaxDetails V,Sal_T_InvoiceTaxDetails T where V.ComputationMode is not null and  V.OUDefinitionID= "& sOrgId &""&_
								           " and V.InvoiceType=T.InvoiceType and V.TaxCategoryCode = T.TaxCategoryCode and V.TaxCode = T.TaxCode "&_
								           " and T.SaleTransactionNo = "& iVNo &" and T.TaxCode <> 0 and T.TaxCategoryCode <> 0 order by V.TaxHierarchy "

'Response.Write "<textarea>"& sQry&"</textarea>"
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
											iTaxPercentage = rs4(9)
											
											'sTaxCrDrIndi = rs4("TransCrDrIndication")

										'	IF Cstr(sTaxCrDrIndi) = "C" Then
										'		iTaxAmount = Cdbl(iTaxAmount) * - 1
										'	End IF
										
										    if sTaxMode = "P" then
										        if cdbl(iTaxPercentage) > 0 then
										            iTaxAmount = cdbl((cdbl(sBasVal) *cdbl(iTaxPercentage))/100)
										            sTaxValue = iTaxPercentage 
										        end if
										    elseif sTaxMode = "F" then
										        iTaxAmount = sTaxValue 
										    end if
										    iTaxAmount = FormatNumber(iTaxAmount,2,0,0,0)
										    Response.Write "<p> Basic Value = "& sBasVal
										    Response.Write "<p> TaxMode =  "&  sTaxMode 
										    Response.Write "<p> Tax Value = "& iTaxPercentage 
										    Response.Write "<p> TaxAmount = "& iTaxAmount 
										    'Response.End 
										    iTotalTaxAmount = cdbl(iTotalTaxAmount) + cdbl(iTaxAmount)
										    
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
										sBasVal = CDbl(sBasVal)+CDbl(iTotalTaxAmount)
										iRoundOff = CDbl(sBasVal) - FormatNumber(sBasVal,0,0,0,0)
										
										iRoundOff = FormatNumber(iRoundOff,2,0,0,0)
										
										Dim iRndAccHead
									'	sQry = "Select TaxAmount,isNull(AccountHead,0) from Acc_T_CreatedVoucherTaxDet where "&_
									'			" isNull(TaxCategoryCode,0) = 0 and isNull(TaxCode,0) = 0  and CreatedTransNo = "&iVNo&" "
									
									sQry = "Select T.TaxAMount,V.AccountHead from VwSalesTaxDetails V,Sal_T_InvoiceTaxDetails T where "&_
									       " V.ComputationMode is not null and  V.OUDefinitionID= "& sOrgId &" and V.InvoiceType=T.InvoiceType and "&_
									       " V.TaxCategoryCode = T.TaxCategoryCode and V.TaxCode = T.TaxCode  and T.SaleTransactionNo ="& iVNo
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


					
		Response.Clear
		Response.ContentType = "text/HTML"
		oDOM.save Server.MapPath("../temp/transaction/Sal_Return_Voucher_xml_"&Session.SessionID&".xml")
		GetXmlForSalesReturn = "../temp/transaction/Sal_Return_Voucher_xml_"&Session.SessionID&".xml"
End Function

%>







