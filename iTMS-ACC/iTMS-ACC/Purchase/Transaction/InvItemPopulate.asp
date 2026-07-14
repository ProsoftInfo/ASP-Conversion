<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	InvItemPopulate.asp
	'Module Name				:	Purchase (Transaction)
	'Author Name				:	Ragavendran R
	'Created On					:	Dec 07,2010
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/purpopulate.asp"-->
<%

Dim oDOM,ndRoot,rsTemp,rsTemp1,rsItem,dcrs,rsAtt,rsPacking
Dim sSql,SubNode,sSuppInvNo,sSuppInvDt,SubNode1,NewElem1,ndPackDetails,ndPack
Dim sPartyName,Curr1,sPartyType,sPartySubType,Mod1,Mop,PayTerm
Dim IssueBank,Bop,Transporter,sPoNo,iEntryNo,iClassCode,dInvQty,dBalQty
Dim saSelRcptNo,sFlag,sSelRcptNo,sRefNum,sRefType,sOrgID,nChekOptVal
Dim iRcptNo,sRcptCode,sRcptDt,sConfNo,sBillType,sCash
Dim iParTypeID,sParTypeName,sParType,sTraUnit,sRecptAg,sCred,sTempRcptCode
Dim sActualReceiptNos,iPartyCode,iItemCode,iQtyRecd,sClassDesc,sItemDesc,UomCode
Dim sStockType,nActItemRate,sAttributeList,sOptName,sTemp,i,iOptVal,iPORefNo
Dim sItemDetExist,nItemRate,sQuery,nNoofSellForm,sInvNo,sAddDesc,iFreeQty,sOtherRefNo



Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set rsTemp = Server.CreateObject("ADODB.Recordset")
Set rsItem = Server.CreateObject("ADODB.Recordset")
Set rsTemp1 = Server.CreateObject("ADODB.Recordset")
Set dcrs = Server.CreateObject("ADODB.Recordset")
Set rsAtt = Server.CreateObject("ADODB.Recordset")
Set rsPacking = Server.CreateObject("ADODB.Recordset")

sRefType = Request.QueryString("RefType")
saSelRcptNo = Request.QueryString("hRecNo")
sOrgID = Request.QueryString("OrgID")
sInvNo = Request.QueryString("InvNo")
sOtherRefNo = Request.QueryString("OthRefNo")

if sFlag = "Multiple" then
	sarrSelRcptno = split(saSelRcptNo,",")
	'sSelRcptno = sarrSelRcptno(0)
	sSelRcptno = saSelRcptno
else
	sSelRcptno = saSelRcptno
end if
if trim(sSelRcptNo) = "" then sSelRcptNo = "0"
sRefNum = sSelRcptNo
if instr(1,sRefNum,",") > 0 then
	sRefNum = replace(sRefNum,",","','")
end if 

sRefNum = "'" & sRefNum & "'"



if trim(sRefType)="8" then ' Actual Receipt

	sSql = "Select isNull(OCNumber,0) From PUR_T_RefferenceNumberDet Where ReceiptNumber in (" & sSelRcptNo  &") "
	'Response.Write sSql
	With rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.Source =  sSql
		.ActiveConnection = con
		.Open
	End With
	Set rsTemp.ActiveConnection = nothing
	do while not rsTemp.EOF
		If trim(rsTemp(0)) <> "" Then
			sConfNo =  trim(sConfNo) & rsTemp(0) & "," 
		end If 'If trim(rsTemp(0)) <> "" Then
		rsTemp.movenext
	loop
	rsTemp.Close
	
	if trim(sConfNo) <> "" then
		sConfNo = mid(sConfNo,1,len(sConfNo) - 1 )
		IF trim(sConfNo) <> "0" then
			sConfNo = replace(sConfNo ,",","','")
			sConfNo = "'" &  sConfNo  & "'"
		end if 
	end if 
	
end if 'if trim(sRefType)="R" then
'Response.Write "saSelRcptNo="& saSelRcptNo
'Response.Write "sRefType = "&sRefType 
if trim(saSelRcptNo) <> "" then
	''to generate the receiptcode inorder to display the Rcpt code in a span
	'Response.Write sRefType = "& sRefType
	if trim(sRefType)="8" then 'Actual Receipt
	sSql ="SELECT ReceiptNumber,ReceiptCode,convert(varchar,ReceiptDate,103),isNull(BillType,0) FROM RCV_T_ActualReceiptHeader where GRNNumber in " &_
		" (Select GRNNumber from RCV_T_GateReceiptHeader where ReceivedForUnit='"& sOrgID& "') and ReceiptNumber in ("&saSelRcptNo&")"
	elseif trim(sRefType)="4" or trim(sRefType)="20" or trim(sRefType)="21" or trim(sRefType)="22" then ' Purchase order
		sSql = "Select PurchaseOrderNo,PurchaseOrderCode,Convert(varchar,PurchaseOrderDate,103),'' as BillType from PUR_T_POHeader where PurchaseOrderNo = "& saSelRcptNo 
		
	elseif trim(sRefType)="13" then
		sSql = "Select GatePassNo,DCCode,Convert(varchar,GeneratedOn,103),'' as BillType,TypeofItems from ForGatePassHeader where GatePassNo ="& saSelRcptNo 
	elseif trim(sRefType)="34" then
	    sSql = "Select WorkCompletionNo,isNull(WorkCompletionCode,WorkCompletionNo),Convert(varchar,WorkCompletionDate,103),'' as BillType from SER_T_ServiceWorkCompletion S, PUR_T_POHeader P "&_
	            " where S.AppRefNo = P.PurchaseOrderNo and S.WorkCompletionNo = "& saSelRcptNo &" and S.AppRefType = 4"
    ''Added By Ragav On Jan 09,2012	for Purchase Invoice Direct based on ActualReceipt Reference
    elseif Trim(sRefType)="23" then ' Sales Invoice Return
		sSql = "Select SalesReturnNo,isNull(InvoiceNumber,SaleTransactionNo),Convert(varchar,InvoiceDate,103),'' as BillType "&_
			   " from Sal_T_SalesReturnHeader H where SalesReturnNo ="& saSelRcptNo  
    elseif Trim(sRefType)="7" then ' Gate Receipt
        sSql = "Select GRNNumber,isNull(GRNCode,GRNNumber),Convert(varchar,GRNDate,103),"&_
               "'' as Bill Type from RCV_T_GateReceiptHeader where GRNNumber ="& saSelRcptNo 
	end if
	''end 
'	Response.Write "<p> sSql =" & sSql 

	With rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.Source =  sSql
		.ActiveConnection = con
		.Open
	End With
	Set rsTemp.ActiveConnection = nothing
	If not rsTemp.EOF then
		iRcptNo	= rsTemp(0)
		sRcptCode = rsTemp(1)
		sRcptDt  = rsTemp(2)
		sBillType  = rsTemp(3)
		
	IF trim(sBillType) = "C" then sCash = "Selected" else sCash = ""
	IF trim(sBillType) = "P" then sCred = "Selected" else sCred = ""
		sTempRcptCode = ""
		Do while not rsTemp.eof
			sActualReceiptNos = trim(sRcptCode)&"--"&trim(sRcptDt)
			rsTemp.MoveNext
			sTempRcptCode = sTempRcptCode + sActualReceiptNos + ","
		Loop
	End if
	rsTemp.Close
end if 'if trim(saSelRcptNo) <> "" then	

Set ndRoot = oDOM.createElement("Root")
oDOM.appendChild ndRoot

if trim(sRefType)="8" then
	sSql = "Select isNull(InvoiceNumber,'0'),InvoiceDate ,PartyCode,PartyType,ReceivedFrom,TRANSFEREDFROM,0 from RCV_T_GateReceiptHeader where ReceivedForUnit='"& sOrgID& "' and " &_
		" GRNNumber in (Select GRNNumber from RCV_T_ActualReceiptHeader where ReceiptNumber in ("& sRefNum & " ) )"

	'Response.Write "<p> sSql = "& sSql
	With rsItem
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	End With
	Set rsItem.ActiveConnection = nothing
	if not rsItem.EOF then
		sSuppInvNo = rsItem(0)
		sSuppInvDt = formatdate(rsItem(1))

		iPartyCode = rsItem(2)
		sPartyType = rsItem(3)
		sPartySubType = rsItem(4)
		sTraUnit = rsItem(5)
		sRecptAg = rsItem(6)
	end if
	rsItem.Close
elseif trim(sRefType)="34" then

    sSql = "Select isNull(PurchaseOrderCode,PurchaseOrderNo),Convert(varchar,PurchaseOrderDate,103),SupplierCode,PurchaseOrderNo"&_
           " from PUR_T_POHeader where PurchaseOrderNo in (Select AppRefNo from SER_T_ServiceWorkCompletion where "&_
           " AppRefType=4 and WorkCompletionNo = "& sRefNum &")"
	'Response.Write "<p> sSql = "& sSql
	With rsItem
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	End With
	Set rsItem.ActiveConnection = nothing
	if not rsItem.EOF then
		sSuppInvNo = rsItem(0)
		sSuppInvDt = formatdate(rsItem(1))
		iPartyCode = rsItem(2)
		iPORefNo = rsItem(3)
	end if
	rsItem.Close
end if ' if trim(sRefType)="R" then
'Response.Write "Party Code = "& iPartyCode 
if trim(iPartyCode)<>"" then
	if trim(sRefType)<>"8" then
		sSql = "Select PartyCode,PartyType,PartySubType,OUDefinitionID from vwOrgParty where PartyCode = "& iPartyCode 
		rsItem.open sSql,con
		if not rsitem.eof then
			iPartyCode = rsItem(0)
			sPartyType = rsItem(1)
			sPartySubType = rsItem(2)
			sTraUnit = rsItem(3)
		end if
		rsItem.close 
	end if
end if

if trim(iPartyCode) = "" then iPartyCode = 0

sSql = "Select PartyName from App_M_PartyMaster where PartyCode=" & trim(iPartyCode) & ""
With rsItem
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSql
	.ActiveConnection = con
	.Open
End With
Set rsItem.ActiveConnection = nothing
if not rsItem.EOF then
	sPartyName = rsItem(0)
End if
rsItem.Close

	Set SubNode= oDOM.createElement("InvoiceHeader")
	ndRoot.appendChild SubNode
	Set SubNode1=oDOM.createElement("ItemDetails")
	ndRoot.appendChild SubNode1
	
	Set NewElem1 = oDOM.createElement("Header")
	newElem1.setAttribute "OrgID", sOrgID
	newElem1.setAttribute "Party", sPartyName
	newElem1.setAttribute "PurchaseType", ""
	newElem1.setAttribute "Currency", Curr1
	newElem1.setAttribute "InvAgainst", "Receipt"
	newElem1.setAttribute "RefNum", sSelRcptNo
	newElem1.setAttribute "PartyCode", iPartyCode
	newElem1.setAttribute "PartyType",sPartyType
	newElem1.setAttribute "PartySubType",sPartySubType
	newElem1.setAttribute "CurrencyNo", ""
	newElem1.setAttribute "DespatchMode",Mod1
	newElem1.setAttribute "PaymentMode",Mop
	newElem1.setAttribute "PayTerms",PayTerm
	newElem1.setAttribute "IssueBank",IssueBank
	newElem1.setAttribute "BenificiaryBank",""
	newElem1.setAttribute "PricingBasis",Bop
	newElem1.setAttribute "Transporter",Transporter
	newElem1.setAttribute "LoadingPort",""
	newElem1.setAttribute "DestPort",""
	newElem1.setAttribute "Remarks",""
	newElem1.setAttribute "SuppInvNo",sSuppInvNo
	newElem1.setAttribute "SuppInvDt",sSuppInvDt
	newElem1.setAttribute "TransporterFlag",""
	newElem1.setAttribute "PoNo",sPoNo
	newElem1.setAttribute "ConfNum",sConfNo
	newelem1.setAttribute "InvoiceFlag",sFlag
	newelem1.setAttribute "InvValue",0
	newelem1.setAttribute "RoundOff",0

	
	newelem1.setAttribute "SuppCode",""
	newelem1.setAttribute "ItemType",""
	if trim(sInvNo)<>"" then
	    newelem1.setAttribute "InvoiceNumber",sInvNo
	end if
	SubNode.appendChild NewElem1

	'==============================================================================
	
	'Response.Write sRefNum 
'	Response.Write "sRefType= "& sRefType 
if (trim(sRefType)<>"" or not IsNull(sRefType)) and trim(sRefType)<>"N" then
	if trim(sRefType)="8" then
		'	sSql = "SELECT DISTINCT AR.ClassificationCode,AR.ItemCode,sum(AR.QuantityReceived)," &_
		'	" IC.GroupName,IM.ItemDescription,AR.EntryNo,AR.UoMCode,isNull(AR.StockType,''),isNull(AR.ItemRate,0),isNull(AR.ItemAttributes,'') FROM RCV_T_ActualRcptItemDet AR,INV_M_ITEMMASTER IM, " &_
		'	"INV_M_CLASSIFICATION IC WHERE AR.OrganisationCode='" & sOrgID & "' AND AR.ReceiptNumber in ( "& sRefNum & " ) " &_
		'	" AND IM.ClassificationCode = AR.ClassificationCode  AND IM.ITEMCODE = AR.ItemCode "&_
		'	" AND IM.ClassificationCode = IC.GroupCode  and IM.OrganisationCode='" & sOrgID & "'" & _
		'	" group by AR.ClassificationCode,AR.ItemCode,IC.GroupName,IM.ItemDescription,AR.EntryNo,AR.UoMCode,isNull(AR.StockType,''),isNull(AR.ItemRate,0),isnull(AR.ItemAttributes,'') "
		    sSql = "SELECT DISTINCT isnull(AR.ClassificationCode,0),isnull(AR.ItemCode,0),sum(AR.QuantityReceived),AR.EntryNo,"&_
		           " AR.UoMCode,isNull(AR.StockType,''),isNull(AR.ItemRate,0),isNull(AR.ItemAttributes,''),isNull(AR.AdditionalDescr,''),isNull(AR.FreeQuantity,0)"&_
		           " FROM RCV_T_ActualRcptItemDet AR where AR.OrganisationCode='" & sOrgID & "' AND "&_
		           " AR.ReceiptNumber in ( "& sRefNum & " ) Group By AR.ClassificationCode,AR.ItemCode,AR.EntryNo,AR.UOMCOde,AR.StockType,AR.ITemRate,AR.ITEMATTRIBUTES,AR.AdditionalDescr,isNull(AR.FreeQuantity,0)"
	elseif trim(sRefType)="13" then
	'	sSql = "Select D.ClassificationCode,D.ItemCode,Quantity,C.GroupName,I.ItemDescription,EntryNo,InvoicedUoM,'S',"&_
	'		   "(ItemValue/Quantity),isNull(ItemAttributes,'') from ForGatePassDetails D,INV_M_ITEMMASTER I,INV_M_CLASSIFICATION C where I.ItemCode = D.ItemCode"&_
	'		   " and D.ClassificationCode = C.GroupCode and GatePassNo in ("& sRefNum &") and I.OrganisationCode = '"&sOrgID&"'"
	
	    sSql = "Select isnull(ClassificationCode,0),isnull(ItemCode,0),Quantity,EntryNo,InvoicedUoM,'S',"&_
               " (ItemValue/Quantity),isNull(ItemAttributes,''),isNull(Description,''),0 from ForGatePassDetails "&_
               " Where GatePassNo in ("& sRefNum &") "
	elseif trim(sRefType)="4" or trim(sRefType)="20" or trim(sRefType)="21" or trim(sRefType)="22" then
'		sSql = "Select D.ClassificationCode,D.ItemCode,QuantityOrdered,C.GroupName,I.ItemDescription,"&_
'			   " EntryNumber,UnitOfMeasure,'S',RatePerUnit,isNull(ItemAttributes,'') from PUR_T_PODetails D,"&_
'			   " INV_M_ITEMMASTER I,INV_M_CLASSIFICATION C where I.ItemCode = D.ItemCode "&_
'			   " and D.ClassificationCode = C.GroupCode and PurchaseOrderNo in ("& sRefNum &") and D.OrganisationCode = '"& sOrgID & "'"

		sSql = "Select isNull(ClassificationCode,0),isNull(ItemCode,0),QuantityOrdered,EntryNumber,UnitOfMeasure,'S',"&_
               " RatePerUnit,isNull(ItemAttributes,''),isNull(AdditionalDescr,''),0 from PUR_T_PODetails where PurchaseOrderNo in ("& sRefNum &") and OrganisationCode = '"& sOrgID &"'"
               
               if trim(sOtherRefNo)<>"" then
                    sSql = "Select IsNull(ClassificationCode,0),IsNull(ItemCode,0),ReleasedQty,EntryNumber,UnitOfMeasure,'S',ReleasedRate,"&_
                           " IsNull(ItemAttributes,''),IsNull(RD.AdditionalDescr,'N/A'),0 from PUR_T_PODetails D,PUR_T_POReleaseDetail RD "&_
                           " where PurchaseOrderNo = "& sRefNum &" and ReleaseEntryNo = "& sOtherRefNo &" and D.EntryNumber = RD.EntryNo and OrganisationCode ='"& sOrgID &"'"
               end if
    elseif trim(sRefType)="34" then
        sSql = "Select isNull(ClassificationCode,0),isNull(ItemCode,0),QuantityOrdered,EntryNumber,UnitOfMeasure,'S',"&_
               " RatePerUnit,isNull(ItemAttributes,''),isNull(AdditionalDescr,''),0 from PUR_T_PODetails where PurchaseOrderNo in ("& iPORefNo &") and OrganisationCode = '"& sOrgID &"'"
	elseIf trim(sRefType) = "23" then ' Sales Invoice Return
		sSql = " Select isNull(ClassificationCode,0),isNull(ItemCode,0),QuantityForReturn,0,D.InvoicedUOM,'S',"&_
		       " D.InvoicedRate,isNull(ItemAttributes,''),(Select ItemDescription from VWITEM where ItemCode = D.ItemCode and ClassificationCode=D.ClassificationCode),0 from Sal_T_SalesReturnDetail D, "&_
			   " Sal_T_SalesReturnHeader H where D.SalesReturnNo=H.SalesReturnNo and H.SalesReturnNo =  "& sRefNum 
    elseif trim(sRefType)="7" then 'gate receipt
        sSql = "Select isNull(ClassificationCode,0),isNull(Itemcode,0),QuantityReceived,'0',QuantityUOM,'S',"&_
               " 0,isNull(ItemAttributes,''),isNull(AdditionalDescr,'N/A'),0 "&_
               "  from RCV_T_GRNItemDetails where GRNNumber = "& sRefNum 
    end if 
'	Response.Write ssql
		With rsItem
			.CursorLocation = 3
			.CursorType = 3
			.Source = sSql
			.ActiveConnection = con
			.Open
		End With
		Set rsItem.ActiveConnection = nothing
		iEntryNo = 0 
		If not rsItem.EOF then
		Do While Not rsItem.EOF		
			 iClassCode = rsItem(0)
			 iItemCode= rsItem(1)
			 iQtyRecd = rsItem(2)
			 nChekOptVal = "0" 
			 
			 sSql = "Select ItemDescription from VWITEM where ItemCode = "& iItemCode &" and ClassificationCode = "& iClassCode 
			 rsTemp.Open ssql,con
			 if not rsTemp.EOF then
			    sItemDesc = rsTemp(0)
			 end if
			 rsTemp.Close 
			 
			 if iItemCode = "0" and iClassCode = "0" then
				sItemDesc = rsItem(8)
			 end if
			 
			 sSql = "Select GroupName from INV_M_CLassification where GroupCode = "& iClassCode
			 rsTemp.Open sSql,con
			 if not rsTemp.EOF then
			    sClassDesc = rsTemp(0)
			 end if
			 rsTemp.Close 
			 
			 
			' sClassDesc = rsItem(3)
			 'sItemDesc = rsItem(4)
			' iEntryNo = rsItem(5)
			 UomCode=rsItem(4)
			 sStockType = rsItem(5)
			 nActItemRate = rsItem(6)
			sAttributeList = rsItem(7)
			sAddDesc = rsItem(8)
			iFreeQty = rsItem(9)
			
			if trim(sRefType)="8" then
				sItemDesc = sAddDesc
			else		
				if trim(sItemDesc)<>"" and trim(sAddDesc)<>"" then 
				    sItemDesc = sItemDesc &" - "& sAddDesc
				elseif trim(sItemDesc)="" and trim(sAddDesc)<>"" then
				    sItemDesc = sAddDesc
				end if
			end if ' if trim(sRefType)="8" then
			
			'Response.Write "sAttributeList="&sAttributeList
			If trim(sAttributeList) <> "" then
				sOptName = ""
				sTemp = split(sAttributeList,",")
				For i = 0 to UBOUND(sTemp) 
					iOptVal = split(sTemp(i),"#")
					'Response.Write sTemp(i)
					If sTemp(i)= "NULL" Then 
						sOptName = ""
					Else
						if UBound(iOptVal)=1 then
							nChekOptVal = split(iOptVal(1),":")(0)
					
							If nChekOptVal <> "0" Then
									'sSql = "Select ItemTypeAttributeID,ItemTypeAttributeName from Inv_M_ItemTypeAttributes where ItemTypeID = '"& sItemType &"'"
									sSql = "Select ItemTypeAttributeID,ItemTypeAttributeName from Inv_M_ItemTypeAttributes"
									'Response.Write sSql
									rsTemp.open sSql,con
									if not rsTemp.eof then
										sSql = "Select OptionName from Inv_M_ItemTypeOptions where OptionValue = "&iOptVal(1)&" and ItemTypeAttributeID ="& rsTemp(0)& " "
										Response.Write "<p>sSql="&sSql
										rsAtt.Open sSql,con
										If not rsAtt.EOF then
											sOptName =sOptName &","& rsAtt(0)
										else
											sOptName =sOptName &","& rsTemp(1)
										End If
										rsAtt.Close 
									end if
									rstemp.close
							Else
								sOptName = ""
							End IF'If nChekOptVal <> "0" Then

						else
								'sSql = "Select ItemTypeAttributeID,ItemTypeAttributeName from Inv_M_ItemTypeAttributes where ItemTypeID = '"& sItemType &"'"
								sSql = "Select ItemTypeAttributeID,ItemTypeAttributeName from Inv_M_ItemTypeAttributes "
								'Response.Write sSql
								rsTemp.open sSql,con
								if not rsTemp.eof then
									sSql = "Select OptionName from Inv_M_ItemTypeOptions where OptionValue = "&iOptVal(0)&" and ItemTypeAttributeID ="& rsTemp(0)& " "
									rsAtt.Open sSql,con
									If not rsAtt.EOF then
										sOptName =sOptName &","& rsAtt(0)
									else
										sOptName =sOptName &","& rsTemp(1)
									End If
									rsAtt.Close 
								end if
								rsTemp.close
							
						end if 'if UBound(iOptVal)=1 then
					End IF	 'If sTemp(i)= "NULL" Then  
				Next
			End If
			IF sOptName <> "" then 
				sOptName = " [" & mid(sOptName,2) &"] "
			End IF	
			iEntryNo = iEntryNo + 1
			'Response.Write "<p> sItemDesc="&sItemDesc &"<BR><BR>"
			'Response.Write "<p> sRefNum = "& sRefNum 
			'' to add Qty Validation for Receipt
			sSql = "Select Sum(B.InvoiceQuantity) from Rcv_T_InvoiceHeader A, Rcv_T_InvoiceDetails B " &_
					" Where isnull(ItemCode,0)=" & trim(iItemCode) & "  and isnull(ClassificationCode,0)= " & trim(iClassCode) & " and " &_
					" OrganisationCode ='" & trim(sOrgID) & "'  and A.InvoiceAgainst = 1 and A.ReferenceNumber in (" & sRefNum & ")  and " &_
					" B.InvoiceNumber=A.InvoiceNumber " &_
					" group by Itemcode,ClassificationCode,OrganisationCode "
			'	Response.Write "<p> sSql = "& sSql 
			
			With dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sSql
				.ActiveConnection = con
				.Open
			End With
			Set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				dInvQty = dcrs(0)
			else
				dInvQty = 0
			end if
			dcrs.Close
			
			'Response.Write "<p> iQtyRecd = " & trim(iQtyRecd)
			'Response.Write "<p>	dInvQty = " & trim(dInvQty)
			
			dBalQty = round(cdbl(iQtyRecd) - cdbl(dInvQty),3)
			
		'	Response.Write "<p> dBalQty = "& trim(dBalQty)

			if cdbl(dBalQty) > 0 then
			
				sItemDetExist = "N"
				if SubNode1.hasChildNodes then
					for each newElem1 in SubNode1.ChildNodes
						if trim(newElem1.getAttribute("EntryNo")) = trim(iEntryNo) and trim(newElem1.getAttribute("ItemCode")) = trim(iItemCode) and trim(newElem1.getAttribute("ClassificationCode")) = trim(iClassCode) then
							sItemDetExist = "Y"
							exit for
						end if 
					next
				end if 'if SubNode1.hasChildNodes then

				if sItemDetExist = "N" then
					if trim(sPoNo) <> "" then
						if instr(1,sPoNo,",") > 0 then
							nItemRate = GetPOItemRate(mid(sPoNo,1,instr(1,sPoNo,",")-1),iItemCode,iClassCode,iEntryNo,sOrgID)
						else
							nItemRate = GetPOItemRate(sPoNo,iItemCode,iClassCode,iEntryNo,sOrgID)
						end if 	
					else
						nItemRate = nActItemRate
					end if 	
					
					if Trim(nItemRate)="" or IsNull(nItemRate) then nItemRate = 0
					
					'Response.Write "nItemRate = "& nItemRate
					
					nItemRate = FormatNumber(nItemRate,5,,,0)
					
					if Trim(sRefType)<>"8" then
						If trim(sAttributeList) <> "" then sItemDesc = sItemDesc & sOptName
					end if 
					Set newElem1 = oDOM.createElement("Item")
					
					newElem1.setAttribute "ItemCode", iItemCode
					newElem1.setAttribute "ClassificationCode", iClassCode
					newElem1.setAttribute "ItmDescription", sItemDesc
					newElem1.setAttribute "Uom", UomCode
					newElem1.setAttribute "Qty", dBalQty
					newElem1.setAttribute "Rate", nItemRate
					newElem1.setAttribute "DisPer", "0"
					newElem1.setAttribute "DisAmount", "0"
					newElem1.setAttribute "NettBasic", "0"
					newElem1.setAttribute "UomDesc", UomCode
					newElem1.setAttribute "EntryNo", iEntryNo
					newElem1.setAttribute "RatePerQtyUoM","0"
					newElem1.setAttribute "SourceEntryNo", iEntryNo
					newElem1.setAttribute "PurchaseType", ""
					newElem1.setAttribute "Amount", "0"
					newElem1.setAttribute "ItemValue", "0"
					newElem1.setAttribute "ItemRate", "0"
					newElem1.setAttribute "RateUOM", ""
					newElem1.setAttribute "StockType", sStockType
					newElem1.setAttribute "VAT", ""
					newElem1.setAttribute "AttributeList",sAttributeList
					newElem1.setAttribute "FreeQty",iFreeQty
					SubNode1.appendChild NewElem1
				else
					newElem1.setAttribute "Qty", CDbl(newElem1.getAttribute("Qty")) + CDbl(dBalQty)	
					newElem1.setAttribute "StockType",trim(newElem1.getAttribute("StockType")) + "," + trim(sStockType)
					newElem1.setAttribute "FreeQty",iFreeQty
				end if 'if sItemDetExist = "N" then
				
				
				sQuery = "Select PackingCode,isNull(MillPackingNumber,0),MillGrossWeight,MillNettWeight,"&_
					"isNull(WeightPerSellingForm,0),isNull(MillLotNo,'N/A'),MillSerialNo,isNull(PackingForm,0) from RCV_T_ActualRcptLotSerial "&_
					" where ReceiptNumber in ("& sSelRcptNo&")"
					'Response.Write sQuery
				rsPacking.Open sQuery,con
				if not rsPacking.EOF then
					set ndPackDetails = oDOM.createElement("PackingDetails")
					newElem1.appendChild ndPackDetails 
					do while not rsPacking.EOF
							set ndPack = oDOM.createElement("Pack")
								ndPack.setAttribute "Packcode",trim(rsPacking(0))
								ndPack.setAttribute "PackNumber",trim(rsPacking(1))
								ndPack.setAttribute "PackGrossQty",trim(rsPacking(2))
								ndPack.setAttribute "PackNettqty",trim(rsPacking(3))
								
								if Cdbl(rsPacking(4))>0 then
									nNoofSellForm = cint(cdbl(rsPacking(3))/cdbl(rsPacking(4)))
								else
									nNoofSellForm = cint(rsPacking(3))
								end if ' if Cdbl(rsPacking(4))>0 then
								
								ndPack.setAttribute "NoofSellingForm", nNoofSellForm
								ndPack.setAttribute "Selected","Y"
								ndPack.setAttribute "WeightperSellingform",Trim(rsPacking(4))
								ndPack.setAttribute "Sellingnumber",trim(rsPacking(7))
							ndPackDetails.appendChild ndPack 
						rsPacking.MoveNext 
					loop
				end if
				rsPacking.Close 
				
			End if 'if cdbl(dBalQty) > 0 then
			rsItem.MoveNext
		Loop
		End if
		rsItem.Close
		'Root.appendChild SubNode1
		'''''To get New Items''''''''''
		sSql = "SELECT A.TempItemCode,Sum(A.QuantityReceived),B.ItemDescription,'',A.UoMCode,isNull(A.StockType,''),isNull(A.ItemRate,0) FROM RCV_T_ActualRcptItemDet A,MS_TemporaryItemMaster B WHERE " &_
		" A.OrganisationCode='" & trim(sOrgID) & "' AND A.ReceiptNumber in ( "& sRefNum& ")  and A.TempItemCode = B.TempItemCode" &_
		" and A.TempItemCode is not null and A.TempItemCode <> '0' group by A.TempItemCode,B.ItemDescription,A.UoMCode,isNull(A.StockType,''),isNull(A.ItemRate,0)"
		'Response.Write ssql
		With rsItem
			.CursorLocation = 3
			.CursorType = 3
			.Source =  sSql
			.ActiveConnection = con
			.Open
		End With
		Set rsItem.ActiveConnection = nothing
		iClassCode = "TEMP"
		If not rsItem.EOF then
		Do While Not rsItem.EOF
			Set iItemCode = rsItem(0)
			Set iQtyRecd  = rsItem(1)
			Set sItemDesc = rsItem(2)
			Set UomCode = rsItem(4)
			set sStockType = rsItem(5)
			set nActItemRate = rsItem(6)
		
			'' To add Qty Validation for Receipt
			sSql = "Select Sum(B.InvoiceQuantity) from Rcv_T_InvoiceHeader A, Rcv_T_InvoiceDetails B Where " &_
				   "isnull(TempItemCode,0)=" & trim(iItemCode) & " and A.InvoiceAgainst = 1 and A.ReferenceNumber" &_
				   "in (" & trim(sRefNum) & " ) and B.InvoiceNumber=A.InvoiceNumber group by TempItemCode "
			'	Response.Write sSql
			With dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sSql
				.ActiveConnection = con
				.Open
			End With
			Set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				dInvQty = dcrs(0)
			else
				dInvQty = 0
			end if
			dcrs.Close
			dBalQty = cdbl(iQtyRecd) - cdbl(dInvQty)
			
			'Response.Write "<p> dBalQty = "& trim(dBalQty)
			if cdbl(dBalQty) > 0 then
			
			
				sItemDetExist = "N"
				if SubNode1.hasChildNodes then
					for each newElem1 in SubNode1.ChildNodes
						if trim(newElem1.getAttribute("EntryNo")) = trim(iEntryNo) and trim(newElem1.getAttribute("ItemCode")) = trim(iItemCode) and trim(newElem1.getAttribute("ClassificationCode")) = trim(iClassCode) then
							sItemDetExist = "Y"
							exit for
						end if 
					next
				end if 'if SubNode1.hasChildNodes then

				if sItemDetExist = "N" then
		
					if trim(sPoNo) <> "" then
						if instr(1,sPoNo,",") > 0 then
							nItemRate = GetPOItemRate(mid(sPoNo,1,instr(1,sPoNo,",")-1),iItemCode,iClassCode,iEntryNo,sOrgID)
						else
							nItemRate = GetPOItemRate(sPoNo,iItemCode,iClassCode,iEntryNo,sOrgID)
						end if 	
					else
						nItemRate = nActItemRate
					end if 	
					
					nItemRate = FormatNumber(nItemRate,5,,,0)
					
				
					iEntryNo = iEntryNo + 1
				
					Set newElem1 = oDOM.createElement("Item")
					newElem1.setAttribute "ItemCode", iItemSode
					newElem1.setAttribute "ClassificationCode", iClassCode
					newElem1.setAttribute "ItmDescription", sItemDesc
					newElem1.setAttribute "Uom", UomCode
					newElem1.setAttribute "Qty", dBalQty
					newElem1.setAttribute "Rate", nItemRate
					newElem1.setAttribute "DisPer", "0"
					newElem1.setAttribute "DisAmount", "0"
					newElem1.setAttribute "NettBasic", "0"
					newElem1.setAttribute "UomDesc", UomCode
					newElem1.setAttribute "EntryNo", iEntryNo
					newElem1.setAttribute "RatePerQtyUoM", "0"
					newElem1.setAttribute "SourceEntryNo", iEntryNo
					newElem1.setAttribute "PurchaseType", ""
					newElem1.setAttribute "Amount", "0"
					newElem1.setAttribute "ItemValue", "0"
					newElem1.setAttribute "ItemRate", "0"
					newElem1.setAttribute "RateUOM", ""
					newElem1.setAttribute "StockType", trim(sStockType)
					newElem1.setAttribute "VAT", ""
					SubNode1.appendChild newElem1
				else
					newElem1.setAttribute "Qty", CDbl(newElem1.getAttribute("Qty")) + CDbl(dBalQty)	
					newElem1.setAttribute "StockType",trim(newElem1.getAttribute("StockType")) + "," + trim(sStockType)
				end if 'if sItemDetExist = "N" then
					
			end if 'if cdbl(dBalQty) > 0 then
			
			rsItem.MoveNext
		Loop
		End if
		rsItem.Close
		
	end if ' if trim(sRefType)<>"" or not IsNull(sRefType) then
	
	Response.ContentType = "text/xml"
	Response.Write oDOM.xml
%>
