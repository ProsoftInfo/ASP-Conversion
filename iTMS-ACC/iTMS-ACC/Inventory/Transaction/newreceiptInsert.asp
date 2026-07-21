<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	newreceiptInsert.asp
	'Module Name				:	Inventory (Internal Receipt Accounting)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 21, 2003
	'Modified By				:	KUMAR K A
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	receiptEntry.asp
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
<!--#include file="../../include/populate.asp"-->
<%
'XML DOM Variables
Dim oDOM,Root,objfs

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

dim newxml,RootNode,HeaderNode,PageNode,SchNode,QtyNode,HeaderNodeO,LotNode,adoCmd

Set newxml = Server.CreateObject("Microsoft.XMLDOM")

newxml.async = false
newxml.load(Request)

newxml.save Server.MapPath("../temp/transaction/RR.xml")
'response.end

dim dcrs,dcrs1,dcrs2,dcrs3,sSql
dim iItemCode,iClass,arrStore,sBin,sLoc,sOrgCode
dim iRecNo,iInvRecNo,iQtyOrd,iQtyRec,iQtyIns,iQtyAcc,iQtyRej,iItmRate,iValue
dim sPressFrom,sPressTo,sQtyIn,iTareQty,sPress,sCrop,iBales,iLot,iSerialFrom,iSerialTo,iSerial
dim arrLotSerial,iLotEntry,iSerialEntry,iQtyRecEntry,iTareQtyEntry,sTareIn,iQtyGross,iQtyNett
dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin,sSellingType,iWeight,sPackingType
dim iOrLot,iIssNo,iMRSNo,sSrcType,iSrcNo,sPackingNum,sTempLot,sSellingFormType
dim iProcessNo,iOrderConfirmationNo,sStage,iAccountedBy,sReceivedOn
dim iPieceNo,iPieceQty,iLotValue,iRate,iDINo,sOrType
dim sIssueEntryNo,sPRODUCTIONORDERNO,sSrcCode
dim iNoofPacks, PackNode, sExp

iProcessNo = "NULL"
iOrderConfirmationNo = "NULL"
sOrType = ""

iAccountedBy = getUserid

Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
Set dcrs3 = Server.CreateObject("ADODB.RecordSet")



con.beginTrans

	Set RootNode = newxml.documentElement
	iRecNo = trim(RootNode.Attributes.Item(0).nodeValue)
	''''''''''''''''''''''''''''''''''''''' only for Finished Goods '''''''''''''''''

	sTempLot = trim(RootNode.Attributes.Item(1).nodeValue)

	''''''''''''''''''''''''''''''''''''''' only for Finished Goods '''''''''''''''''

	sReceivedOn = trim(RootNode.Attributes.Item(2).nodeValue)
	
	if len(mid(sReceivedOn,4,2)) = 1 then
	    sTempMonYr = "0"&mid(sReceivedOn,4,2)
    else
	    sTempMonYr = mid(sReceivedOn,4,2)
    end if
    sMonYr = sTempMonYr&Year(sReceivedOn)

    arrFin = split(GetFinancialYear(sMonYr),":")
    sFinFrom = arrFin(0)
    sFinTo = arrFin(1)
	

	sPRODUCTIONORDERNO = ""
	with dcrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT PRODUCTIONORDERNO FROM PRD_T_DAILYPACKINGHEADER WHERE DAILYPACKINGCODE = (SELECT SRCREFERENCENO FROM APP_T_INTERNALRECEIPTHEADER WHERE INTERNALRECEIPTNO = " & iRecNo & ")"
		.ActiveConnection = con
		.Open
	end with
	set dcrs1.ActiveConnection = nothing

	if not dcrs1.EOF then
		sPRODUCTIONORDERNO = trim(dcrs1(0))
	end if
	dcrs1.Close
	
	with dcrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ORDERTYPE FROM PRD_T_PRODUCTIONORDERHEADER WHERE PRODUCTIONORDERNO = (SELECT PRODUCTIONORDERNO FROM PRD_T_DAILYPACKINGHEADER WHERE DAILYPACKINGCODE = (SELECT SRCREFERENCENO FROM APP_T_INTERNALRECEIPTHEADER WHERE INTERNALRECEIPTNO = " & iRecNo & "))"
		.ActiveConnection = con
		.Open
	end with
	set dcrs1.ActiveConnection = nothing

	if not dcrs1.EOF then
		sOrType = trim(dcrs1(0))
	else
		sOrType = ""
	end if
	dcrs1.Close
	iNoofPacks = "Null"
'Response.Write newxml.xml
	if RootNode.HaschildNodes() then
		Response.Write "Y<BR>" ' pls do not block this line
		For Each HeaderNode In RootNode.childNodes
			iItemCode = trim(HeaderNode.Attributes.Item(0).nodeValue)
			iClass = trim(HeaderNode.Attributes.Item(1).nodeValue)
			sOrgCode = trim(HeaderNode.Attributes.Item(2).nodeValue)

			iQtyRec = 0

			iMRSNo = trim(HeaderNode.Attributes.Item(4).nodeValue)
			iOrLot = trim(HeaderNode.Attributes.Item(5).nodeValue)
			iIssNo = trim(HeaderNode.Attributes.Item(6).nodeValue)
			iQtyRec = trim(HeaderNode.Attributes.Item(7).nodeValue)

			arrStore = split(trim(HeaderNode.Attributes.Item(8).nodeValue),":")
			iItmRate = trim(HeaderNode.Attributes.Item(9).nodeValue)
			sSrcType = trim(HeaderNode.Attributes.Item(11).nodeValue)
			'Response.Write "sSrcType="& sSrcType
			
			
			iDINo = trim(HeaderNode.Attributes.Item(12).nodeValue)
			'sIssueEntryNo = trim(HeaderNode.Attributes.Item(13).nodeValue)
			sLoc = arrStore(0)

		' This Code is Added For the Purpose of to Add No of Packs In Ledger Entries
		' To Find Out No of Packs For the Item
			sExp ="//ITEM [ @CCODE = "&iClass&" and @ICODE = "&iItemCode&"]/STAGE/LOTSERIAL"
			Set PackNode = RootNode.Selectnodes(sExp)
			If PackNode.Length > 0 Then
				iNoofPacks = PackNode.Length
			End If
		'	Response.Write "Ok" & iNoofPacks & "ok"
		'	Response.End
		'  --------------------------------------------------------------------------



			if UBound(arrStore) = 1 then
				sBin = arrStore(1)
			else
				sBin = "NULL"
			end if

			' Manufactured Based
			if sSrcType = "F" then
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT SOURCECONFIRMNO FROM VWRECEIPTTYPECHECK WHERE INTERNALRECEIPTNO = " & iRecNo & ""
					'Response.Write dcrs1.Source
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then
					if IsNull(trim(dcrs1(0))) or trim(dcrs1(0)) = "" then
						sSrcType = "RF"
					else
						sSrcType = "RS"
					end if
				else
					sSrcType = "RF"
				end if
				dcrs1.Close
				
				sSql = "Select SrcTypeCode from INV_M_SrcType where SrcType = "& Pack(sSrcType) &""
			    'Response.Write sSql
			    dcrs1.Open sSql,con
			    if not dcrs1.EOF then
			        sSrcCode = dcrs1(0)
			    end if
			    dcrs1.Close 
				
				

				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT ISNULL(PROCESSINGNO,0),ISNULL(SOURCECONFIRMNO,'-') FROM APP_T_INTERNALRECEIPTHEADER WHERE INTERNALRECEIPTNO = " & iRecNo & ""
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing
				if not dcrs.EOF then
					iProcessNo = trim(dcrs(0))
					if trim(dcrs(1)) <> "-" then
						iOrderConfirmationNo = Pack(trim(dcrs(1)))
					else
						iOrderConfirmationNo = "NULL"
					end if
				else
					iOrderConfirmationNo = "NULL"
					iProcessNo = "NULL"
				end if
				dcrs.close

				'Response.Write vbCrLf & sOrType & " >> " & iProcessNo

				'sSrcType = "RF"

				For Each PageNode In HeaderNode.childNodes
					if StrComp(PageNode.nodeName,"STAGE") = 0 then
						iValue = cdbl(iValue) + cdbl(PageNode.Attributes.getNamedItem("IVALUE").Value)
					end if
				next
			' Repacking Based Based
			elseif sSrcType = "R" then
				sSrcType = "RE"
				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT ISNULL(PROCESSINGNO,0),ISNULL(SOURCECONFIRMNO,'-') FROM APP_T_INTERNALRECEIPTHEADER WHERE INTERNALRECEIPTNO = " & iRecNo & ""
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing
				if not dcrs.EOF then
					iProcessNo = trim(dcrs(0))
					if trim(dcrs(1)) <> "-" then
						iOrderConfirmationNo = Pack(trim(dcrs(1)))
					else
						iOrderConfirmationNo = "NULL"
					end if
				else
					iOrderConfirmationNo = "NULL"
					iProcessNo = "NULL"
				end if
				dcrs.close

				'Response.Write vbCrLf & sOrType & " >> " & iProcessNo

				For Each PageNode In HeaderNode.childNodes
					if StrComp(PageNode.nodeName,"STAGE") = 0 then
						iValue = cdbl(iValue) + cdbl(PageNode.Attributes.getNamedItem("IVALUE").Value)
					end if
				next
			' None Based
			elseif sSrcType = "N" then
				sSrcType = "RR"
			' MR or Direct Issue Based
			elseif sSrcType = "M" then
				sSrcType = "RR"
				For Each LotNode In HeaderNode.childNodes
					if StrComp(LotNode.nodeName,"STAGE") = 0 then
						For Each PageNode In LotNode.childNodes
							if StrComp(PageNode.nodeName,"LOTSERIAL") = 0 and StrComp(PageNode.Attributes.Item(0).nodeName,"SERIAL") = 0 then
								iSerial = trim(PageNode.Attributes.Item(0).nodeValue)
								iTareQty = trim(PageNode.Attributes.Item(1).nodeValue)
								If iSerial = "" Then iSerial = 0
								with dcrs1
									.CursorLocation = 3
									.CursorType = 3
									.Source = "SELECT ISNULL(RATE,0) FROM INV_T_RECEIPTLOTDETAILS WHERE SERIALNUMBER = " & iSerial & ""
									.ActiveConnection = con
									'Response.Write  dcrs1.Source
									.Open

								end with
								set dcrs1.ActiveConnection = nothing

								if not dcrs1.EOF then
									iRate = cdbl(dcrs1(0))
								else
									iRate = "0"
								end if
								dcrs1.Close

								iValue = cdbl(iValue) + (cdbl(iTareQty) * cdbl(iRate))
							end if
						next
					end if
				next
			' CCI Released based
			elseif sSrcType = "C" then
				sSrcType = "R"
				For Each PageNode In HeaderNode.childNodes
					if StrComp(PageNode.nodeName,"STAGE") = 0 then
						iValue = cdbl(PageNode.Attributes.getNamedItem("IVALUE").Value)
					end if
				next
			end if

			'Response.Write sSrcType

			if IsNull(iOrderConfirmationNo) or trim(iOrderConfirmationNo) = "" then iOrderConfirmationNo = "NULL"
			if IsNull(iProcessNo) or trim(iProcessNo) = "" then iProcessNo = "NULL"

			iRate = FormatNumber(cdbl(iValue) / cdbl(iQtyRec),2,,,0)

			' None based Internal Receipts
			if iMRSNo = "0" and iDINo = "0" then
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					'.Source = "SELECT ISNULL(MAX(INVENTORYRECEIPTNO)+1,1) FROM INV_T_RECEIPTDETAILS"
					.Source = "SELECT ISNULL(MAX(INVENTORYRECEIPTNO)+1,1) FROM INV_T_LocationLot"
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing
				'Response.Write Pack(FormatDate(date())) & iAccountedBy
				'blocked by Ragav on March 18,2011
'				sSql = "INSERT INTO INV_T_RECEIPTDETAILS (INVENTORYRECEIPTNO,RECEIPTNUMBER,ORGANISATIONCODE," &_
'					"ITEMCODE,CLASSIFICATIONCODE,RECEIPTQUANTITY,ITEMRATE,SRCTYPE," &_
'					"PROCESSINGNO,ORDERCONFIRMATIONNO,ACCOUNTEDON,ACCOUNTEDBY,RECEIVEDON) VALUES " &_
'					"(" & trim(dcrs1(0)) & "," & iRecNo & "," & Pack(sOrgCode) & "," & iItemCode & "," & iClass & "," &_
'					"" & iQtyRec & "," & iRate & "," & Pack(sSrcType) & "," & iProcessNo & "," & iOrderConfirmationNo & "," &_
'					"CONVERT(DATETIME," & Pack(sReceivedOn) & ",103)," & iAccountedBy & ",CONVERT(DATETIME," & Pack(sReceivedOn) & ",103))"
				
				iInvRecNo = trim(dcrs1(0))

				dcrs1.Close
			else
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT ISNULL(SUM(TRANSACTVALUE),0),ISNULL(SUM(TRANSACTQUANTITY),0) FROM INV_T_ITEMLEDGER WHERE ORGANISATIONCODE = " & Pack(sOrgCode) & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND TRANSACTIONTYPE = 'I' and TransactionNo = " & iMRSNo & ""
					.ActiveConnection = con
					.Open
				end with
				'Response.Write dcrs1.Source
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then
					If cDbl(dcrs1(1)) <> 0 then
						iItmRate = cdbl(dcrs1(0)) / cdbl(dcrs1(1))
					Else
						iItmRate = cdbl(dcrs1(0))
					End If
				end if
				dcrs1.Close

				if not iOrLot = "NULL" and sSrcType = "M" then

					with dcrs1
						.CursorLocation = 3
						.CursorType = 3
						if iOrLot <> "NULL" or iOrLot <> "0" then
							.Source = "SELECT ISNULL(INVENTORYRECEIPTNO,0) FROM INV_T_LOCATIONLOT WHERE ORGANISATIONCODE = " & Pack(sOrgCode) & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND LOTNUMBER = " & Pack(iOrLot) & ""
						else
							.Source = "SELECT ISNULL(INVENTORYRECEIPTNO,0) FROM INV_T_LOCATIONLOT WHERE ORGANISATIONCODE = " & Pack(sOrgCode) & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND LOTNUMBER IS NULL)"
						end if
						.ActiveConnection = con
						.Open
					end with

					set dcrs1.ActiveConnection = nothing

					if not dcrs1.EOF then
						iInvRecNo = trim(dcrs1(0))
					end if
					dcrs1.Close

					with dcrs1
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ISNULL(SUM(TRANSACTVALUE),0),ISNULL(SUM(TRANSACTQUANTITY),0) FROM INV_T_ITEMLEDGER WHERE ORGANISATIONCODE = " & Pack(sOrgCode) & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND TRANSACTIONTYPE = 'I' AND TRANSACTIONNO = " & iMRSNo & ""
						.ActiveConnection = con
						.Open
					end with
					set dcrs1.ActiveConnection = nothing

					if not dcrs1.EOF then
						'iItmRate = cdbl(dcrs1(0)) / cdbl(dcrs1(1))
					end if
					dcrs1.Close

				'	sSql = "UPDATE INV_T_RECEIPTDETAILS SET RECEIPTQUANTITY = (RECEIPTQUANTITY + " & iQtyRec & ") WHERE " &_
				'		"INVENTORYRECEIPTNO = " & iInvRecNo & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
				'		"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & ""
				'	'Response.Write sSql & vbCrLf & vbCrLf
				'	'con.Execute sSql

				else
					with dcrs1
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL)"
						.ActiveConnection = con
						.Open
						'Response.Write dcrs1.Source & vbCrLf
					end with

					set dcrs1.ActiveConnection = nothing

					if not dcrs1.EOF then
						If  cdbl(trim(dcrs1(0))) = 0 Then
							'iItmRate = 0
						Else
							'iItmRate = cdbl(trim(dcrs1(1))) / cdbl(trim(dcrs1(0)))
						End If
					end if
					dcrs1.Close

					with dcrs1
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ISNULL(MAX(INVENTORYRECEIPTNO)+1,1) FROM INV_T_LocationLot"
						.ActiveConnection = con
						.Open
					end with
					set dcrs1.ActiveConnection = nothing

			'		sSql = "INSERT INTO INV_T_RECEIPTDETAILS (INVENTORYRECEIPTNO,RECEIPTNUMBER,ORGANISATIONCODE," &_
			'			"ITEMCODE,CLASSIFICATIONCODE,RECEIPTQUANTITY," &_
			'			"ITEMRATE,SRCTYPE,ACCOUNTEDON,ACCOUNTEDBY,RECEIVEDON) VALUES " &_
			'			"(" & trim(dcrs1(0)) & "," & iRecNo & "," & Pack(sOrgCode) & "," & iItemCode & "," & iClass & "," &_
			'			"" & iQtyRec & "," & iItmRate & "," & Pack(sSrcType) & "," &_
			'			"CONVERT(DATETIME," & Pack(FormatDate(sReceivedOn)) & ",103)," & iAccountedBy & ",CONVERT(DATETIME," & Pack(sReceivedOn) & ",103))"
			'		'Response.Write sSql & vbCrLf & vbCrLf
			'		con.Execute sSql

					iInvRecNo = trim(dcrs1(0))

					dcrs1.Close

				end if

			end if

			iValue = cdbl(iItmRate) * cdbl(iQtyRec)

			sSql = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
				"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,NOOFPACKS) VALUES " &_
				"(" & Pack(sOrgCode) & "," & iItemCode & "," & iClass & "," &_
				"" & Pack(sSrcType) & "," & iInvRecNo & ",CONVERT(DATETIME," & Pack(sReceivedOn) & ",103)," & iQtyRec & "," & iValue & "," & iNoofPacks & ")"
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'if len(Month(sReceivedOn)) = 1 then
			'	sTempMonYr = "0"&Month(sReceivedOn)
			'else
			'	sTempMonYr = Month(sReceivedOn)
			'end if
			'sMonYr = sTempMonYr&Year(sReceivedOn)

			if len(mid(sReceivedOn,4,2)) = 1 then
				sTempMonYr = "0"&mid(sReceivedOn,4,2)
			else
				sTempMonYr = mid(sReceivedOn,4,2)
			end if
			sMonYr = sTempMonYr&Year(sReceivedOn)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ITEMCODE FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing
			if dcrs.EOF then
				sSql = "INSERT INTO INV_T_ITEMYEARLYSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
					"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
					"(" & Pack(sOrgCode) & "," & iClass & "," & iItemCode & "," &_
					"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," & iQtyRec & "," & iValue & "," & iQtyRec & "," & iValue & ")"
				'Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql
			else
				sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iQtyRec & ")," &_
					"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ")," &_
					"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iQtyRec & "), " &_
					"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
					"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
					"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
					"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
					"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
				'Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql
			end if
			dcrs.Close

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ITEMCODE FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
				.ActiveConnection = con
				.Open
			end with

			set dcrs.ActiveConnection = nothing

			if dcrs.EOF then
				sSql = "INSERT INTO Inv_T_ItemLocationStock (ORGANISATIONCODE,CLASSIFICATIONCODE," &_
					"ITEMCODE,FINANCIALYEARFROM,FINANCIALYEARTO,LOCATIONNUMBER,BINNUMBER," &_
					"YEARRECEIPTQUANTITY,YEARRECEIPTVALUE,YEARCLOSINGSTOCK,YEARCLOSINGVALUE) VALUES " &_
					"(" & Pack(sOrgCode) & "," & iClass & "," & iItemCode & "," &_
					"CONVERT(DATETIME," & Pack(sFinFrom) & ",103),CONVERT(DATETIME," & Pack(sFinTo) & ",103)," &_
					"" & sLoc & "," & sBin & "," &_
					"" & iQtyRec & "," & iValue & "," & iQtyRec & "," & iValue & ")"
				'Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql

			else
				sSql = "UPDATE Inv_T_ItemLocationStock SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iQtyRec & ")," &_
					"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ")," &_
					"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iQtyRec & "), " &_
					"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
					"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
					"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
					"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
					"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
					"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
				'Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql
			end if
			dcrs.Close

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT SRCREFERENCENO FROM APP_T_INTERNALRECEIPTHEADER WHERE INTERNALRECEIPTNO = " & iRecNo & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs1.ActiveConnection = nothing

			if not dcrs1.EOF then
				iSrcNo = trim(dcrs1(0))
			end if
			dcrs1.Close

			if not iOrLot = "NULL" and sSrcType = "M" then
				For Each LotNode In HeaderNode.childNodes
					For Each PageNode In LotNode.childNodes
						if StrComp(PageNode.nodeName,"LOTSERIAL") = 0 and StrComp(PageNode.Attributes.Item(0).nodeName,"SERIAL") = 0 then
							iSerialEntry = trim(PageNode.Attributes.Item(0).nodeValue)
							iQtyRecEntry = trim(PageNode.Attributes.Item(1).nodeValue)
							sPackingNum = trim(PageNode.Attributes.Item(2).nodeValue)

							if sPackingNum = "" or sPackingNum = "NULL" then
									
								sSql = "UPDATE INV_T_LocationLot SET QUANTITYISSUED = (QUANTITYISSUED - " & iQtyRecEntry & ") WHERE " &_
									"INVENTORYRECEIPTNO = " & iInvRecNo & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
									"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
									"STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND " &_
									"LOTNUMBER = " & Pack(iOrLot) & " AND SERIALNUMBER = " & iSerialEntry & ""' AND PACKINGNUMBER IS NULL"
							else
							
							sSql = "UPDATE INV_T_LocationLot SET QUANTITYISSUED = (QUANTITYISSUED - " & iQtyRecEntry & ") WHERE " &_
									"INVENTORYRECEIPTNO = " & iInvRecNo & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
									"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
									"STORAGELOCATIONNO = " & sLoc & " AND (STORAGEBINNUMBER = " & sBin & " OR STORAGEBINNUMBER IS NULL) AND " &_
									"LOTNUMBER = " & Pack(iOrLot) & " AND SERIALNUMBER = " & iSerialEntry & ""' AND PACKINGNUMBER = " & Pack(sPackingNum) & ""
							end if

							'Response.Write sSql & vbCrLf & vbCrLf
							con.Execute sSql
						elseif StrComp(PageNode.nodeName,"LotSerial") = 0 then
							sQtyIn = trim(PageNode.Attributes.Item(0).nodeValue)
							iTareQty = trim(PageNode.Attributes.Item(1).nodeValue)
							iLot = trim(PageNode.Attributes.Item(2).nodeValue)
							iSerialFrom = trim(PageNode.Attributes.Item(3).nodeValue)
							iSerialTo = trim(PageNode.Attributes.Item(4).nodeValue)
							sTareIn = trim(PageNode.Attributes.Item(5).nodeValue)

							For Each SchNode In PageNode.childNodes
								if StrComp(SchNode.nodeName,"LotSerialDetails") = 0 then
									arrLotSerial = split(trim(SchNode.Attributes.Item(0).nodeValue)," - ")
									iLotEntry = trim(arrLotSerial(0))
									iSerialEntry = trim(arrLotSerial(1))
									iQtyRecEntry = trim(SchNode.Attributes.Item(1).nodeValue)
									iTareQtyEntry = trim(SchNode.Attributes.Item(2).nodeValue)
									iQtyNett = "0"

									iQtyGross = cdbl(iQtyRecEntry) + cdbl(iTareQtyEntry)
									iQtyNett = iQtyRecEntry

									with dcrs
										.CursorLocation = 3
										.CursorType = 3
										.Source = "SELECT ISNULL(MAX(SERIALNUMBER)+1,1) FROM INV_T_LocationLot"
										.ActiveConnection = con
										.Open
									end with
									set dcrs.ActiveConnection = nothing

									if not dcrs.EOF then
										iSerialEntry = dcrs(0)
									end if
									dcrs.Close

									sSql = "INSERT INTO INV_T_LocationLot (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE," &_
										"CLASSIFICATIONCODE,STORAGELOCATIONNO,STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER," &_
										"LOTQUANTITYGROSS,LOTQUANTITYNETT,LOTQUANTITYTARE,PROCESSINGNO,ORDERCONFIRMATIONNO,SRCTYPE,SRCTYPECODE) VALUES " &_
										"(" & iInvRecNo & "," & Pack(sOrgCode) & "," & iItemCode & "," & iClass & "," &_
										"" & sLoc & "," & sBin & "," & Pack(iLotEntry) & "," & iSerialEntry & "," &_
										"" & iQtyGross & "," & iQtyNett & "," & iTareQtyEntry & ","& iProcessNo &","& Pack(iOrderConfirmationNo) &","& Pack(sSrcType) &","& sSrcCode &")"
									'Response.Write sSql & vbCrLf & "LotSerial" & vbCrLf
									con.Execute sSql

								end if
							next
						end if
					next
				next
			else
				if sSrcType = "R" then sTempLot = iOrLot

				For Each LotNode In HeaderNode.childNodes
					sStage = trim(LotNode.Attributes.Item(0).nodeValue)
					iLotValue = trim(LotNode.Attributes.Item(1).nodeValue)
					iQtyNett = trim(LotNode.Attributes.Item(2).nodeValue)
				'	Response.Write iLotValue & "ok"
				'	Response.Write iQtyNett & "ok"
					iRate = FormatNumber(cdbl(iLotValue) / cdbl(iQtyNett),4,,,0)

					if sStage = "" or sStage = "NULL" then
						sStage = "NULL"
					else
						sStage = Pack(sStage)
					end if

					For Each PageNode In LotNode.childNodes
						if StrComp(PageNode.nodeName,"LOTSERIAL") = 0 and StrComp(PageNode.Attributes.Item(0).nodeName,"SERIAL") = 0 then
							iSerial = trim(PageNode.Attributes.Item(0).nodeValue)
							iQtyRecEntry = trim(PageNode.Attributes.Item(1).nodeValue)
							sPackingNum = trim(PageNode.Attributes.Item(2).nodeValue)
							sPackingType = trim(PageNode.Attributes.Item(3).nodeValue)
							sSellingType = trim(PageNode.Attributes.Item(4).nodeValue)
							iWeight = trim(PageNode.Attributes.Item(5).nodeValue)
							iQtyGross = trim(PageNode.Attributes.Item(6).nodeValue)
							sSellingFormType = trim(PageNode.Attributes.Item(7).nodeValue)

							iQtyNett = "0"
							iQtyNett = iQtyRecEntry

							if iQtyGross = "" then iQtyGross = iQtyNett

							if cdbl(iQtyGross) > 0 then
								iTareQtyEntry = cdbl(iQtyGross) - cdbl(iQtyNett)
							else
								iTareQtyEntry = "0"
							end if

							if iSerialEntry = "" then iSerialEntry = "NULL"

							with dcrs
								.CursorLocation = 3
								.CursorType = 3
								.Source = "SELECT ISNULL(MAX(SERIALNUMBER)+1,1) FROM INV_T_LocationLOT"
								.ActiveConnection = con
								.Open
							end with
							set dcrs.ActiveConnection = nothing

							if not dcrs.EOF then
								iSerialEntry = dcrs(0)
							end if
							dcrs.Close

							sSql = "INSERT INTO INV_T_LocationLOT (INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE," &_
								"CLASSIFICATIONCODE,STORAGELOCATIONNO,STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER," &_
								"LOTQUANTITYGROSS,LOTQUANTITYNETT,LOTQUANTITYTARE,PACKINGNUMBER,PACKINGCODE," &_
								"SELLINGNUMBER,WEIGHTPERSELLINGFORM,SELLINGFORM,STAGE,RATE,SRCTYPE,SRCTYPECODE,PROCESSINGNO,ORDERCONFIRMATIONNO) VALUES " &_
								"(" & iInvRecNo & "," & Pack(sOrgCode) & "," & iItemCode & "," & iClass & "," &_
								"" & sLoc & "," & sBin & "," & Pack(sTempLot) & "," & iSerialEntry & "," &_
								"" & iQtyGross & "," & iQtyNett & "," & iTareQtyEntry & "," & Pack(sPackingNum) & "," &_
								"" & sPackingType & "," & sSellingType & "," & iWeight & "," & sSellingFormType & "," & sStage & "," & iRate & ","&Pack(sSrcType) &","& sSrcCode &","& iProcessNo &","& iOrderConfirmationNo &")"
							'Response.Write sSql & vbCrLf & "LotSeria" & vbCrLf
							con.Execute sSql

							iPieceNo = 0
							If iSerial = "" Then iSerial = 0

							with dcrs
								.CursorLocation = 3
								.CursorType = 3
								.Source = "SELECT ISNULL(PACKUNITWEIGHT,0) FROM PRD_T_DAILYPACKINGFORMWEIGHTDETAILS WHERE DAILYPACKINGCODE = " & iSrcNo & " AND PACKINGCODE = " & sPackingType & " AND PACKSERIALNO = " & iSerial & ""
								.ActiveConnection = con
								.Open

							end with
							set dcrs.ActiveConnection = nothing

							do while not dcrs.EOF

								iPieceQty = trim(dcrs(0))

						'		if cdbl(iPieceQty) > 0 then
						'			iPieceNo = iPieceNo + 1
						'			sSql = "INSERT INTO INV_T_RECEIPTFABDETAILS (SERIALNUMBER,PIECENO,QUANTITY) VALUES " &_
						'				"(" & iSerialEntry & "," & iPieceNo & "," & iPieceQty & ")"
						'			'Response.Write sSql & vbCrLf & vbCrLf
						'			con.Execute sSql
						'		end if
							dcrs.MoveNext
							loop
							dcrs.Close

						elseif StrComp(PageNode.nodeName,"LotSerial") = 0 then
							sQtyIn = trim(PageNode.Attributes.Item(0).nodeValue)
							iTareQty = trim(PageNode.Attributes.Item(1).nodeValue)
							iLot = trim(PageNode.Attributes.Item(2).nodeValue)
							iSerialFrom = trim(PageNode.Attributes.Item(3).nodeValue)
							iSerialTo = trim(PageNode.Attributes.Item(4).nodeValue)
							sTareIn = trim(PageNode.Attributes.Item(5).nodeValue)

							For Each SchNode In PageNode.childNodes
								if StrComp(SchNode.nodeName,"LotSerialDetails") = 0 then
									arrLotSerial = split(trim(SchNode.Attributes.Item(0).nodeValue)," - ")
									iLotEntry = trim(arrLotSerial(0))
									iSerialEntry = trim(arrLotSerial(1))
									iQtyRecEntry = trim(SchNode.Attributes.Item(1).nodeValue)
									iTareQtyEntry = trim(SchNode.Attributes.Item(2).nodeValue)
									iQtyNett = "0"

									iQtyGross = cdbl(iQtyRecEntry) + cdbl(iTareQtyEntry)
									iQtyNett = iQtyRecEntry

									with dcrs
										.CursorLocation = 3
										.CursorType = 3
										.Source = "SELECT ISNULL(MAX(SERIALNUMBER)+1,1) FROM INV_T_LocationLOT"
										.ActiveConnection = con
										.Open
									end with
									set dcrs.ActiveConnection = nothing

									if not dcrs.EOF then
										iSerialEntry = dcrs(0)
									end if
									dcrs.Close

									sSql = "INSERT INTO INV_T_LocationLOT(INVENTORYRECEIPTNO,ORGANISATIONCODE,ITEMCODE," &_
										"CLASSIFICATIONCODE,STORAGELOCATIONNO,STORAGEBINNUMBER,LOTNUMBER,SERIALNUMBER," &_
										"LOTQUANTITYGROSS,LOTQUANTITYNETT,LOTQUANTITYTARE,SRCTYPE,SRCTYPECODE,PROCESSINGNO,ORDERCONFIRMATIONNO) VALUES " &_
										"(" & iInvRecNo & "," & Pack(sOrgCode) & "," & iItemCode & "," & iClass & "," &_
										"" & sLoc & "," & sBin & "," & Pack(iLotEntry) & "," & iSerialEntry & "," &_
										"" & iQtyGross & "," & iQtyNett & "," & iTareQtyEntry & ","& Pack(sSrcType) &","& sSrcCode &","& iProcessNo &","& iOrderConfirmationNo &")"
									'Response.Write sSql & vbCrLf & "LotSerialDet" & vbCrLf
									con.Execute sSql

								end if
							next
						end if
					next
				next

			end if
		next

		sSql = "UPDATE APP_T_INTERNALRECEIPTHEADER SET STATUS = 'Y' WHERE " &_
			"INTERNALRECEIPTNO = " & iRecNo & ""
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		''''''''''''''''''''''''''' ' In case of Sales Return - To be resent
		if sOrType = "R" then

			dim iIssQty,iSerialNo,iSerialQty,sUoM,iIssueNo,sDeptNo,iLineNo,sPickLot
			dim dTDate,sItmType

			dTDate = FormatDate(Date)
			sDeptNo = "SAL"

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ITEMTYPEID FROM VWITEM WHERE ClassificationCode = " & iClass & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				sItmType = trim(dcrs(0))
			end if
			dcrs.Close

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(IssueEntryNo)+1,1) FROM INV_T_MaterialIssueHeader"
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				iDINo = dcrs(0)

				sSql = "INSERT INTO INV_T_MaterialIssueHeader (OrganisationCode,IssueEntryNo,IssueDATE," &_
					"IssueTYPE,ITEMTYPEID,ISSUEDFORCODE,REMARKS,IssuedBy,CREATEDON) VALUES " &_
					"(" & Pack(sOrgCode) & "," & iDINo & ",CONVERT(DATETIME,'" & sReceivedOn & "',103),'1'," &_
					"" & Pack(sItmType) & ",'SAL','Sales Return - To be resent'," & iAccountedBy & ",CONVERT(DATETIME,'" & sReceivedOn & "',103))"
				'Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql
			end if
			dcrs.Close

			For Each HeaderNode In RootNode.childNodes
				iItemCode = trim(HeaderNode.Attributes.Item(0).nodeValue)
				iClass = trim(HeaderNode.Attributes.Item(1).nodeValue)
				sOrgCode = trim(HeaderNode.Attributes.Item(2).nodeValue)

				iQtyRec = 0

				iQtyRec = trim(HeaderNode.Attributes.Item(7).nodeValue)

				with dcrs3
					.CursorLocation = 3
					.CursorType = 3
					.Source = "SELECT SERIALNUMBER,LOTQUANTITYNETT,STORAGELOCATIONNO,ISNULL(STORAGEBINNUMBER,0),ISNULL(LOTNUMBER,'-') FROM INV_T_LocationLot WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOTQUANTITYNETT > 0 ORDER BY 1"
					.ActiveConnection = con
					.Open
				end with
				set dcrs3.ActiveConnection = nothing

				do while not dcrs3.EOF
					iSerialNo = trim(dcrs3(0))
					iSerialQty = trim(dcrs3(1))
					sLoc = trim(dcrs3(2))
					sBin = trim(dcrs3(3))
					sPickLot = trim(dcrs3(4))

					if sBin = "0" then sBin = "NULL"
					if sPickLot = "-" then
						sPickLot = "NULL"
					else
						sPickLot = Pack(sPickLot)
					end if

					iIssQty = cdbl(iSerialQty)

					sSql = "UPDATE INV_T_LocationLOT SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iSerialQty & ")" &_
						" WHERE SERIALNUMBER = " & iSerialNo & ""
					'Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql

					' Insert the quantity issued into Direct Issue Details Table
					with dcrs
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT IssueEntryNo FROM INV_T_MaterialISSUEDETAILS WHERE IssueEntryNo = " & iDINo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND (LOTNUMBER = " & sPickLot & " OR LOTNUMBER IS NULL) AND SERIALNUMBER = " & iSerialNo & ""
						.ActiveConnection = con
						.Open
					end with
					set dcrs.ActiveConnection = nothing

					if dcrs.EOF then
						sSql = "INSERT INTO INV_T_MaterialISSUEDETAILS (IssueEntryNo,ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
							"LOCATIONNUMBER,BINNUMBER,LOTNo,SERIALNo,QUANTITYISSUED,ItemEntryNo) VALUES " &_
							"(" & iDINo & "," & Pack(sOrgCode) & "," & iClass & "," & iItemCode & "," &_
							"" & sLoc & "," & sBin & "," & sPickLot & "," & iSerialNo & "," & iIssQty & ",1)"
						'Response.Write sSql & vbCrLf & vbCrLf
						con.Execute sSql
					else
						sSql = "UPDATE INV_T_MaterialISSUEDETAILS SET QUANTITYISSUED = (QUANTITYISSUED + " & iIssQty & ") " &_
							"WHERE IssueEntryNo = " & iDINo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
							"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
							"(LOTNo = " & sPickLot & " OR LOTNo IS NULL) AND SERIALNo = " & iSerialNo & ""
						'Response.Write sSql & vbCrLf & vbCrLf
						con.Execute sSql
					end if
					dcrs.Close

					
					with dcrs
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ISNULL(RATE,0) FROM INV_T_LocationLOT WHERE SERIALNUMBER = " & iSerialNo & ""
						.ActiveConnection = con
						.Open
					end with
					set dcrs.ActiveConnection = nothing

					if not dcrs.EOF then
						iItmRate = cdbl(dcrs(0))
					end if
					dcrs.Close

					iValue = iIssQty * iItmRate

					sSql = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
						"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE,NOOFPACKS) VALUES " &_
						"(" & Pack(sOrgCode) & "," & iItemCode & "," & iClass & "," &_
						"'ID'," & iDINo & ",CONVERT(DATETIME," & Pack(dTDate) & ",103)," & iIssQty & "," & iValue & "," & iNoofPacks & ")"
					'Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql

					
					'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

					with dcrs
						.CursorLocation = 3
						.CursorType = 3
						'.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
						.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
						.ActiveConnection = con
						.Open
					end with
					set dcrs.ActiveConnection = nothing

					if not dcrs.EOF then
						if cdbl(dcrs(0)) < iValue then
							iValue = cdbl(dcrs(0))
						else
							iValue = iValue
						end if
					end if
					dcrs.Close

					sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
						"YEARISSUEVALUE = (YEARISSUEVALUE + " & iValue & ")," &_
						"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
						"YEARCLOSINGVALUE = (YEARCLOSINGVALUE - " & iValue & ") WHERE " &_
						"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
						"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
						"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
						"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
					'Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql


					sSql = "UPDATE Inv_T_ItemLocationStock SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
						"YEARISSUEVALUE = (YEARISSUEVALUE + " & iValue & ")," &_
						"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
						"YEARCLOSINGVALUE = (YEARCLOSINGVALUE - " & iValue & ") WHERE " &_
						"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
						"ORGANISATIONCODE = " & Pack(sOrgCode) & " AND " &_
						"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
						"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
						"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
					'Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

				dcrs3.MoveNext
				loop
				dcrs3.Close

			next

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ACTIONNO FROM SAL_T_SALESRETURNACTION WHERE SALESRETURNNO = (SELECT SOURCEREFFERENCENO FROM RCV_T_GATERECEIPTHEADER WHERE RECEIPTAGAINST = '07' AND GRNNUMBER = (SELECT GRNNUMBER FROM RCV_T_ACTUALRECEIPTHEADER WHERE RECEIPTNUMBER = " & iRecNo & "))"
				'Response.Write dcrs.Source
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing
			if not dcrs.EOF then
				' Sales Return - To be reworked,resent
				if trim(dcrs(0)) = "6" then
					InsertInvoiceDetails "I"
				end if
			end if
			dcrs.close

			InsertInvoiceDetails "I"

		end if
	end if

'upaditing packing status for job order case
if trim(sPRODUCTIONORDERNO) <> "" then 
	
	
	with dcrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "select OrderNumber from Sub_T_SaleOrderDetails where OrderNumber in (select SourceReffNo from Prd_T_ProductionOrderHeader where ProductionORderNo = '" & trim(sPRODUCTIONORDERNO) & "')"
		.ActiveConnection = con
		.Open
	end with
	set dcrs1.ActiveConnection = nothing

	if not dcrs1.EOF then
		Set adoCmd = Server.CreateObject("ADODB.Command")

		Set adoCmd.ActiveConnection = con

		adoCmd.CommandText = "JWK_UpdatePackingStatus"
		adoCmd.CommandType = 4 'adCmdStoredProc
			
		adoCmd.Parameters.Append adoCmd.CreateParameter("@iInvRcptNo", 129,1,250, cint(iInvRecNo))

		adoCmd.Execute()
		'Response.Write "<p>calling new sp"
	end if
	dcrs1.Close
end if 
	
if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & vbCrLf
	next
	'Redirect to Error Handling System
else
	con.RollbackTrans
	Response.End 
	'Response.Write "<p> data not saved"

	con.CommitTrans
end if

con.close
set con = nothing
%>

<%
	Function InsertInvoiceDetails(sInvoiceType)
		Dim iForInvoiceNo, iQty, iTotValue, iRate, iLotQtyGross, iLotQtyNett, sPackNumber
		Dim iCommodity, iPackCode, iSellNo, iWeightSellForm, iSellForm, iCtr
		Dim iItemCtr, iGatePassNo,sSISupplier

		iItemCtr = 0

		sSql = "SELECT ISNULL(PARTYCODE,0) FROM SAL_T_SALESRETURNHEADER WHERE SALETRANSACTIONNO IN (SELECT SALETRANSACTIONNO FROM SAL_T_INVOICEFORDIS WHERE PROCESSINGNO = " & iProcessNo & ")"
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sSql
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then sSISupplier = trim(dcrs(0))
		dcrs.close


		'Fetching ForInvoiceNo
		sSql = "SELECT ISNULL(MAX(FORINVOICENO),0)+1 FROM FORINVOICE_HEADER"
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sSql
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then iForInvoiceNo = trim(dcrs(0))
		dcrs.close

		'Inserting into FORINVOICE_HEADER
		sSql = "INSERT INTO FORINVOICE_HEADER(FORINVOICENO,INVOICEDFORUNIT,REFERENCENO," _
		& "PARTYTYPE,PARTYSUBTYPE,AGENTCODE,PARTYCODE,PARTYLOCATION,TYPEOFSALE,TYPEOFINVOICE," _
		& "INVDESPMODE,INVPAYMENTMODENO,INVTRANSPORTERCODE,INVPYMTTERMS,INVBASISOFPRICINGNO," _
		& "INVLOADINGPORT,INVDESTINATIONPORT,INVCURRENCY,INVPARTYBANKNO,GROSSWEIGHT,NETTWEIGHT," _
		& "NOOFCASES,REMARKS,SALETRANSACTIONNO,CREATEDBY,CREATEDON,TYPEOFITEMS) VALUES(" _
		& iForInvoiceNo & "," & Pack(sOrgCode) & "," & iDINo & ",NULL,NULL,NULL," & sSISupplier & ",NULL," _
		& "NULL," & Pack(sInvoiceType) & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL," _
		& "NULL," & Pack(sReceivedOn) & "," & pack(sItmType) & ")"
		'Response.Write sSql & vbcrlf & vbcrlf
		con.Execute sSql

		'Fetching GatePass No
		sSql = "SELECT ISNULL(MAX(GATEPASSNO),0)+1 FROM FORGATEPASSHEADER"
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sSql
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then iGatePassNo = trim(dcrs(0))
		dcrs.close

		'Inserting GatePass Header
		sSql = "INSERT INTO FORGATEPASSHEADER (GATEPASSNO,FORUNIT,INVOICETYPE," &_
			"FORINVOICENO,PARTYCODE,TYPEOFITEMS,APPLICATIONCODE,MARKEDON,STATUS) VALUES(" &_
			iGatePassNo & "," & Pack(sOrgCode) & "," & Pack(sInvoiceType) & "," &_
			iForInvoiceNo & "," & sSISupplier & "," & Pack(sItmType) & ",4,CONVERT(DATETIME,'" & sReceivedOn & "',103),'N')"

		'Response.Write sSql & vbcrlf & vbcrlf
		con.Execute sSql

		For Each HeaderNode In RootNode.childNodes

			iItemCtr = iItemCtr + 1
			iItemCode = trim(HeaderNode.Attributes.Item(0).nodeValue)
			iClass = trim(HeaderNode.Attributes.Item(1).nodeValue)
			sOrgCode = trim(HeaderNode.Attributes.Item(2).nodeValue)

			iQty = 0

			iQty = trim(HeaderNode.Attributes.Item(7).nodeValue)

			iTotValue  = 0

			with dcrs3
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT SERIALNUMBER,LOTQUANTITYNETT,STORAGELOCATIONNO,ISNULL(STORAGEBINNUMBER,0),ISNULL(LOTNUMBER,'-') FROM INV_T_LOCATIONLOT WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOTQUANTITYNETT > 0 ORDER BY 1"
				.ActiveConnection = con
				.Open
			end with
			set dcrs3.ActiveConnection = nothing

			do while not dcrs3.EOF
				iSerialNo = trim(dcrs3(0))
				iSerialQty = trim(dcrs3(1))

				sSql = "SELECT ISNULL(RATE,0) FROM INV_T_LOCATIONLOT WHERE SERIALNUMBER = " & iSerialNo
				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = sSql
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing

				if not dcrs.eof then iTotValue = cdbl(iTotValue) + (cdbl(iSerialQty) * cdbl(dcrs(0)))
				dcrs.Close
			dcrs3.MoveNext
			loop
			dcrs3.Close

			iRate = cdbl(iTotValue) / cdbl(iQty)

			'Fetching CommodityCode
			sSql = "SELECT ISNULL(COMMODITYCODE,0) FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & pack(sOrgCode)
			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sSql
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				if trim(dcrs(0)) <> "0" then
					iCommodity = trim(dcrs(0))
				else
					iCommodity = "NULL"
				end if
			else
				iCommodity = "NULL"
			end if
			dcrs.close

			'Inserting into FORINVOICE_DETAILS
			sSql = "INSERT INTO FORINVOICE_DETAILS(FORINVOICENO,ITEMCODE,CLASSIFICATIONCODE," _
			& "COMMODITYCODE,QUANTITYFORINVOICE,INVOICEDUOM,INVOICEDRATE,BASICAMOUNT," _
			& "DISCOUNTPERCENT,DISCOUNTAMOUNT,QUANTITYINVOICED,QUANTITYPAID) VALUES(" _
			& iForInvoiceNo & "," & iItemCode & "," & iClass & "," & iCommodity & "," _
			& iQty & "," & pack(sUoM) & "," & iRate & "," & iTotValue & ",0,0,0,0)"
			'Response.Write sSql & vbcrlf & vbcrlf
			con.Execute sSql

			'Inserting into GatePass Details
			sSql = "INSERT INTO FORGATEPASSDETAILS(GATEPASSNO,ENTRYNO,ITEMCODE,CLASSIFICATIONCODE," &_
			"QUANTITY,INVOICEDUOM) VALUES(" & iGatePassNo & "," & iItemCtr & "," & iItemCode  & "," &_
			iClass  & "," & iQty  & "," & pack(sUom) & ")"
			'Response.Write sSql & vbcrlf & vbcrlf
			con.Execute sSql

			with dcrs3
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT SERIALNUMBER,LOTQUANTITYNETT,STORAGELOCATIONNO,ISNULL(STORAGEBINNUMBER,0),ISNULL(LOTNUMBER,'-') FROM INV_T_LOCATIONLOT WHERE INVENTORYRECEIPTNO = " & iInvRecNo & " AND ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND LOTQUANTITYNETT > 0 ORDER BY 1"
				.ActiveConnection = con
				.Open
			end with
			set dcrs3.ActiveConnection = nothing

			do while not dcrs3.EOF
				iSerialNo = trim(dcrs3(0))
				iSerialQty = trim(dcrs3(1))

				sSql = "SELECT LOTQUANTITYGROSS,LOTQUANTITYNETT,PACKINGNUMBER,PACKINGCODE,SELLINGNUMBER,WEIGHTPERSELLINGFORM,SELLINGFORM FROM INV_T_LOCATIONLOT WHERE SERIALNUMBER = " & iSerialNo
				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = sSql
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing

				if not dcrs.eof then
					iLotQtyGross = trim(dcrs(0))
					iLotQtyNett = trim(dcrs(1))
					sPackNumber = trim(dcrs(2))
					iPackCode = trim(dcrs(3))
					if IsNull(dcrs(4)) then
						iSellNo = "NULL"
					else
						iSellNo = trim(dcrs(4))
					end if
					if IsNull(dcrs(5)) then
						iWeightSellForm = "NULL"
					else
						iWeightSellForm = trim(dcrs(5))
					end if
					if IsNull(dcrs(6)) then
						iSellForm = "NULL"
					else
						iSellForm = trim(dcrs(6))
					end if
				end if
				dcrs.Close

				'Inserting into FORINVOICE_PACKDETAILS
				sSql = "INSERT INTO FORINVOICE_PACKDETAILS(FORINVOICENO,ITEMCODE,CLASSIFICATIONCODE," _
				& "PACKINGCODE,PACKNUMBER,PACKGROSSWEIGHT,PACKNETTWEIGHT,PACKNOOFSELLINGFORM," _
				& "WEIGHTPERSELLINGFORM,SELLINGNUMBER,INVENTORYRECEIPTSERIALNO) VALUES(" _
				& iForInvoiceNo & "," & iItemCode & "," & iClass & "," & iPackCode & "," _
				& Pack(sPackNumber) & "," & iLotQtyGross & "," & iLotQtyNett & "," & iSellForm _
				& "," & iWeightSellForm & "," & iSellNo & "," & iSerialNo &")"
				'Response.Write sSql & vbcrlf & vbcrlf
				con.execute sSql

			dcrs3.MoveNext
			loop
			dcrs3.Close

		next

	End Function
%>

