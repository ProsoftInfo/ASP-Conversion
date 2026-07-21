<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	stkMgmtPAInsert.asp
	'Module Name				:	Inventory (Stock Management Physical Adjustment)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	May 28, 2003
	'Modified By				:	Ragavendran R
	'Modified On				:	Dec 22,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	stkMgmtPAEntry.asp
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
<!--#include file="../../include/mrsStatus.asp"-->
<%
'XML DOM Variables
Dim oDOM,Root,objfs,RootO

dim OutData

Set OutData = Server.CreateObject("Microsoft.XMLDOM")
Set RootO = OutData.createElement("Root")
OutData.appendChild RootO

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

dim newxml,RootNode,HeaderNode,PickNode,PickDetNode,ndItem

Set newxml = Server.CreateObject("Microsoft.XMLDOM")

newxml.async = false
newxml.load(Request)
newxml.save server.MapPath("../temp/transaction/PA.xml")
dim dcrs,dcrs1,sSql,bFlag
dim iItemCode,iClass,arrStore,sBin,sLoc,iReqQty,iIssQty,iInvRecNoRec
dim iValue,sOrgID,iTransBy,dTraDate,sMonYr,sMethod,iRecQty,iItmRate
dim arrFin,sFinFrom,sFinTo,sTempMonYr,iYrOpStock,iYrIssQty,iYrCloQty,iYrCloValue
dim sType,iTraNo,iSerialQty,iSerialNo,iSerialAdjQty
dim iPickNo,sPickLoc,sPickBin,sPickLot,sPickQty,iInvRecNo,sSTOrgID,iSTQty
dim iStkQty,iAdjustQty,iIntRecNo,sReason
dim iMRSNo,iMRSAmdNo,iIssueNo,sUoM,iConNo

iTransBy = getUserid
bFlag = true

dTraDate = FormatDate(date())

sTempMonYr = mid(dTraDate,4,2)
sMonYr = sTempMonYr&Year(dTraDate)

arrFin = split(GetFinancialYear(sMonYr),":")
sFinFrom = arrFin(0)
sFinTo = arrFin(1)

Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

Set RootNode = newxml.documentElement

'iItemCode = trim(RootNode.Attributes.getNamedItem("ITEM").Value)
'iClass = trim(RootNode.Attributes.getNamedItem("CLASS").Value)
'sOrgID = trim(RootNode.Attributes.getNamedItem("ORG").Value)
'sReason = trim(RootNode.Attributes.getNamedItem("REASON").Value)

con.beginTrans

if RootNode.HaschildNodes() then
	For Each ndItem In RootNode.childNodes
		iItemCode = ndItem.getAttribute("ICode")
		iClass = ndItem.getAttribute("CCode")
		sOrgID = ndItem.getAttribute("Unit")
		sReason = ndItem.getAttribute("Reason")
		For each HeaderNode in ndItem.childNodes
				if StrComp(HeaderNode.nodeName,"LOCDET") = 0 then
					sLoc = trim(HeaderNode.Attributes.getNamedItem("LOC").Value)
					sBin = trim(HeaderNode.Attributes.getNamedItem("BIN").Value)

					bFlag = true

					if sBin <> "0" then
						sBin = sBin
					else
						sBin = "NULL"
					end if

					For Each PickNode In HeaderNode.childNodes
						if StrComp(PickNode.nodeName,"PICK") = 0 then
							sPickLot = trim(PickNode.Attributes.getNamedItem("LOTNO").Value)
							iInvRecNo = trim(PickNode.Attributes.getNamedItem("INVRECNO").Value)
							iStkQty = trim(PickNode.Attributes.getNamedItem("QTYSTK").Value)
							iIssQty = cdbl(PickNode.Attributes.getNamedItem("QTYISS").Value)
							iAdjustQty = cdbl(PickNode.Attributes.getNamedItem("ADJUSTED").Value)

							if iInvRecNo = "" or IsNull(iInvRecNo) then iInvRecNo = "NULL"

							if sPickLot <> "NULL" then sPickLot = Pack(sPickLot)

							'with dcrs
							'	.CursorLocation = 3
							'	.CursorType = 3
							'	.Source = "SELECT ISNULL(MAX(TRANSFERNO)+1,1) FROM INV_T_INTERNALSTOCKTRANSFER"
							'	.ActiveConnection = con
							'	.Open
							'end with
							'set dcrs.ActiveConnection = nothing

							'if not dcrs.EOF then
							'	iTraNo = dcrs(0)
							'end if
							'dcrs.close
							iValue = 0
							' If PA is positive (Receipt)
							if iAdjustQty > 0 then
								'Check for Lot
								'if not sPickLot = "NULL" then
								if PickNode.hasChildNodes() then
									For Each PickDetNode In PickNode.childNodes
										iSerialNo = trim(PickDetNode.Attributes.getNamedItem("SERIALNO").Value)
										iSerialAdjQty = trim(PickDetNode.Attributes.getNamedItem("QTY").Value)

										with dcrs
											.CursorLocation = 3
											.CursorType = 3
											'.Source = "SELECT ISNULL(RATE,0) FROM INV_T_RECEIPTLOTDETAILS WHERE SERIALNUMBER = " & iSerialNo & ""
											
											If iSerialNo <> "" Then
												.Source = "SELECT ISNULL(RATE,0) FROM INV_T_LOCATIONLOT WHERE SERIALNUMBER = " & iSerialNo & ""
											Else
												.Source = "SELECT ISNULL(RATE,0) FROM INV_T_LOCATIONLOT WHERE ITEMCODE="& iItemCode &" AND CLASSIFICATIONCODE = "& iClass &" AND ORGANISATIONCODE ='"& sOrgID &" AND LOTNUMBER IS NULL AND STORAGELOCATIONNO ="& sLoc &"' AND (STORAGEBINNUMBER = "& sBin &" OR STORAGEBINNUMBER IS NULL) "
											End IF
											'Response.Write dcrs.source
											.ActiveConnection = con
											.Open
										end with
										set dcrs.ActiveConnection = nothing

										if not dcrs.EOF then
											iItmRate = trim(dcrs(0))
											iValue = cdbl(iValue) + (cdbl(iItmRate) * cdbl(iSerialAdjQty))
										end if
										dcrs.Close

									next

									with dcrs
										.CursorLocation = 3
										.CursorType = 3
										.Source = "SELECT ISNULL(MAX(INTERNALRECEIPTNO)+1,1) FROM APP_T_INTERNALRECEIPTHEADER"
										.ActiveConnection = con
										.Open
									end with
									set dcrs.ActiveConnection = nothing

									if not dcrs.EOF then
										iIntRecNo = trim(dcrs(0))
									end if
									dcrs.Close

									ReceiptInsert sPickLot,"0",iAdjustQty,iInvRecNo,iTraNo,iValue,iIntRecNo,"H",sReason

									For Each PickDetNode In PickNode.childNodes
										iSerialNo = trim(PickDetNode.Attributes.getNamedItem("SERIALNO").Value)
										iSerialQty = trim(PickDetNode.Attributes.getNamedItem("STKQTY").Value)
										iSerialAdjQty = trim(PickDetNode.Attributes.getNamedItem("QTY").Value)
										if not (cdbl(iSerialAdjQty)) = 0 then
											ReceiptInsert sPickLot,iSerialNo,iSerialAdjQty,iInvRecNo,iTraNo,iValue,iIntRecNo,"D",sReason
										end if
									next
								' not from the Lot
								else
									StockReceiptInsert iAdjustQty,iTraNo,sReason
								end if
							' If PA is negative (Issue)
							elseif iAdjustQty < 0 then
								'Check for Lot
								'if not sPickLot = "NULL" then
									if PickNode.hasChildNodes() then
										with dcrs
											.CursorLocation = 3
											.CursorType = 3
											.Source = "SELECT ISNULL(MAX(INTERNALRECEIPTNO)+1,1) FROM APP_T_INTERNALRECEIPTHEADER"
											.ActiveConnection = con
											.Open
										end with
										set dcrs.ActiveConnection = nothing

										if not dcrs.EOF then
											iIntRecNo = trim(dcrs(0))
										end if
										dcrs.Close

										'''''''''''''''' MRS CREATION ''''''''''''''''''''''''''''''''''''''''''
										with dcrs
											.CursorLocation = 3
											.CursorType = 3
											.Source = "SELECT ISNULL(MAX(MRSNUMBER)+1,1) FROM INV_T_MRSHEADER"
											.ActiveConnection = con
											.Open
										end with
										set dcrs.ActiveConnection = nothing

										if not dcrs.EOF then
											iMRSNo = trim(dcrs(0))
										end if
										dcrs.close

										'Status  MRS Created
										sSql = "INSERT INTO INV_T_MRSHEADER (MRSFORUNIT,MRSNUMBER,MRSDATE," &_
											"MRSTYPE,MRSHEADERSTATUS,ISSUEDFORCODE,CREATEDBY,CREATEDON,DEPTNO,SOURCEREFNO,GENERATEDFROM) VALUES " &_
											"(" & Pack(sOrgID) & "," & iMRSNo & ",CONVERT(DATETIME,GETDATE(),103),'1','040101'," &_
											"'OTH'," & iTransBy & ",CONVERT(DATETIME,GETDATE(),103)," &_
											"'OTH'," & Pack(iTraNo) & ",4)"
										'Response.Write sSql & vbCrLf & vbCrLf
										con.Execute sSql

										'Status  MRS Created
										sSql = "INSERT INTO INV_T_MRSITEMDETAILS (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
											"ITEMCODE,LOCATIONNUMBER,BINNUMBER,REQUIREDBY,REQUIREDVALUE,QUANTITYREQUESTED,MRSITEMSTATUS) VALUES " &_
											"(" & iMRSNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
											"" & sLoc & "," & sBin & ",'I'," & Pack(dTraDate) & "," & iAdjustQty * -1 & ",'040101')"
										'Response.Write sSql & vbCrLf & vbCrLf
										con.Execute sSql

										'''''''''''''''' MRS AMEND CREATION ''''''''''''''''''''''''''''''''''''''''''
										with dcrs
											.CursorLocation = 3
											.CursorType = 3
											.Source = "SELECT ISNULL(MAX(MRSAMENDNUMBER)+1,1) FROM INV_A_MRSHEADER"
											.ActiveConnection = con
											.Open
										end with
										set dcrs.ActiveConnection = nothing

										if not dcrs.EOF then
											iMRSAmdNo = trim(dcrs(0))
										end if
										dcrs.close

										sSql = "INSERT INTO INV_A_MRSHEADER EXECUTE ('SELECT " & iMRSAmdNo & " AS MRSAMENDNUMBER," &_
											"CONVERT(DATETIME,GETDATE(),103) AS MRSAMENDDATE,''Physical Adjustment Approval'' AS AMENDREMARKS," &_
											"" & iTransBy & " AS AMENDENDBY,CONVERT(DATETIME,GETDATE(),103) AS AMENDENDON," &_
											"''AP'' AS AMENDORAPPROVE,* FROM INV_T_MRSHEADER WHERE MRSNUMBER = " & iMRSNo & "') "
										'Response.Write sSql & vbCrLf & vbCrLf
										con.Execute sSql

										sSql = "INSERT INTO INV_A_MRSITEMDETAILS EXECUTE ('SELECT " & iMRSAmdNo & "," &_
											"* FROM INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNo & "') "
										'Response.Write sSql & vbCrLf & vbCrLf
										con.Execute sSql

										'Status  MRS Approved
										sSql = "UPDATE INV_T_MRSHEADER SET APPROVEDBY = " & iTransBy & ", APPROVEDON = CONVERT(DATETIME,GETDATE(),103) WHERE MRSNUMBER = " & iMRSNo & ""
										con.Execute sSql

										sSql = "DELETE INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNo & ""
										con.Execute sSql

										'Status  Item Issued
										sSql = "INSERT INTO INV_T_MRSITEMDETAILS (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
											"ITEMCODE,LOCATIONNUMBER,BINNUMBER,REQUIREDBY,REQUIREDVALUE,QUANTITYREQUESTED,QUANTITYAPPROVED) VALUES " &_
											"(" & iMRSNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
											"" & sLoc & "," & sBin & ",'I'," & Pack(dTraDate) & "," & iAdjustQty * -1 & "," & iAdjustQty * -1 & ")"
										'Response.Write sSql & vbCrLf & vbCrLf
										con.Execute sSql

										' Function Call to Update the Line Status of an MR for Inventory Application
										MRLineStatusUpdate "Requisition","Approved",iMRSNo,iItemCode,iClass,sOrgID,"4","",""

										' Function Call to Update the Header Status of an MR
										'MRStatusUpdate "Requisition","Approved",iMRSNo
										MRStatusUpdate "Requisition","Approved",iMRSNo,"","","",sOrgID

								'''''''''''''''''''''''''''''''INTERNAL ST INSERT OF PA ISSUE'''''''''''''''''''''''''''

										'sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFER (TRANSFERNO,ITEMCODE,ORGANISATIONCODE," &_
										'	"CLASSIFICATIONCODE,FROMLOCATIONNUMBER," &_
										'	"FROMBINNUMBER,TRANSFERTYPE,TRANSFERREDON,TRANSFERREDBY,REASON) VALUES " &_
										'	"(" & iTraNo & "," & iItemCode & "," & Pack(sOrgID) & "," & iClass & "," & sLoc & "," &_
										'	"" & sBin & ",'PI',CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iTransBy & "," & Pack(sReason) & ")"
										'Response.Write sSql & vbCrLf & vbCrLf
										'con.Execute sSql
								'''''''''''''''''''''''''''''''DEPARTMENT STOCK INSERT''''''''''''''''''''''''''''''
										with dcrs
											.CursorLocation = 3
											.CursorType = 3
											.Source = "SELECT ISNULL(MAX(ISSUENO)+1,1) FROM INV_T_DEPARTMENTSTOCK"
											.ActiveConnection = con
											.Open
										end with
										set dcrs.ActiveConnection = nothing

										if not dcrs.EOF then
											iIssueNo = trim(dcrs(0))
										end if
										dcrs.close

										with dcrs
											.CursorLocation = 3
											.CursorType = 3
											.Source = "SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
											.ActiveConnection = con
											.Open
										end with
										set dcrs.ActiveConnection = nothing

										if not dcrs.EOF then
											sUoM = trim(dcrs(0))
										end if
										dcrs.Close

										'sSql = "INSERT INTO INV_T_DEPARTMENTSTOCK (ISSUENO,DEPTNO,MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
										'	"ITEMCODE,QUANTITYISSUED,QUANTITYUOM) VALUES " &_
										'	"(" & iIssueNo & ",'OTH'," & iMRSNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
										'	"" & iAdjustQty * -1 & "," & Pack(sUoM) & ")"
										'Response.Write sSql & vbCrLf & vbCrLf
										'con.Execute sSql

								''''''''''''''''''''''''''''''''''MAT CONSUMPTION UPDATE ''''''''''''''''''''''''''''''''''
										'with dcrs
										'	.CursorLocation = 3
										'	.CursorType = 3
										'	.Source = "SELECT ISNULL(MAX(CONSUMPTIONNO)+1,1) FROM APP_T_MATERIALCONSUMPTION "
										'	.ActiveConnection = con
										'	.Open
										'end with
										'set dcrs.ActiveConnection = nothing
										'if not dcrs.EOF then
										'	iConNo = trim(dcrs(0))
										'end if
										'dcrs.Close

										'sSql = "INSERT INTO APP_T_MATERIALCONSUMPTION (CONSUMPTIONNO,ISSUENO,CONSUMEDBYDEPT,ORGANISATIONCODE," &_
										'	"CLASSIFICATIONCODE,ITEMCODE,QUANTITYCONSUMED,ENTEREDON,ENTEREDBY) VALUES " &_
										'	"(" & iConNo & "," & iIssueNo & ",'OTH'," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
										'	"" & iAdjustQty * -1 & ",CONVERT(DATETIME,GETDATE(),103)," & iTransBy & ")"
										'Response.Write sSql & vbCrLf & vbCrLf
										'con.Execute sSql

										'sSql = "UPDATE INV_T_DEPARTMENTSTOCK SET QUANTITYCONSUMED = " &_
										'	"(ISNULL(QUANTITYCONSUMED,0) + " & iAdjustQty * -1 & ") WHERE ISSUENO = " & iIssueNo & " AND " &_
										'	"MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
										'	"CLASSIFICATIONCODE = " & iClass & " AND ITEMCODE = " & iItemCode & ""
										'Response.Write sSql & vbCrLf
										'con.Execute sSql


										For Each PickDetNode In PickNode.childNodes
											iSerialNo = trim(PickDetNode.Attributes.Item(0).nodeValue)
											iSerialQty = trim(PickDetNode.Attributes.Item(1).nodeValue)
											iSerialAdjQty = trim(PickDetNode.Attributes.Item(2).nodeValue)
											if not (cdbl(iSerialAdjQty)) = 0 then
												IssueInsert sPickLot,iSerialNo,iSerialAdjQty * -1,iInvRecNo,iTraNo,iIntRecNo,iIssueNo,iConNo
											end if
										next

										''''''''''''''''''''''''''''''''''''Stock Status Updation'''''''''''''''''''''''''''''''''''''''''''''''''''''
											' Function Call to Update the Header Status of an MR
											'MRStatusUpdate "Issue","Create",iMRSNo
											MRStatusUpdate "Issue","Create",iMRSNo,"","","",sOrgID
										''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
									'not from the Lot
									else
										StockIssueInsert iAdjustQty * -1,iTraNo
									end if
								'end if
							end if
						end if
					next
				' end if for LOCDET Node check
				end if
		next 'For each HeaderNode in ndItem.childNodes
	next
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
	response.end
	response.clear
	con.CommitTrans
end if

con.close
set con = nothing

' Function to Insert Internal receipt Details and account it
Function ReceiptInsert(sPickLotP,iSerialNoP,iIssQtyP,iInvRecNoP,iTraNoP,iItmRateP,iIntRecNoP,bCheck,sReason)
	' Declaration of variables
	Dim dcrs
	dim iLineNo,sUoM,iItmRate,iValue,iRecNo,iQtyRec

	dim iIssQty,sPickLot,iIntRecNo,sDept
	dim iInvRecNo,iTraNo,iSerialNo,iSerialQty,iAppCode

	' OTHERS Department
	sDept = "OTH"
	' Inventory Application
	iAppCode = "4"

	sPickLot = trim(sPickLotP)
	iSerialNo = trim(iSerialNoP)
	iIssQty = cdbl(trim(iIssQtyP))
	iInvRecNo = trim(iInvRecNoP)
	iTraNo = trim(iTraNoP)
	iItmRate = trim(iItmRateP)
	iIntRecNo = iIntRecNoP

	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	if bCheck = "H" then

		'sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFER (TRANSFERNO,ITEMCODE,ORGANISATIONCODE," &_
		'	"CLASSIFICATIONCODE,FROMLOCATIONNUMBER," &_
		'	"FROMBINNUMBER,TRANSFERTYPE,TRANSFERREDON,TRANSFERREDBY,SRCREFNO,REASON) VALUES " &_
		'	"(" & iTraNo & "," & iItemCode & "," & Pack(sOrgID) & "," & iClass & "," & sLoc & "," &_
		'	"" & sBin & ",'PR'," &_
		'	"CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iTransBy & "," & iIntRecNo & "," & Pack(sReason) & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		'con.Execute sSql

		sSql = "INSERT INTO APP_T_INTERNALRECEIPTHEADER (INTERNALRECEIPTNO,APPLICATIONCODE," &_
			"ORGANISATIONCODE,CREATEDFROMDEPT,REFTYPE,CREATEDON,CREATEDBY,STATUS) VALUES " &_
			"(" & iIntRecNo & "," & iAppCode & "," & Pack(sOrgID) & "," & Pack(sDept) & ",'P'," &_
			"CONVERT(DATETIME,GETDATE(),103)," & iTransBy & ",'Y')"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		'sSql = "UPDATE INV_T_LOCATIONLOT SET ACCEPTQUANTITY = (ACCEPTQUANTITY + " & iIssQty & ") WHERE " &_
		'	"INVENTORYRECEIPTNO = " & iInvRecNo & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
		'	"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & ""
		'Response.Write sSql & vbCrLf & vbCrLf
		'con.Execute sSql

		iQtyRec = 0
		iQtyRec = iIssQty

		iValue = cdbl(iItmRate)

		sSql = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
			"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE) VALUES " &_
			"(" & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
			"'RP'," & iInvRecNo & ",CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iQtyRec & "," & iValue & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) "	'AND MONTHANDYEAR = " & Pack(sMonYr) & "
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
				"LOCATIONNUMBER,BINNUMBER,YearReceiptQuantity,YearReceiptValue) VALUES " &_
				"(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
				"" & sLoc & "," & sBin & "," & iQtyRec & "," & iValue & ")"
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YearReceiptQuantity = (YearReceiptQuantity + " & iQtyRec & ")," &_
				"YearReceiptValue = (YearReceiptValue + " & iValue & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
				"CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
				"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL)" 
				
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close

		
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iQtyRec & ")," &_
			"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ")," &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iQtyRec & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql
		
		sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iQtyRec & ")," &_
			"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ")," &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iQtyRec & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	elseif bCheck = "D" then
		'sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFERDETAILS (INTERTRANSFERNO,LOTNUMBER,SERIALNUMBER," &_
		'	"QUANTITY) VALUES " &_
		'	"(" & iTraNo & "," & sPickLot & "," & iSerialNo & "," & iIssQty & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		'con.Execute sSql
		If iSerialNo = "" Then iSerialNo = 0
		sSql = "INSERT INTO APP_T_INTERNALRECEIPTDETAILS (INTERNALRECEIPTNO,MRSNUMBER," &_
			"CLASSIFICATIONCODE,ITEMCODE,LOTNO,SERIALNO,QUANTITYRETURN) VALUES " &_
			"(" & iIntRecNo & ",NULL," & iClass & "," & iItemCode & "," & sPickLot & "," &_
			"" & iSerialNo & "," & iIssQty & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_LOCATIONLOT SET LOTQUANTITYNETT = (LOTQUANTITYNETT + " & iIssQty & ") WHERE " &_
			"SERIALNUMBER = " & iSerialNo & ""

		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_LOCATIONLOT SET LOTQUANTITYGROSS = (LOTQUANTITYNETT + LOTQUANTITYTARE) WHERE " &_
			"SERIALNUMBER = " & iSerialNo & ""

		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql
	end if
end function

'''***************************************************************************''''''''

' Function to Insert in case of Clean Stock PA Receipt
Function StockReceiptInsert(iIssQtyP,iTraNoP,sReason)
	' Declaration of variables
	Dim dcrs
	dim iValue,sMonYr
	dim iYrOpStock,iYrIssQty,iYrCloQty,iYrCloValue
	dim iIssQty,iTraNo,iQtyRec

	iIssQty = cdbl(iIssQtyP)
	iTraNo = iTraNoP

	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(YEARRECEIPTQUANTITY,0),YEAROPENINGSTOCK,YEARISSUEQUANTITY,YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM INV_T_ITEMLOCYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iRecQty = cdbl(dcrs(0))
		iYrOpStock = cdbl(dcrs(1))
		iYrIssQty = cdbl(dcrs(2))
		iYrCloQty = cdbl(dcrs(3))
		iYrCloValue = cdbl(dcrs(4))
	end if
	dcrs.Close

	' check for Year Opening and Issue Stock, if Issue Stock is there then
	' Issue from the Stock available
	if iYrOpStock > iYrIssQty then
		' check for Issue Qty greater than Available Qty
		if iIssQty > (iYrOpStock - iYrIssQty) then
			' Issue the remaining Qty
			iIssQty = iYrOpStock - iYrIssQty
		else
			iIssQty = iIssQty
		end if
		iValue = iIssQty * (iYrCloValue / iYrCloQty)
	elseif iYrCloQty > 0 then
		' check for Issue Qty greater than Closing Qty
		if iIssQty > iYrCloQty then
			' Issue the remaining Qty
			iIssQty = iYrCloQty
		else
			iIssQty = iIssQty
		end if
		iValue = iIssQty * (iYrCloValue / iYrCloQty)
	end if

	if 1 = 1 then

		'sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFER (TRANSFERNO,ITEMCODE,ORGANISATIONCODE," &_
		'	"CLASSIFICATIONCODE,FROMLOCATIONNUMBER," &_
		'	"FROMBINNUMBER,TRANSFERTYPE,TRANSFERREDON,TRANSFERREDBY,REASON) VALUES " &_
		'	"(" & iTraNo & "," & iItemCode & "," & Pack(sOrgID) & "," & iClass & "," & sLoc & "," &_
		'	"" & sBin & ",'PR',CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iTransBy & "," & Pack(sReason) & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		'con.Execute sSql

		'sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFERDETAILS (INTERTRANSFERNO,QUANTITY) VALUES " &_
		'	"(" & iTraNo & "," & iIssQty & ")"
		''Response.Write sSql & vbCrLf & vbCrLf
		'con.Execute sSql

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		iQtyRec = iIssQty

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
				"LOCATIONNUMBER,BINNUMBER,YearReceiptQuantity,YearReceiptValue) VALUES " &_
				"(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
				"" & sLoc & "," & sBin & "," & iQtyRec & "," & iValue & ")"
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iQtyRec & ")," &_
				"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
				"CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
				"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) "
				
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close

		
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		sSql = "UPDATE INV_T_ITEMYEARLYSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iQtyRec & ")," &_
			"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ")," &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iQtyRec & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARRECEIPTQUANTITY = (YEARRECEIPTQUANTITY + " & iQtyRec & ")," &_
			"YEARRECEIPTVALUE = (YEARRECEIPTVALUE + " & iValue & ")," &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK + " & iQtyRec & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE + " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

	end if

end function

'''***************************************************************************''''''''
' Function to Insert in case of LOT Qty PA Issue
Function IssueInsert(sPickLotP,iSerialNoP,iIssQtyP,iInvRecNoP,iTraNoP,iIntRecNoP,iIssueNoP,iConNoP)
	' Declaration of variables
	Dim dcrs
	dim iLineNo,sUoM,iItmRate,iValue,iRecNo,iQtyRec,iIssueNo

	dim iIssQty,sPickLot,iIntRecNo,sDept
	dim iInvRecNo,iTraNo,iSerialNo,iSerialQty,iAppCode
	dim iConNo
	' OTHERS Department
	sDept = "OTH"
	' Inventory Application
	iAppCode = "4"

	sPickLot = trim(sPickLotP)
	iSerialNo = trim(iSerialNoP)
	iIssQty = cdbl(trim(iIssQtyP))
	iInvRecNo = trim(iInvRecNoP)
	iTraNo = trim(iTraNoP)
	iIntRecNo = iIntRecNoP
	iIssueNo = iIssueNoP
	iConNo = iConNoP

	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(RATE,0) FROM INV_T_LOCATIONLOT WHERE SERIALNUMBER = " & iSerialNo & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iItmRate = trim(dcrs(0))
	end if
	dcrs.Close

	iValue = cdbl(iIssQty) * cdbl(iItmRate)

	sSql = "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iIssQty & ")" &_
		" WHERE SERIALNUMBER = " & iSerialNo & ""
	'Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql

'''''''''''''''''''''''''''''''INTERNAL ST INSERT OF PA ISSUE'''''''''''''''''''''''''''

	'sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFERDETAILS (INTERTRANSFERNO,LOTNUMBER,SERIALNUMBER,QUANTITY) VALUES " &_
	'	"(" & iTraNo & "," & sPickLot & "," & iSerialNo & "," & iIssQty & ")"
	'Response.Write sSql & vbCrLf & vbCrLf
	'con.Execute sSql

'''''''''''''''''''''''''''''''DEPARTMENT STOCK INSERT''''''''''''''''''''''''''''''

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(MAX(LINENUMBER)+1,1) FROM INV_T_DEPARTMENTSTOCKISSUEDETAILS"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iLineNo = trim(dcrs(0))
	end if
	dcrs.Close

	' insert for Department Stock Breakup
	'sSql = "INSERT INTO INV_T_DEPARTMENTSTOCKISSUEDETAILS (LINENUMBER,ISSUENO,ISSUEDATE,LOTNO,SERIALNO,QUANTITYISSUED) VALUES " &_
	'	"(" & iLineNo & "," & iIssueNo & ",CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & sPickLot & "," & iSerialNo & "," & iIssQty & ")"
	'Response.Write sSql & vbCrLf & vbCrLf
	'con.Execute sSql

''''''''''''''''''''''''''''''''''''''MAT CONSUMPTION DETAILS UPDATE'''''''''''''''''''''''''''''''''''''''''''
	'with dcrs
	'	.CursorLocation = 3
	'	.CursorType = 3
	'	.Source = "SELECT ISNULL(MAX(LINENUMBER)+1,1) FROM APP_T_MATERIALCONSUMPTIONDETAIL"
	'	.ActiveConnection = con
	'	.Open
	'end with
	'set dcrs.ActiveConnection = nothing

	'if not dcrs.EOF then
	'	iLineNo = trim(dcrs(0))
	'end if
	'dcrs.Close

	'sSql = "INSERT INTO APP_T_MATERIALCONSUMPTIONDETAIL (LINENUMBER,CONSUMPTIONNO,ISSUENO,ISSUEDATE,LOTNO,SERIALNO,QUANTITYCONSUMED) VALUES " &_
	'	"(" & iLineNo & "," & iConNo & "," & iIssueNo & ",CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & sPickLot & "," & iSerialNo & "," & iIssQty & ")"
	'Response.Write sSql & vbCrLf & vbCrLf
	'con.Execute sSql

	'if sPickLot <> "NULL" then
	'	sSql = "UPDATE INV_T_DEPARTMENTSTOCKISSUEDETAILS SET QUANTITYCONSUMED = " &_
	'		"(ISNULL(QUANTITYCONSUMED,0) + " & iIssQty & ") WHERE ISSUENO = " & iIssueNo & " AND " &_
	'		"CONVERT(CHAR,ISSUEDATE,103) = CONVERT(CHAR," & Pack(dTraDate) & ",103) AND " &_
	'		"LOTNO = " & sPickLot & " AND SERIALNO = " & iSerialNo & ""
	'else
	'	sSql = "UPDATE INV_T_DEPARTMENTSTOCKISSUEDETAILS SET QUANTITYCONSUMED = " &_
	'		"(ISNULL(QUANTITYCONSUMED,0) + " & iIssQty & ") WHERE ISSUENO = " & iIssueNo & " AND " &_
	'		"CONVERT(CHAR,ISSUEDATE,103) = CONVERT(CHAR," & Pack(dTraDate) & ",103) AND " &_
	'		"LOTNO IS NULL AND SERIALNO = " & iSerialNo & ""
	'end if

	'Response.Write sSql & vbCrLf
	'con.Execute sSql
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	sSql = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
		"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE) VALUES " &_
		"(" & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
		"'IP'," & iMRSNo & ",CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iIssQty & "," & iValue & ")"
	'Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql

	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL)"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if dcrs.EOF then
		sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
			"LOCATIONNUMBER,BINNUMBER,YearIssueQuantity,YearIssueValue) VALUES " &_
			"(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
			"" & sLoc & "," & sBin & "," & iIssQty & "," & iValue & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql
	else
		sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
			"YEARISSUEVALUE = (YEARISSUEVALUE + " & iValue & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
			"CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) "
			
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql
	end if
	dcrs.Close
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		'.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
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
		"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
		"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
		"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
	'Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql

	sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
		"YEARISSUEVALUE = (YEARISSUEVALUE + " & iValue & ")," &_
		"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
		"YEARCLOSINGVALUE = (YEARCLOSINGVALUE - " & iValue & ") WHERE " &_
		"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
		"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
		"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
		"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
		"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
	'Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql
	
	sSql = "UPDATE INV_T_MRSITEMDETAILS SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iIssQty & ") " &_
		"WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
		"ORGANISATIONCODE = " & Pack(sOrgID) & " AND MRSNUMBER = " & iMRSNo & ""
	'Response.Write sSql & vbCrLf & vbCrLf
	con.execute sSql

	'''''''''''''Status Updation'''''''''''''''''''''''''''''''''''''''''''''
	' Function Call to Update the Line Status of an MR for Inventory Application
	MRLineStatusUpdate "Issue","Create",iMRSNo,iItemCode,iClass,sOrgID,"4","F",""
	''''''''''''''''''''''''''''''''''''''''''''''''''''

end function

'''***************************************************************************''''''''
' Function to Insert in case of Clean Stock PA Issue
Function StockIssueInsert(iIssQtyP,iTraNoP)
	' Declaration of variables
	Dim dcrs
	dim iLineNo,iItmRate,iValue
	dim iYrOpStock,iYrIssQty,iYrCloQty,iYrCloValue,iCounter

	dim iIssQty,iMRSAmdNo,sUoM,iIssueNo
	dim iTraNo,iMRSNo,sDept,iAppCode,sAmendRemarks,iConNo

	iIssQty = cdbl(iIssQtyP)
	iTraNo = iTraNoP

	' OTHERS Department
	sDept = "OTH"
	' Inventory Application
	iAppCode = "4"
	' Amended Remarks
	sAmendRemarks = "Physical Adjustment Approval"

	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	'''''''''''''''' MRS CREATION ''''''''''''''''''''''''''''''''''''''''''
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(MAX(MRSNUMBER)+1,1) FROM INV_T_MRSHEADER"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iMRSNo = trim(dcrs(0))
	end if
	dcrs.close

	'Status  MRS Created for Approval
	sSql = "INSERT INTO INV_T_MRSHEADER (MRSFORUNIT,MRSNUMBER,MRSDATE," &_
		"MRSTYPE,MRSHEADERSTATUS,ISSUEDFORCODE,CREATEDBY,CREATEDON,DEPTNO,SOURCEREFNO,GENERATEDFROM) VALUES " &_
		"(" & Pack(sOrgID) & "," & iMRSNo & ",CONVERT(DATETIME,GETDATE(),103),'1','040101'," &_
		"" & Pack(sDept) & "," & iTransBy & ",CONVERT(DATETIME,GETDATE(),103)," &_
		"" & Pack(sDept) & "," & Pack(iTraNo) & "," & iAppCode & ")"
	'Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql

	'Status  MRS Created
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(MAX(ICounter)+1,1) FROM INV_T_MRSITEMDETAILS where ItemCode = "& iItemCode &" and Classificationcode="& iClass &" "
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iCounter = trim(dcrs(0))
	end if
	dcrs.close
	sSql = "INSERT INTO INV_T_MRSITEMDETAILS (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
		"ITEMCODE,LOCATIONNUMBER,BINNUMBER,REQUIREDBY,REQUIREDVALUE,QUANTITYREQUESTED,MRSITEMSTATUS,ICounter) VALUES " &_
		"(" & iMRSNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
		"" & sLoc & "," & sBin & ",'I'," & Pack(dTraDate) & "," & iIssQty & ",'040101',"& ICounter &")"
	Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql

	'''''''''''''''' MRS AMEND CREATION ''''''''''''''''''''''''''''''''''''''''''

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(MAX(MRSAMENDNUMBER)+1,1) FROM INV_A_MRSHEADER"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iMRSAmdNo = trim(dcrs(0))
	end if
	dcrs.close

	sSql = "INSERT INTO INV_A_MRSHEADER EXECUTE ('SELECT " & iMRSAmdNo & " AS MRSAMENDNUMBER," &_
		"CONVERT(DATETIME,GETDATE(),103) AS MRSAMENDDATE,'" & Pack(sAmendRemarks) & "' AS AMENDREMARKS," &_
		"" & iTransBy & " AS AMENDENDBY,CONVERT(DATETIME,GETDATE(),103) AS AMENDENDON," &_
		"''AP'' AS AMENDORAPPROVE,* FROM INV_T_MRSHEADER WHERE MRSNUMBER = " & iMRSNo & "') "
	'Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql

	sSql = "INSERT INTO INV_A_MRSITEMDETAILS EXECUTE ('SELECT " & iMRSAmdNo & "," &_
		"* FROM INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNo & "') "
	'Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql

	sSql = "UPDATE INV_T_MRSHEADER SET APPROVEDBY = " & iTransBy & ", APPROVEDON = CONVERT(DATETIME,GETDATE(),103) WHERE MRSNUMBER = " & iMRSNo & ""
	con.Execute sSql

	sSql = "DELETE INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNo & ""
	con.Execute sSql

	sSql = "INSERT INTO INV_T_MRSITEMDETAILS (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
		"ITEMCODE,LOCATIONNUMBER,BINNUMBER,REQUIREDBY,REQUIREDVALUE,QUANTITYREQUESTED,QUANTITYAPPROVED) VALUES " &_
		"(" & iMRSNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
		"" & sLoc & "," & sBin & ",'I'," & Pack(dTraDate) & "," & iIssQty & "," & iIssQty & ")"
	'Response.Write sSql & vbCrLf & vbCrLf
	con.Execute sSql

	' Function Call to Update the Line Status of an MR for Inventory Application
	'MRLineStatusUpdate "Requisition","Approved",iMRSNo,iItemCode,iClass,sOrgID,"4","",""
	MRLineStatusUpdate "Requisition","Approved",iMRSNo,iItemCode,iClass,iCounter,sOrgID,"4","","",""

	' Function Call to Update the Header Status of an MR
	'MRStatusUpdate "Requisition","Approved",iMRSNo
	MRStatusUpdate "Requisition","Approved",iMRSNo,"","","",sOrgID

	'''''''''''''''' END OF MRS CREATION ''''''''''''''''''''''''''''''''''''''''''
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(YEARRECEIPTQUANTITY,0),YEAROPENINGSTOCK,YEARISSUEQUANTITY,YEARCLOSINGSTOCK,YEARCLOSINGVALUE FROM Inv_T_ItemLocationStock WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iRecQty = cdbl(dcrs(0))
		iYrOpStock = cdbl(dcrs(1))
		iYrIssQty = cdbl(dcrs(2))
		iYrCloQty = cdbl(dcrs(3))
		iYrCloValue = cdbl(dcrs(4))
	end if
	dcrs.Close
	' check for Year Opening and Issue Stock, if Issue Stock is there then
	' Issue from the Stock available
	if iYrOpStock > iYrIssQty then
		' check for Issue Qty greater than Available Qty
		if iIssQty > (iYrOpStock - iYrIssQty) then
			' Issue the remaining Qty
			iIssQty = iYrOpStock - iYrIssQty
		else
			iIssQty = iIssQty
		end if
		iValue = iIssQty * (iYrCloValue / iYrCloQty)
	elseif iYrCloQty > 0 then
		' check for Issue Qty greater than Closing Qty
		if iIssQty > iYrCloQty then
			' Issue the remaining Qty
			iIssQty = iYrCloQty
		else
			iIssQty = iIssQty
		end if
		iValue = iIssQty * (iYrCloValue / iYrCloQty)
	end if

	if 1 = 1 then

'''''''''''''''''''''''''''''''INTERNAL ST INSERT OF PA ISSUE'''''''''''''''''''''''''''
		'sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFER (TRANSFERNO,ITEMCODE,ORGANISATIONCODE," &_
		'	"CLASSIFICATIONCODE,FROMLOCATIONNUMBER," &_
		'	"FROMBINNUMBER,TRANSFERTYPE,TRANSFERREDON,TRANSFERREDBY,REASON) VALUES " &_
		'	"(" & iTraNo & "," & iItemCode & "," & Pack(sOrgID) & "," & iClass & "," & sLoc & "," &_
		'	"" & sBin & ",'PI',CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iTransBy & "," & Pack(sReason) & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		'con.Execute sSql

		'sSql = "INSERT INTO INV_T_INTERNALSTOCKTRANSFERDETAILS (INTERTRANSFERNO,QUANTITY) VALUES " &_
		'	"(" & iTraNo & "," & iIssQty & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		'con.Execute sSql

'''''''''''''''''''''''''''''''DEPARTMENT STOCK INSERT''''''''''''''''''''''''''''''
	If 1 = 2 Then
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(MAX(ISSUENO)+1,1) FROM INV_T_DEPARTMENTSTOCK"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			iIssueNo = trim(dcrs(0))
		end if
		dcrs.close

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			sUoM = trim(dcrs(0))
		end if
		dcrs.Close

		sSql = "INSERT INTO INV_T_DEPARTMENTSTOCK (ISSUENO,DEPTNO,MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
			"ITEMCODE,QUANTITYISSUED,QUANTITYUOM) VALUES " &_
			"(" & iIssueNo & "," & Pack(sDept) & "," & iMRSNo & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
			"" & iIssQty & "," & Pack(sUoM) & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(MAX(LINENUMBER)+1,1) FROM INV_T_DEPARTMENTSTOCKISSUEDETAILS"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			iLineNo = trim(dcrs(0))
		end if
		dcrs.Close

		' insert for Department Stock Breakup
		sSql = "INSERT INTO INV_T_DEPARTMENTSTOCKISSUEDETAILS (LINENUMBER,ISSUENO,ISSUEDATE,QUANTITYISSUED) VALUES " &_
			"(" & iLineNo & "," & iIssueNo & ",CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iIssQty & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

''''''''''''''''''''''''''''''''''MAT CONSUMPTION UPDATE ''''''''''''''''''''''''''''''''''
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(MAX(CONSUMPTIONNO)+1,1) FROM APP_T_MATERIALCONSUMPTION "
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if not dcrs.EOF then
			iConNo = trim(dcrs(0))
		end if
		dcrs.Close

		sSql = "INSERT INTO APP_T_MATERIALCONSUMPTION (CONSUMPTIONNO,ISSUENO,CONSUMEDBYDEPT,ORGANISATIONCODE," &_
			"CLASSIFICATIONCODE,ITEMCODE,QUANTITYCONSUMED,ENTEREDON,ENTEREDBY) VALUES " &_
			"(" & iConNo & "," & iIssueNo & "," & Pack(sDept) & "," & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
			"" & iIssQty & ",CONVERT(DATETIME,GETDATE(),103)," & iTransBy & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_DEPARTMENTSTOCK SET QUANTITYCONSUMED = " &_
			"(ISNULL(QUANTITYCONSUMED,0) + " & iIssQty & ") WHERE ISSUENO = " & iIssueNo & " AND " &_
			"MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CLASSIFICATIONCODE = " & iClass & " AND ITEMCODE = " & iItemCode & ""
		'Response.Write sSql & vbCrLf
		con.Execute sSql

''''''''''''''''''''''''''''''''''''''MAT CONSUMPTION DETAILS UPDATE'''''''''''''''''''''''''''''''''''''''''''
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(MAX(LINENUMBER)+1,1) FROM APP_T_MATERIALCONSUMPTIONDETAIL"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			iLineNo = trim(dcrs(0))
		end if
		dcrs.Close

		sSql = "INSERT INTO APP_T_MATERIALCONSUMPTIONDETAIL (LINENUMBER,CONSUMPTIONNO,ISSUENO,ISSUEDATE,LOTNO,QUANTITYCONSUMED) VALUES " &_
			"(" & iLineNo & "," & iConNo & "," & iIssueNo & ",CONVERT(DATETIME," & Pack(dTraDate) & ",103),NULL," & iIssQty & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_DEPARTMENTSTOCKISSUEDETAILS SET QUANTITYCONSUMED = " &_
			"(ISNULL(QUANTITYCONSUMED,0) + " & iIssQty & ") WHERE ISSUENO = " & iIssueNo & " AND " &_
			"CONVERT(CHAR,ISSUEDATE,103) = CONVERT(CHAR," & Pack(dTraDate) & ",103) AND " &_
			"LOTNO IS NULL AND (SERIALNO = 0 OR SERIALNO IS NULL)"

		'Response.Write sSql & vbCrLf
		con.Execute sSql
	End If 'If 1 = 2 Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		sSql = "INSERT INTO INV_T_ITEMLEDGER (ORGANISATIONCODE,ITEMCODE,CLASSIFICATIONCODE," &_
			"TRANSACTIONTYPE,TRANSACTIONNO,TRANSACTIONDATE,TRANSACTQUANTITY,TRANSACTVALUE) VALUES " &_
			"(" & Pack(sOrgID) & "," & iItemCode & "," & iClass & "," &_
			"'IP'," & iMRSNo & ",CONVERT(DATETIME," & Pack(dTraDate) & ",103)," & iIssQty & "," & iValue & ")"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ITEMCODE FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) "
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if dcrs.EOF then
			sSql = "INSERT INTO INV_T_ITEMLOCATIONSTOCK (ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
				"LOCATIONNUMBER,BINNUMBER,YEARISSUEQUANTITY,YEARISSUEVALUE) VALUES " &_
				"(" & Pack(sOrgID) & "," & iClass & "," & iItemCode & "," &_
				"" & sLoc & "," & sBin & "," & iIssQty & "," & iValue & ")"
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		else
			sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
				"YEARISSUEVALUE = (YEARISSUEVALUE + " & iValue & ") WHERE ITEMCODE = " & iItemCode & " AND " &_
				"CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
				"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
				"MONTHANDYEAR = " & Pack(sMonYr) & ""
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if
		dcrs.Close

		
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMYEARLYSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
			.Source = "SELECT ISNULL(YEARCLOSINGVALUE,0) FROM INV_T_ITEMLOCATIONSTOCK WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
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
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "UPDATE INV_T_ITEMLOCATIONSTOCK SET YEARISSUEQUANTITY = (YEARISSUEQUANTITY + " & iIssQty & ")," &_
			"YEARISSUEVALUE = (YEARISSUEVALUE + " & iValue & ")," &_
			"YEARCLOSINGSTOCK = (YEARCLOSINGSTOCK - " & iIssQty & "), " &_
			"YEARCLOSINGVALUE = (YEARCLOSINGVALUE - " & iValue & ") WHERE " &_
			"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
			"ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
			"LOCATIONNUMBER = " & sLoc & " AND (BINNUMBER = " & sBin & " OR BINNUMBER IS NULL) AND " &_
			"CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND " &_
			"CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
		'Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		
	end if

	sSql = "UPDATE INV_T_MRSITEMDETAILS SET QUANTITYISSUED = (ISNULL(QUANTITYISSUED,0) + " & iIssQty & ") " &_
		"WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND " &_
		"ORGANISATIONCODE = " & Pack(sOrgID) & " AND MRSNUMBER = " & iMRSNo & ""
	'Response.Write sSql & vbCrLf & vbCrLf
	con.execute sSql

	'''''''''''''Status Updation'''''''''''''''''''''''''''''''''''''''''''''
	' Function Call to Update the Line Status of an MR for Inventory Application
	MRLineStatusUpdate "Issue","Create",iMRSNo,iItemCode,iClass,sOrgID,"4","F",""
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	'''''''''''''''''''Status Updation'''''''''''''''''''''''''''''''''''''''
	' Function Call to Update the Header Status of an MR
	'MRStatusUpdate "Issue","Create",iMRSNo
	MRStatusUpdate "Issue","Create",iMRSNo,"","","",sOrgID
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

end function

'''***************************************************************************''''''''
%>

