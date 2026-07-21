<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ArrBinDetInsert.asp
	'Module Name				:	Inventory (Inventory)
	'Author Name				:	UmaMaheswari S
	'Created On					:	June 01, 2011
	'Modified By				:	
	'Modified On				:	
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!--#include file="../../include/populate.asp"-->
<%
'XML DOM Variables
Dim Root,objfs,newxml,RootNode,HeaderNode,ndItem,BinNode,Node
dim dcrs,dcrs1,dcrs2,sSql,sSql1,iItemCode,iClass,arrStore,sBinNo,nBinQty,sStorageLocNo
dim iValue,sOrgID,iTransBy,dTraDate,sMonYr,arrFin,sFinFrom,sFinTo,sTempMonYr
Dim sBinNoStatus,nInvRecNo

Set objfs = CreateObject("Scripting.FileSystemObject")
Set newxml = Server.CreateObject("Microsoft.XMLDOM")

newxml.async = false
newxml.load(Request)

iTransBy = getUserid
dTraDate = FormatDate(date())

sTempMonYr = mid(dTraDate,4,2)
sMonYr = sTempMonYr&Year(dTraDate)

arrFin = split(GetFinancialYear(sMonYr),":")
sFinFrom = arrFin(0)
sFinTo = arrFin(1)
'Response.Write "<p>sFinFrom="&sFinFrom &","& sfinto
Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

Set RootNode = newxml.documentElement

'newxml.save server.MapPath("../temp/Transaction/BinDetData.xml")
con.beginTrans

if RootNode.HaschildNodes() then
	For Each ndItem In RootNode.childNodes
		iItemCode = ndItem.getAttribute("ICode")
		iClass = ndItem.getAttribute("CCode")
		sOrgID = ndItem.getAttribute("Unit")
		
		For Each HeaderNode in ndItem.childNodes
			if StrComp(HeaderNode.nodeName,"LOCDET") = 0 then
				
				sStorageLocNo = trim(HeaderNode.Attributes.getNamedItem("LOC").Value)
				sBinNoStatus = ""
				
				'Check Storage Bin No Is There or Not
				with dcrs
					.CursorLocation = 3
					.CursorType = 3
					.Source ="Select isNull(StorageBinNumber,0) from Inv_T_LocationLot where Itemcode = "& iItemCode &" and classificationcode ="& iClass &" and StorageLocationNo = '"& sStorageLocNo &"' AND Year(DateofReceipt) >= Year('"&sFinFrom&"') AND Year(DateofReceipt) <= '"&Year(sFinTo)&"'"
					.ActiveConnection = con
					.Open
				end with
				set dcrs.ActiveConnection = nothing
				If Not dcrs.EOF Then
					If dcrs(0) <> "0" Then
						sBinNoStatus = "Y"
					Else
						sBinNoStatus = "N"
					End IF
				Else
					sBinNoStatus = "N"
				End IF
				dcrs.Close 
				
				For Each BinNode in HeaderNode.childNodes 
					If BinNode.nodeName = "STOREBINDET" Then
						For Each Node in BinNode.childNodes 
						
						'sBinNo = trim(BinNode.Attributes.getNamedItem("BINNO").Value)
						'nBinQty = trim(BinNode.Attributes.getNamedItem("QTY").Value)
						sBinNo = trim(Node.Attributes.getNamedItem("NO").Value)
						nBinQty = trim(Node.Attributes.getNamedItem("QTY").Value)
						
						'Response.Write "<p>Data="&nBinQty & "---"& sBinNo & "--" & sBinNoStatus
						'Response.Write "<p>sBinNoStatus="&sBinNoStatus
						
						'Again Check Storage Bin No Is There or Not For selected Bin
						If sBinNo <> "0" Then
							with dcrs
								.CursorLocation = 3
								.CursorType = 3
								.Source ="Select isNull(StorageBinNumber,0) from Inv_T_LocationLot where Itemcode = "& iItemCode &" and classificationcode ="& iClass &" and StorageLocationNo = '"& sStorageLocNo &"' AND StorageBinNumber = '"& sBinNo &"' AND Year(DateofReceipt) >= Year('"&sFinFrom&"') AND Year(DateofReceipt) <= '"&Year(sFinTo)&"'"
								.ActiveConnection = con
								.Open
							end with
							set dcrs.ActiveConnection = nothing
							If Not dcrs.EOF Then
								If dcrs(0) <> "0" Then
									sBinNoStatus = "Y"
								Else
									sBinNoStatus = "N"
								End IF
							Else
								sBinNoStatus = "N"
							End IF
							dcrs.Close 
						
						End IF 'If sBinNo <> "0" Then
						'Response.Write "<p>Data="&nBinQty & "---"& sBinNo & "--" & sBinNoStatus
						
						'Inv_T_LocationLot
						
							with dcrs
								.CursorLocation = 3
								.CursorType = 3
								If sBinNoStatus = "Y" Then
									.Source ="Select Distinct InventoryReceiptNo from Inv_T_LocationLot where Itemcode = "& iItemCode &" and classificationcode ="& iClass &" and StorageLocationNo = '"& sStorageLocNo &"' and StorageBinNumber = '"& sBinNo &"' AND Year(DateofReceipt) >= Year('"&sFinFrom&"') AND Year(DateofReceipt) <= '"&Year(sFinTo)&"'"
								Else
									.Source ="Select Distinct InventoryReceiptNo from Inv_T_LocationLot where Itemcode = "& iItemCode &" and classificationcode ="& iClass &" and StorageLocationNo = '"& sStorageLocNo &"'  AND Year(DateofReceipt) >= Year('"&sFinFrom&"') AND Year(DateofReceipt) <= '"&Year(sFinTo)&"'"
								End IF
								'Response.Write dcrs.Source 
								.ActiveConnection = con
								.Open
							end with
							
							set dcrs.ActiveConnection = nothing
							
							If Not dcrs.EOF Then
								Do while Not dcrs.EOF 
									If sBinNoStatus = "Y" Then
										
										'sSql = "Update Inv_T_LocationLot SET LotQuantityGross = "& nBinQty &" ,LotQuantityNett = "& nBinQty &" where Itemcode = "& iItemCode &" and classificationcode ="& iClass &" and StorageLocationNo = '"& sStorageLocNo &"' and StorageBinNumber = '"& sBinNo &"' AND InventoryReceiptNo = "& dcrs(0) &" "
										
										sSql = "select LotQuantityGross,LotQuantityNett,InventoryReceiptNo From Inv_T_LocationLot Where Itemcode = "& iItemCode &" and classificationcode ="& iClass &" and StorageLocationNo = '"& sStorageLocNo &"' and StorageBinNumber = '"& sBinNo &"' AND InventoryReceiptNo = "& dcrs(0) &" "
										with dcrs2 
											.CursorLocation = 3
											.CursorType = 3
											.Source = sSql 
											.ActiveConnection = con
											.Open 
										End with
										If Not dcrs2.EOF Then
											sSql1 = "Update Inv_T_LocationLot SET LotQuantityGross = "& dcrs2(0) &" ,LotQuantityNett = "& dcrs2(1) &" where Itemcode = "& iItemCode &" and classificationcode ="& iClass &" and StorageLocationNo = '"& sStorageLocNo &"' and StorageBinNumber = '"& sBinNo &"' AND InventoryReceiptNo = "& dcrs2(2) &" "
											'Response.Write "<p>Query1"&sSql1 & vbCrLf 
											con.Execute sSql1
										End IF
										dcrs2.Close 
										
									Else
										nInvRecNo = nInvRecNo & "," & dcrs(0)
									End IF	'If sBinNoStatus = "Y" Then
									dcrs.MoveNext 
								Loop
							End IF
							dcrs.Close 
							
							If sBinNoStatus = "N" Then
								If nInvRecNo <> "" Then  nInvRecNo = Mid(nInvRecNo,2)
							
								If nBinQty <> "0" Then
									sSql = "Update Inv_T_LocationLot SET StorageBinNumber='"& sBinNo &"' where Itemcode = "& iItemCode &" and classificationcode ="& iClass &" and StorageLocationNo = '"& sStorageLocNo &"' AND InventoryReceiptNo IN ("& nInvRecNo &") "
									'Response.Write "<p>AAAAAAA"&sSql & vbCrLf 
									con.Execute sSql
								End IF
							End IF	'If sBinNoStatus = "Y" Then
							
						'Inv_t_ItemLocationStock
						
						If sBinNoStatus = "Y" Then
							
							with dcrs
								.CursorLocation = 3
								.CursorType = 3
								.Source ="Select YearOpeningStock,YearOpeningValue,YearClosingStock,YearClosingValue from Inv_t_ItemLocationStock where Itemcode = "& iItemCode &" and classificationcode ="& iClass &" and LocationNumber = '"& sStorageLocNo &"' and BinNumber = '"& sBinNo &"' AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
								.ActiveConnection = con
								.Open
							end with
							set dcrs.ActiveConnection = nothing
							If Not dcrs.EOF Then
								'sSql = "Update Inv_t_ItemLocationStock  SET YearOpeningStock = "& nBinQty &" ,YearOpeningValue="& nBinQty &" , YearClosingStock="& nBinQty &" , YearClosingValue="& nBinQty &" where Itemcode = "& iItemCode &" and classificationcode ="& iClass &" and LocationNumber = '"& sStorageLocNo &"' and BinNumber = '"& sBinNo &"'"
								sSql = "Update Inv_t_ItemLocationStock  SET YearOpeningStock = "& dcrs(0) &" ,YearOpeningValue="& dcrs(1) &" , YearClosingStock="& dcrs(2) &" , YearClosingValue="& dcrs(3) &" where Itemcode = "& iItemCode &" and classificationcode ="& iClass &" and LocationNumber = '"& sStorageLocNo &"' and BinNumber = '"& sBinNo &"'"
								'Response.Write "<p>Query2"&sSql & vbCrLf 
								con.Execute sSql
							End IF
							dcrs.Close 
							
						Else
							
							If nBinQty <> "0" Then
								sSql = "Update Inv_t_ItemLocationStock  SET BinNumber = '"& sBinNo &"' where Itemcode = "& iItemCode &" and classificationcode ="& iClass &" and LocationNumber = '"& sStorageLocNo &"' AND CONVERT(DATETIME," & Pack(sFinFrom) & ",103) >= CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) <= CONVERT(DATETIME,FINANCIALYEARTO,103)"
								'Response.Write "<p>XXXX"&sSql & vbCrLf 
								con.Execute sSql
							Else
								'Insert 1 Row 
								sSql = " Insert into Inv_t_ItemLocationStock(OrganisationCode,ClassificationCode,ItemCode,FinancialYearFrom,FinancialYearTo,LocationNumber,BinNumber,"&_
									   " YearOpeningStock,YearOpeningValue,YearReceiptQuantity,YearReceiptValue,YearIssueQuantity,YearIssueValue,YearClosingStock,YearClosingValue,"&_
									   " YearReserved,YearOnHold,YearRejected,YearConsumed,SrcType) VALUES('"& sOrgID &"',"& iClass &","& iItemCode &",Convert(DateTime,'"& sFinFrom &"',103),"&_
									   " convert(DateTime,'"& sFinTo &"',103),'"& sStorageLocNo &"','"& sBinNo &"',0,0,0,0,0,0,0,0,0,0,0,0,NULL)"
								'Response.Write "<p>YYYY"&sSql & vbCrLf 
								con.Execute sSql
							End IF	'If nBinQty <> "0" Then
							
						End IF	'If sBinNoStatus = "Y" Then
						
						'Inv_M_ItemStorage Table Updation
						
						If sBinNoStatus = "Y" Then
						Else
							If nBinQty <> "0" Then
								sSql = "Update Inv_M_ItemStorage SET BinNumber ='"& sBinNo &"' Where Itemcode = "& iItemCode &" and classificationcode ="& iClass &" and LocationNumber = '"& sStorageLocNo &"'"
								'Response.Write "<p>MMMM"&sSql & vbCrLf 
								con.Execute sSql
							Else
								sSql = "Select ApplicableFor,AllowTransfers From Inv_M_ItemStorage Where Itemcode = "& iItemCode &" and classificationcode ="& iClass &" and LocationNumber = '"& sStorageLocNo &"' "
								dcrs.Open sSql,con
								If Not dcrs.EOF Then
									sSql = " Insert into Inv_M_ItemStorage(ItemCode,ClassificationCode,OrganisationCode,ApplicableFor,LocationNumber,BinNumber,AllowTransfers) VALUES"&_
										   " ("& iItemCode &","& iClass &",'"& sOrgID &"','"& dcrs(0)&"','"& sStorageLocNo &"','"& sBinNo &"','"& dcrs(1) &"')"
									'Response.Write "<p>NNNN"&sSql & vbCrLf 
									con.Execute sSql
								End IF
								dcrs.Close 
							End IF
						End IF	'If sBinNoStatus = "Y" Then
						
						Next 'For Each Node in BinNode.childNodes 
					End IF	'If BinNode.nodeName = "STROEBINDET" Then
					
				Next
				
			end if	'if StrComp(HeaderNode.nodeName,"LOCDET") = 0 then
		next
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
	'con.RollbackTrans
	'Response.End 
	con.CommitTrans
end if

con.close
set con = nothing
%>