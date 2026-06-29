<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XMLModuleActivity.asp
	'Module Name				:	Admin  (Master)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	FEB 17,2010
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
<!--#include file="../../include/populate.asp"-->
<%
	Dim sOrgUnit,sAppCode,drSet,Root,objDOM,SubNode,sQuery
	Dim nCount,arrSize
	Dim arrActCode(),arrActName()
	Set drSet = Server.CreateObject("ADODB.RecordSet")
	
	Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
	
	objDOM.load(request)
	sOrgUnit = Request.QueryString("sUnitId")
	sAppCode = Request.QueryString("hApplicationNo")
	
	set Root = objDOM.createElement("Root")
		objDOM.appendChild Root
	
	if sAppCode = 1 then ' Accounts
		Redim arrActCode(8)
		Redim arrActName(8)
		arrActCode(0) = "0" :arrActCode(1) = "01":	arrActCode(2) = "02":arrActCode(3) = "04"
		arrActCode(4) = "05":arrActCode(5) = "06":arrActCode(6) = "07":arrActCode(7) = "08"
		
		arrActName(0) = "Select":arrActName(1) = "Cash Book":arrActName(2) = "Bank Day Book":arrActName(3) = "Purchase Journal"
		arrActName(4) = "Sales Journal":arrActName(5) = "Debit Note":arrActName(6) = "Credit Note":	arrActName(7) = "General Journals"
		
		for nCount = 0 to ubound(arrActCode) - 1 			
			set SubNode = objDOM.createElement("Activity")
				SubNode.setAttribute "ActCode",arrActCode(nCount)
				SubNode.setAttribute "ActName",arrActName(nCount)
				Root.appendChild SubNode
		next
		
	elseif sAppCode = 2 then ' Purchase
		sQuery = "SELECT ActivityNumber,ActivityName FROM Pur_M_Activities order by ActivityNumber"
	elseif sAppCode = 3 then ' Sales
		Redim arrActCode(12)
		Redim arrActName(12)
		arrActCode(0) = "0" :arrActCode(1) = "QUT":	arrActCode(2) = "ORD":arrActCode(3) = "OCR"
		arrActCode(4) = "ORP":arrActCode(5) = "DIS":arrActCode(6) = "PIS":arrActCode(7) = "PUR"
		arrActCode(8) = "RIS":arrActCode(9) = "PFO":arrActCode(10) = "INV":	arrActCode(11) = "FJJ"
		
		arrActName(0) = "Select":arrActName(1) = "Quotation":arrActName(2) = "Order Creation":arrActName(3) = "Order Confirmation"
		arrActName(4) = "Order Processing":	arrActName(5) = "Despatch Instruction Slip":arrActName(6) = "Production Instruction Slip"
		arrActName(7) = "Purchase Instruction Slip":arrActName(8) = "Repack Instruction Slip":arrActName(9) = "Proforma Invoice"
		arrActName(10) = "Invoice":	arrActName(11) = "Form JJ"

		for nCount = 0 to ubound(arrActCode) - 1 			
			set SubNode = objDOM.createElement("Activity")
				SubNode.setAttribute "ActCode",arrActCode(nCount)
				SubNode.setAttribute "ActName",arrActName(nCount)
				Root.appendChild SubNode
		next

	elseif sAppCode = 4 then ' Inventory
		Redim arrActCode(7)
		Redim arrActName(7)
		arrActCode(0) = "0" :arrActCode(1) = "LO":	arrActCode(2) = "MR":arrActCode(3) = "IS"
		arrActCode(4) = "PN":arrActCode(5) = "SL":arrActCode(6) = "DC"

		arrActName(0) = "Select":arrActName(1) = "Lot Number":arrActName(2) = "MR Number":arrActName(3) = "Issue Number"
		arrActName(4) = "Packing Number":	arrActName(5) = "Sample Label":arrActName(6) = "DC - Gate Pass"

		for nCount = 0 to ubound(arrActCode) - 1 			
			set SubNode = objDOM.createElement("Activity")
				SubNode.setAttribute "ActCode",arrActCode(nCount)
				SubNode.setAttribute "ActName",arrActName(nCount)
				Root.appendChild SubNode
		next

	elseif sAppCode = 5 then ' Maintenance
		sQuery = "SELECT ActivityNumber,ActivityName FROM MTN_M_ActivitiesForNoSeries order by ActivityNumber"
	elseif sAppCode = 6 then ' Production
	elseif sAppCode = 7 then ' Fixed Assets
	elseif sAppCode = 8 then ' Fixed Deposits
		Redim arrActCode(4)
		Redim arrActName(4)
		arrActCode(0) = "0" :arrActCode(1) = "FDR":	arrActCode(2) = "FLN":arrActCode(3) = "LON"
		
		arrActName(0) = "Select":arrActName(1) = "Deposit Receipt No":arrActName(2) = "Folio No":arrActName(3) = "Loan Sanction No"
		
		for nCount = 0 to ubound(arrActCode) - 1 			
			set SubNode = objDOM.createElement("Activity")
				SubNode.setAttribute "ActCode",arrActCode(nCount)
				SubNode.setAttribute "ActName",arrActName(nCount)
				Root.appendChild SubNode
		next

	elseif sAppCode = 9 then ' Trade Deposits
		Redim arrActCode(3)
		Redim arrActName(3)
		arrActCode(0) = "0" :arrActCode(1) = "TDR":	arrActCode(2) = "TLN"
		
		arrActName(0) = "Select":arrActName(1) = "Deposit Receipt No":arrActName(2) = "Trade Deposit Holder No"
		
		for nCount = 0 to ubound(arrActCode) - 1 			
			set SubNode = objDOM.createElement("Activity")
				SubNode.setAttribute "ActCode",arrActCode(nCount)
				SubNode.setAttribute "ActName",arrActName(nCount)
				Root.appendChild SubNode
		next
	elseif sAppCode = 10 then ' Tax Deduction
	elseif sAppCode = 11 then ' Quality control
	elseif sAppCode = 12 then ' Engineering Services
	elseif sAppCode = 13 then ' Central Excise
	elseif sAppCode = 14 then ' Job Work
	elseif sAppCode = 15 then ' Canteen
	elseif sAppCode = 16 then ' VAT
	end if

	
	if sQuery<>"" then
		with drset
			.CursorLocation = 3
			.CursorType = 3
			.ActiveConnection = con
			.Source = sQuery
			.Open
		end with
		if not drSet.EOF then
		arrSize = drSet.RecordCount + 1
			Redim arrActCode(arrSize)
			Redim arrActName(arrSize)
		
			set SubNode = objDOM.createElement("Activity")
				SubNode.setAttribute "ActCode","0"
				SubNode.setAttribute "ActName","Select"
				Root.appendChild SubNode
			do while not drSet.EOF 
				set SubNode = objDOM.createElement("Activity")
				SubNode.setAttribute "ActCode",drset(0)
				SubNode.setAttribute "ActName",drset(1)
				Root.appendChild SubNode
				drSet.MoveNext 
			loop
		end if
		drSet.Close 
	end if 	'if sQuery<>"" then
	
	Response.ContentType ="text/xml"
	Response.Write Root.xml
%>