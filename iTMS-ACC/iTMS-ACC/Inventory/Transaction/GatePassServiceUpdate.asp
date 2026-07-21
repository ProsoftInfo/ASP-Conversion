<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GatePassServiceUpdate.asp
	'Module Name				:	INVENTORY (Transaction)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	APRIL 05,2010
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
<!-- #include File="../../include/purpopulate.asp" -->
<!--#include file="../../include/NoSeries.asp"-->
<%
	dim newxml
	dim dcrs,sSql,RootNode,ItemNode,HeaderNode,dcrs1
	dim iQty,sOtherDesc,sItemType,sOrgCode,sRemarks,iSup,sTransport,sTakenBy,sDeliveryBy
	dim sExp,iGPNo,iSeriesNo,iSeriesCode,iEntryNo,iCtr,sGatePassDate,sReason
	dim iItem, iClass,nItemValue
	Dim sInvoicedUoM,sDCNo,sApplFormJJ,SupplierNode
	Dim sFinPeriod, sFinPeriodFrom,sFinPeriodTo, sFinFrom,sFinTo
	Dim sDate, arrFin
	sFinPeriod = Session("FinPeriod")
	sFinPeriodFrom = FormatDate("04/01/" & Mid(sFinPeriod,1,4))
	sFinPeriodTo = FormatDate("03/31/" & Mid(sFinPeriod,6,4))
	sFinFrom = FormatDate("04/01/" & Mid(sFinPeriod,1,4))
	sFinTo = FormatDate("03/31/" & Mid(sFinPeriod,6,4))

	sDate = FormatDate(date())

	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	' Create our DOM Document Objects
	Set newxml = Server.CreateObject("Microsoft.XMLDOM")

	newxml.async = false
	newxml.load(Request)
	iGPNo = Request.QueryString("GatePassNo")

'	Response.Write newxml.xml

	Set RootNode = newxml.documentElement

	sExp ="//HEADER"
	Set HeaderNode = RootNode.selectSingleNode(sExp)
	sItemType = trim(HeaderNode.Attributes.getNamedItem("ITEMTYPE").Value)
	sOrgCode = trim(HeaderNode.Attributes.getNamedItem("FORUNIT").Value)
	sRemarks = PackQuote(trim(HeaderNode.Attributes.getNamedItem("REMARKS").Value))
	iSup = trim(HeaderNode.Attributes.getNamedItem("SUPPAGENT").Value)
	
	sTransport	= PackQuote(trim(HeaderNode.Attributes.getNamedItem("Transport").Value))
	sTakenBy	= PackQuote(trim(HeaderNode.Attributes.getNamedItem("TakenBy").Value))
	sDeliveryBy = PackQuote(trim(HeaderNode.Attributes.getNamedItem("DeliveryBy").Value))

		
	con.beginTrans
	
	sExp = "//Supplier"
	Set SupplierNode = RootNode.selectSingleNode(sExp)
	iSup = trim(SupplierNode.Attributes.getNamedItem("PartyCode").Value)
	
	sGatePassDate = sDate

	sSql = "UPDATE FORGATEPASSHEADER set ORGANISATIONCODE = "& Pack(sOrgCode) &",PARTYCODE="& iSup &",TYPEOFITEMS="& Pack(sItemType) &",MARKEDON = CONVERT(DATETIME,'" & sDate & "',103),REMARKS='"&sRemarks&"',Transport='"&  sTransport &"',TakenBy='"&  sTakenBy  &"',DeliveryBy='"&  sDeliveryBy &"' WHERE  GATEPASSNO = "& iGPNo 
	'Response.write sSql & "<BR>"
	con.Execute sSql
	
	 sSql ="Delete from  FORGATEPASSDETAILS where GATEPASSNO = " & iGPNo
	 'Response.Write sSql
	con.execute sSql
	
	sExp ="//DETAILS"
	Set ItemNode = RootNode.Selectnodes(sExp)
	if ItemNode.Length > 0 then
		For iCtr = 0 to ItemNode.Length - 1
			sOtherDesc = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("OTHERDESC").Value)
			iQty = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("QTY").Value)
			iItem = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ITEMCODE").Value)
			iClass = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("CLASSCODE").Value) 
			sInvoicedUoM = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("UOM").Value) 
			
			nItemValue = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("VALUE").Value) 
			sApplFormJJ = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("FORMJJ").Value) 
			sReason = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("REASON").Value)
			
						
			If iItem = "" Then iItem = "Null"
			If iClass = "" Then iClass = "Null"
			If Trim(sInvoicedUoM) = "" Then sInvoicedUoM = "Null"
			If Trim(nItemValue) = "" Then nItemValue = "Null"
			If sApplFormJJ = "" Then sApplFormJJ = "N"
			
			
			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(ENTRYNO)+1,1) FROM FORGATEPASSDETAILS WHERE GATEPASSNO = " & iGPNo & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing
			if not dcrs.EOF then
				iEntryNo = trim(dcrs(0))
			end if
			dcrs.Close
			
			sSql = "INSERT INTO FORGATEPASSDETAILS (GATEPASSNO,ENTRYNO,ITEMCODE," &_
				"CLASSIFICATIONCODE,QUANTITY,DESCRIPTION,INVOICEDUOM,ItemValue,FormJJ,Reason) VALUES " &_
				"(" & iGPNo & "," & iEntryNo & "," & iItem & "," & iClass & "," & iQty & "," &_
				"" & Pack(sOtherDesc) & ",'" & sInvoicedUoM & "'," & nItemValue & ",'" & sApplFormJJ  & "','"&trim(sReason)&"')"
'			Response.write sSql & "<BR>"
			con.Execute sSql
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

