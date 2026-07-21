<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	InvGetItemDetForRefType.asp
	'Module Name				:	Inventory (MRS Issue)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Dec 27,2010
	'							:
	'Connects To				:	DirectIssueItemEntry.asp
	'Internal Variables			:
	'Database					:
	'Queries Used				:
	'Counters					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!--#include file="../../include/IncludeDatePicker.asp"-->
<%
	Dim	rsItem,objDOM,rsStore
	Dim ndRoot,ndItem,ndSubcontract,ndDetails
	Dim sOrgCode,sItemType,sStock,sRefCodes,sQuery,sRefType
	Dim iEntryNo
	Set rsItem	=	Server.CreateObject("ADODB.Recordset")
	set rsStore =   Server.CreateObject("ADODB.Recordset")
	Set objDOM	=	Server.CreateObject("Microsoft.XMLDOM")

	sOrgCode = Request.QueryString("orgID")
	sItemType= Request.QueryString("sIType")
	sStock = Request.QueryString("Stock")
	sRefCodes = Request.QueryString("RefCodes")
	sRefCodes = Replace(sRefCodes,",","','")
	sRefType = Request.QueryString("RefType")
	set ndRoot = objDOM.createElement("Root")
	objDOM.appendChild(ndRoot)

	'Response.Write "Reftype = "& sRefType
	if sRefType = "11" then 'Material Requisition
		sQuery = "Select V.CompanyItemCode,(V.ItemDescription + '--'+ isnull(V.CatalogueNo,'N/A')),"&_
			     " V.GroupName,V.StoresUOM,V.ItemCode,V.ClassificationCode,V.DecimalAllowed,V.ReceiptNumbering,"&_
			     " isNull(M.ItemAttributes,''),V.ItemTypeID,MRSNumber,QuantityApproved from Inv_T_MRSItemDetails M,VWITEM V where MRSNumber in ("& sRefCodes &")"&_
			     " and V.ItemCode=M.ITemCode and V.ClassificationCode = M.ClassificationCode and M.OrganisationCode = V.OrganisationCode"&_
			     " and M.OrganisationCode = '"& sOrgCode &"'"
    elseif sRefType = "14" then ' Mix Code
        sQuery = " Select CompanyItemCode,ItemDescription,GroupName,StoresUoM,ItemCode,ClassificationCode,"&_
                 " DecimalAllowed,ReceiptNumbering,isNull(Attributelist,''),ItemTypeID,MixCode,0 from VW_PRD_MixVareity"&_
                 " Where OrganisationCode = '"& sOrgCode &"' and MixCode in ('"& sRefCodes &"')"
    elseif sRefType = "15" or sRefType = "42" then ' Sales Order, Sales Order - Jobwork
        sQuery = " Select V.CompanyItemCode,V.ItemDescription,V.GroupName,V.StoresUoM,V.ItemCode,"&_
                 " V.ClassificationCode,V.DecimalAllowed,V.ReceiptNumbering,'',V.ItemTypeID,OrderNumber,QuantityOrdered "&_
                 " from Sal_T_OrdersDetails M,VwItem V where OrderNumber in("& sRefCodes &") and V.ItemCode = M.ItemCode"&_
                 " and V.ClassificationCode=M.ClassificationCode and V.OrganisationCode = '"& sOrgCode &"'"
    elseif sRefType = "17" then ' Production Order
        sQuery = " Select V.CompanyItemCode,V.ItemDescription,V.GroupName,V.StoresUoM,V.ItemCode,"&_
                 " V.ClassificationCode,V.DecimalAllowed,V.ReceiptNumbering,isNull(AttributeList,''),"&_
                 " V.ItemTypeID,ProductionOrderNo,RequiredQuantity from PRD_T_ProductDetails P,VWITEM V where V.ItemCode=P.ItemCode and "&_
                 " V.ClassificationCode = P.ClassificationCode and ProductionOrderNo in ('"& sRefCodes &"') and V.OrganisationCode = '"& sOrgCode &"'"
    elseif sRefType ="36" then 'Actual Receipt for Purchase return
        sQuery = "Select V.CompanyItemCode,V.ItemDescription,V.GroupName,V.StoresUOM,V.ItemCode, "&_
                 " V.ClassificationCode,V.DecimalAllowed,V.ReceiptNumbering,isNull(H.ItemAttributes,''),"&_
                 " V.ItemTypeID,H.ReceiptNumber,QuantityReceived from RCV_T_ActualRcptItemDet H, VW_Purchase_ActionOnRcptItem A,VWITEM V where "&_
                 " H.ReceiptNumber = A.receiptNumber and V.ItemCode = H.ItemCode and V.ClassificationCode = H.ClassificationCode "&_
                 " and V.OrganisationCode = H.OrganisationCode and A.ReceiptNumber in ("& sRefCodes  &") and H.OrganisationCode = '"& sOrgCode &"'"
        sQuery = " Select V.CompanyItemCode,V.ItemDescription,V.GroupName,V.StoresUOM,V.ItemCode, "&_
                 " V.ClassificationCode,V.DecimalAllowed,V.ReceiptNumbering,isNull(D.ItemAttributes,''),'',"&_
                 " H.ReceiptNumber,H.ActionOnQty from RCV_T_ActualRcptItemDet D join VWItem V on V.ItemCode = D.ItemCode join "&_
                 " VW_Purchase_ActionOnRcptItem H on D.ReceiptNumber = H.ReceiptNumber and D.EntryNo=H.EntryNo "&_
                 " and H.ReceiptNumber in ("& sRefCodes &") and D.OrganisationCode = '"& sOrgCode &"'"
    elseif sRefType ="22" then 'Purchase Order - Subcontract
    	sQuery = ""
	end if

	'Response.Write sQuery
	rsItem.Open sQuery,con
	if not rsItem.EOF then
		iEntryNo = 0
		do while not rsItem.EOF
			ndRoot.setAttribute "ItemType",rsItem(9)
			iEntryNo = iEntryNo + 1
			set ndItem = objDOM.createElement("Item")
				ndItem.setAttribute "EntryNo",iEntryNo
				ndItem.setAttribute "CompanyItemCode",rsItem(0)
				ndItem.setAttribute "ItemCode",rsItem(4)
				ndItem.setAttribute "ClassCode",rsItem(5)
				ndItem.setAttribute "ItemName",rsItem(1)
				ndItem.setAttribute "ClassName",rsItem(2)
				ndItem.setAttribute "StoresUoM",rsItem(3)
				ndItem.setAttribute "Decimal",rsItem(6)
				ndItem.setAttribute "ReceiptNum",rsItem(7)
				ndItem.setAttribute "AttributeList",rsItem(8)
				ndItem.setAttribute "No",rsItem(10)
				ndItem.setAttribute "Qty",rsItem(11)
			ndRoot.appendChild(ndItem)
			rsItem.MoveNext
		loop
	end if
	rsItem.Close

	if Trim(sRefType)="11" then
	    sQuery =" Select H.SubConProcessID,Instruction,SubConLabourCharge,HardWaste,IntWaste,SubconProcessName,PurchaseOrderNo from PUR_T_POHeader H join App_M_SubContractProcess P on H.SubConProcessID=P.SubconProcessID where PurchaseOrderNo in (Select AppRefNo from INV_T_MRSHeader where MRSNUmber =  "& sRefCodes &")"
        rsItem.Open sQuery,con
        if not rsItem.EOF then
            set ndSubcontract = objDOM.createElement("SubContract")
            ndSubcontract.setAttribute "SCProcess",rsItem(0)
            ndSubcontract.setAttribute "Instruct",rsItem(1)
            ndSubcontract.setAttribute "LabourCharge",rsItem(2)
            ndSubcontract.setAttribute "Currency","INR"
            ndSubcontract.setAttribute "HardWaste",rsItem(3)
            ndSubcontract.setAttribute "InvWaste",rsItem(4)
            ndSubcontract.setAttribute "ProcessName",rsItem(5)
            ndRoot.appendChild ndSubcontract

            sQuery = "Select ItemReceivedAs,ClassReceivedAs,V.ItemDescription from PUR_T_POItemAdditionalDetail D join VWITEM V on D.ItemReceivedas=V.ItemCode where PurchaseOrderNo = "& rsItem(6)
            'Response.Write sQuery
            rsStore.Open sQuery,con
            if not rsStore.EOF then
                do while not rsStore.EOF
                    set ndDetails = objDOM.createElement("Details")
                    ndDetails.setAttribute "MatRecdAsItem",rsStore(0)
                    ndDetails.setAttribute "MatRecdAsCode",rsStore(1)
                    ndDetails.setAttribute "MatRecdAsDescr",rsStore(2)
                    ndSubcontract.appendChild ndDetails
                    rsStore.MoveNext
                loop
            end if
            rsStore.Close
        end if
        rsItem.Close
	end if  'if Trim(sRefType)="11" then

	Response.ContentType = "text/xml"
	Response.Write objDOM.xml
%>
