

<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLGetTaxDetails.asp
	'Module Name				:	Purchase (Transactions-Invoice)
	'Author Name				:	
	'Created On					:	November 16 , 2005
	'Modified By				:	Ragavendran R
	'Modified On				:	Jun 14,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:	InvPurInvoiceEntry.asp
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
<!--#include virtual="/include/sessionVerify.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/purpopulate.asp"-->

<%
Dim iToBeAccounted,iTRndoff, sTaxCatType,sTotpackvalue,sOrghdno,sSumpack,sPackvalue
Dim nDisplayTotal,dTotal
Dim objRs,objRs2,sSql,sSql1,sOrgID,sPurType,rsTemp
dim oDom,Root,oNodTaxRoot,newElem,newElem1
dim sTaxName,sCatCode,sTaxCode,sFormula,dTaxValue,sAccHead,iCrEligible,iRegCode
dim staxmode,sInvNumber,sRcptNumber,sPackingCode,sNoofPack

'dTotal=10
sOrgID=Request.QueryString("ForUnit")
sPurType=Request.QueryString("PurType")
sInvNumber = Request.QueryString("InvNo")
sRcptNumber = Request.QueryString("RcptNum") 
if trim(sPurType) = "" then sPurType = 0
'Response.Write"sPurType="&sPurType&"<BR>"
Set rsTemp = Server.CreateObject("Adodb.recordset")
Set objRs = Server.CreateObject("Adodb.recordset")
Set objRs2 = Server.CreateObject("Adodb.recordset")

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

Set Root = oDOM.createElement("Root")
oDOM.appendChild Root

nDisplayTotal =  dTotal

Set oNodTaxRoot = oDOM.createElement("TaxDetails")
oNodTaxRoot.setAttribute "Basicvalue",dTotal
oNodTaxRoot.setAttribute "NettValue",dTotal
'oNodTaxRoot.setAttribute "InvValue","0"
'oNodTaxRoot.setAttribute "Roundoff","0"
oNodTaxRoot.setAttribute "TotalTax","0"
oNodTaxRoot.setAttribute "SubTotal","0"
oNodTaxRoot.setAttribute "PurchaseType",sPurType
Root.appendChild oNodTaxRoot

sSql = "Select TaxShortName,TaxCategoryCode,TaxCode,ComputationMode,isnull(SumOfFields,''),isnull(FlatAmount,0),"&_
	" AccountHead,TaxCreditEligibility,isnull(RegisterCode,0),AccountTaxAccHead,isnull(Roundoff,'0'),TaxCategoryType, isnull(orgTaxAccHdNo,0) as OrgTaxAccHdNo from VwPurchaseTaxDetails where OUDefinitionID='"&sOrgID&"' and PurchaseType="&sPurType& " order by TaxHierarchy"

'Response.Write "sSql ="+ sSql

With objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSql
	.ActiveConnection = con
	.Open
End with
Set objRs.ActiveConnection = nothing

Set sTaxName=objRs(0)
Set sCatCode=objRs(1)
Set sTaxCode=objRs(2)
Set sTaxMode=objRs(3)
Set sFormula=objRs(4)
Set dTaxValue=objRs(5)
Set sAccHead=objRs(6)
Set iCrEligible=objRs(7)
Set iRegCode=objRs(8)
Set iToBeAccounted = objRs(9)
Set iTRndoff = objrs(10)
Set sTaxCatType  = objrs(11)
Set sOrghdno = objRs(12)

Do while not objRs.EOF
	Set newElem = oDOM.createElement("Tax")
	newElem.setAttribute "CatCode",sCatCode
	newElem.setAttribute "TaxCode",sTaxCode
	newElem.setAttribute "TaxMode",sTaxMode
	newElem.setAttribute "TaxFormula",sFormula
	newElem.setAttribute "TaxValue",CStr(dTaxValue)
	newElem.setAttribute "TaxAmount","0"
	newElem.setAttribute "AccHead",sAccHead
	newElem.setAttribute "CrEligible",iCrEligible
	newElem.setAttribute "RegisterCode",iRegCode
	newElem.setAttribute "ToBeAccounted",iToBeAccounted
	newElem.SetAttribute "Rndoff",iTRndoff
	newElem.SetAttribute "TaxCategoryType",sTaxCatType
	newElem.setAttribute "OrgTaxAccHdNo",sOrghdno
	newElem.Text= sTaxName
	oNodTaxRoot.appendChild newElem
	
	''Added by Ragav for getting Tax Packing Details
	''Begin
	If objrs("ComputationMode") = "K" then
	
			
			if trim(sRcptNumber)<>"" then
				sSql = "Select isNull(SUM(NoofPackage),0),PackingCode from RCV_T_ActualRcptItemLot where ReceiptNumber = "& sRcptNumber &" Group by PackingCode"
			'	Response.Write sSql 
				rsTemp.Open sSql,con
				if not rsTemp.EOF then
					sSumpack = rsTemp(0)
					sPackingCode = rsTemp(1)
				else
					sSumpack = 0
					sPackingCode = 0
				end if
				rsTemp.Close 
			end if 'if trim(sRcptNumber)<>"" then
			if trim(sSumpack)="" or IsNull(sSumpack) then
				sSumpack =0
			end if
			
			
			sSql  = "Select Distinct PackingCode,RatePerPack From RCV_T_InvoiceTaxPacDet where OrgTaxAccHdNo = "& objrs("OrgTaxAccHdNo") &" and InvoiceNumber = " & sInvNumber  & ""
			'Response.Write sSql 
			rsTemp.Open sSql,con
			
			If rsTemp.EOF then
				rsTemp.Close 
				sSql = "Select Distinct PackingCode,RatePerPack From APP_R_PurOrgnTaxPackDetails where OrgTaxAccHdNo = "& objrs("OrgTaxAccHdNo") &" "
				rsTemp.Open sSql,con
			end if
			
			If not rsTemp.EOF then

				Do while not rsTemp.EOF
					Set newElem1 = oDOM.createElement("Taxpack")
					newElem1.setAttribute "PackCode",rsTemp(0)
					newElem1.setAttribute "Packrate",rsTemp(1)
					if StrComp(sPackingCode,rsTemp(0))=0 then
						sNoofPack = sSumpack 
					else
						sNoofPack = 0
					end if
					newElem1.setAttribute "Noofpack",sNoofPack
					sPackvalue = cdbl(sNoofPack) * cdbl(rsTemp(1))
					newElem1.setAttribute "Packvalue",sPackvalue
					newElem.appendchild Newelem1
					rsTemp.MoveNext
				loop
			End if
			rsTemp.Close
	end if 'If objrs("ComputationMode") = "K" then
	''End
	
	objRs.MoveNext
Loop
objRs.Close
'oDOM.Save server.MapPath("../temp/transaction/yyy.xml")

Response.ContentType="text/xml"
Response.Write oDOM.xml

%>
<%
Function Packsum(Objroot,Packcode)
Dim sTempnode,sExp,Qty,ICtr,sTot
sExp = "//Pack[@PackCode = "&Packcode&"]"
Set sTempnode = Objroot.selectnodes(sExp)
if sTempnode.Length > 0 then
	sTot = sTempnode.Length
Else
	sTot  = 0
End if
Packsum = sTot
End function
%>