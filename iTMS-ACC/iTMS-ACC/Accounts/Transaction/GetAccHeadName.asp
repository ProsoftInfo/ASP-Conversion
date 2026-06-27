
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GetAccHeadName.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Ragavendran R
	'Created On					:	Feburary 11, 2011
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
<%
Dim sOrgID,iInvType,objRs,sMod,iAccHead,sAccHdName,sQuery,iBookNo

set objRs = server.CreateObject("ADODB.Recordset")
sOrgID  = session("organizationcode")
iInvType  = Request.QueryString("InvType")
sMod    = Request.QueryString("Mod")
iBookNo = Request.QueryString("BookNo")
Set objRs = Server.CreateObject("ADODB.RecordSet")

if trim(sMod)="PUR" then
    sQuery = "Select M.AccountHead From Acc_M_GLAccountHead M, "&_
	    	 "VwOrgBookNames V Where V.BookCode = '04' and V.BookNumber = "&iBookNo&" "&_
		     "and V.OUDefinitionID = '"&sOrgId&"' and V.BookAccountHead = M.AccountHead "
else
    sQuery = "Select AccountHead From App_R_OrgnTaxAccountHead Where TaxCode is Null and TaxCategoryCode  "&_
		     "is Null and InvoiceType = "&iInvType&" and OUDefinitionID = '"&sOrgId&"'  "
end if		     
'Response.Write sQuery
    objRs.Open sQuery,Con

    IF Not objRs.EOF Then
	    iAccHead = objRs(0)
    Else
	    iAccHead = 0
    End IF
    objRs.Close
    
if trim(iAccHead)<>"" then
    sQuery = "Select AccountDescription From Acc_M_GLAccountHead Where AccountHead = "&iAccHead&" "
    objRs.Open sQuery,Con
    IF Not objRs.EOF Then
	    sAccHdName = objRs(0)
    Else
	    sAccHdName = ""
    End IF
    objRs.Close
end if    
if trim(sAccHdName)<>"" and trim(iAccHead)<>"" then
    response.write iAccHead&":"&sAccHdName
end if'if trim(sAccHdName)<>"" and trim(iAccHead)<>"" then
%>