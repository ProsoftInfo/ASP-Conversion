<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	AppVouStatusUpdateAll.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	Sre Hari.M
	'Created On					:	Feb 20, 2006
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
<!--#include file="../../include/Accpopulate.asp"-->
<%
	dim iTransNo,sVouStatus,sQuery,i,sTransNo,sBookCode,Objrs,sFormVal,sSelVouTy

	set Objrs=server.CreateObject("ADODB.Recordset")
	iTransNo=Request("hTransNo")
	sTransNo=Split(iTransNo,":")
	sFormVal = Request("hFormVal")
	sSelVouTy = Request("voutype")
	

	sVouStatus="010103"

	sQuery="select isnull(bookcode,0) from Acc_T_CreatedVoucherHeader where CreatedTransNo="&sTransNo(0)
	Objrs.Open sQuery,con
		sBookCode=Objrs(0)
	Objrs.close

		
	for i=0 to UBound(sTransNo)
		sQuery="update Acc_T_CreatedVoucherHeader set CreatedVouchStatus='"&sVouStatus&"'" &_
		" where CreatedTransNo="&sTransNo(i)
		con.execute(sQuery)

		sQuery="insert into Acc_T_VoucherApprovalTracking (TransactionNumber,CreatedTransNo,ApproverLevel,"&_
			"ApprovedBy,ApprovedOn) values (NULL,"&sTransNo(i)&",'1',"&getUserId&",getdate())"
		con.execute(sQuery)

		sQuery="delete Acc_T_VouchersForApproval where CreatedTransNo="&sTransNo(i)
		con.execute(sQuery)
	next

	if sBookCode="01" then
		'Response.Redirect("CashVouchers.asp?hFormVal="&sFormVal&"&voutype="&sSelVouTy)
		Response.Redirect("CashVouchers.asp?hFormVal="&sFormVal&"&voutype="&sSelVouTy&"&ACTN="&Session("ACTN"))
	elseif sBookCode = "02" then
		Response.Redirect("BankVouchers.asp?hFormVal="&sFormVal&"&voutype="&sSelVouTy&"&ACTN="&Session("ACTN"))
	elseif sBookCode = "08" then
		Response.Redirect("GJVOUCHERS.ASP?hFormVal="&sFormVal&"&voutype="&sSelVouTy)	
	elseif sBookCode = "04" then
		Response.Redirect("PURCHASEVOUCHERS.ASP?hFormVal="&sFormVal&"&voutype="&sSelVouTy)
	elseif sBookCode = "05" then
		Response.Redirect("SALESVOUCHERS.ASP?hFormVal="&sFormVal&"&voutype="&sSelVouTy)
	elseif sBookCode = "06" then
		Response.Redirect("DebitVouchers.ASP?hFormVal="&sFormVal&"&voutype="&sSelVouTy)	
	elseif sBookCode = "07" then
		Response.Redirect("CreditVouchers.ASP?hFormVal="&sFormVal&"&voutype="&sSelVouTy)		
	end if

%>