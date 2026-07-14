<%@ Language="VBScript" %>
<% option explicit %>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	BANKVOUCHERDELETE.ASP
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	UMAMAHESWARI
	'Created On					:	April 21, 2011
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
<!--#include virtual="/include/Databaseconnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<%
	Dim rsObj,adoCmd
	Dim sQuery
	Dim iCrTransNo,iFromVouTransNo,iToVouTransNo,iToCrTransNo

	set rsObj = Server.CreateObject("ADODB.Recordset")
	iCrTransNo = Request.QueryString("Val")
	Response.Write "<p>iCrTransNo="&iCrTransNo

	con.beginTrans

	If iCrTransNo > 0 then

	    'sQuery = "Select TransactionNumber,ContraTransactionNumber from Acc_T_VoucherHeader where CreatedTransNo = "& iCrTransNo
	    'Response.Write "<P>"& sQuery
	    'With rsObj
	    '    .CursorLocation = 3
	    '    .CursorType = 3
	    '    .ActiveConnection = con
	    '    .Source = sQuery
	    '    .Open
 	    'End With
 	    'Response.Write "<p> Record Count = "&rsObj.RecordCount
	    'if not rsObj.EOF then
	    '    iFromVouTransNo = rsObj(0)
	    '    iToVouTransNo =  rsObj(1)
	    'end if
	    'rsObj.Close

    	'if Trim(iToVouTransNo)<>"0" and Trim(iToVouTransNo)<>"" then
	    '    sQuery = "Select CreatedTransNo from Acc_T_VoucherHeader where TransactionNumber= "& iToVouTransNo
	    '    Response.Write "<P>"& sQuery
	    '    rsObj.Open sQuery,con
	    '    if not rsObj.EOF then
	    '        iToCrTransNo = rsObj(0)
	    '    end if
	    '    rsObj.Close

	    '    sQuery = "Delete from Acc_T_CreatedVoucherDetails where CreatedTransNo = "& iToCrTransNo
	    '    Response.Write "<p>"& sQuery
	    '    con.execute sQuery

        '    sQuery = "Delete from Acc_T_CreatedVoucherHeader where CreatedTransNo = "& iToCrTransNo
	    '    Response.Write "<p>"& sQuery
	    '    con.execute sQuery

    	'    sQuery = "Delete from Acc_T_VoucherDetails where TransactionNumber = "& iToVouTransNo
	    '    Response.Write "<p>"& sQuery
	    '    con.execute sQuery

	    '    sQuery = "Delete from Acc_T_VoucherHeader where TransactionNumber = "& iToVouTransNo
	    '    Response.Write "<p>"& sQuery
	    '    con.execute sQuery

	    '    sQuery = "Delete from Acc_T_GLTransSummaryBreakup where VoucherNo = "& iFromVouTransNo
	    '    Response.Write "<p>"& sQuery
	    '    con.execute sQuery

	    '    sQuery = "Delete from Acc_T_GLTransSummaryBreakup where VoucherNo = "& iToVouTransNo
	    '    Response.Write "<p>"& sQuery
	    '    con.execute sQuery

	    '    sQuery = "Delete from Acc_T_GLTransactions where TransactionNumber = "& iToVouTransNo
	    '    Response.Write "<p>"& sQuery
	    '    con.execute sQuery
	    'end if 'if Trim(iToVouTransNo)<>"0" and Trim(iToVouTransNo)<>"" then

		set adoCmd = Server.CreateObject("ADODB.Command")
		set adoCmd.ActiveConnection = con

		'adoCmd.CommandText = "SP_BankVoucherDeleteAfterApproval"
		adoCmd.CommandText = "Proc_ACC_ACCCashBankVoucherDel"
		adoCmd.CommandType = 4 'adCmdStoredProc

		adoCmd.Parameters.Append adoCmd.CreateParameter("@CreatedTransNo",3,1,200,iCrTransNo)

		adoCmd.Execute()
	End IF

	'Con.rollbacktrans
	'Response.End
	con.CommitTrans


	con.close
	Set con = Nothing
	Response.Redirect "MANAGEBANKVOUCHERS.ASP"
%>
