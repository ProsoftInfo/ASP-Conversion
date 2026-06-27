<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AccVoucherNo_Generate.asp
	'Module Name				:	Accounts 
	'Author Name				:	Kalaiselvi R
	'Created On					:	June 28,2010
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	
	'							:
	'Connects To				:	
	'Procedures/Functions Used	:	
	'Internal Variables			:	
	
	'Database					:	ITMS
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
	'Declare Variables
	Dim sUnit,sBookCode,sBookNo,sVoucherType,sTransactionType,sFinPeriod
	Dim dtFromDate,sGetBookCode
	
	Dim adoCmd
	
	sUnit		= Request.Form("hUnitNo")
	sGetBookCode	= Request.QueryString("BookCode")
	sTransactionType = Request.Form("hVocType")
	dtFromDate = Request.Form("hFromDate")
	
	sFinPeriod = Session("FinPeriod")
	
	If sGetBookCode = "01" Then
		if sTransactionType ="C" then sTransactionType = "CAR"
		if sTransactionType ="D" then sTransactionType = "CAP"
		if sTransactionType = "CAP" or sTransactionType = "CAR" then
			sBookCode= "01"
		end if 
	Else
		if sTransactionType ="C" then sTransactionType = "BAR"
		if sTransactionType ="D" then sTransactionType = "BAP"
		if sTransactionType = "BAP" or sTransactionType = "BAR" then
			sBookCode= "02"
		end if 
	End IF
		
	Con.BeginTrans
	
	
	''-----------------------------------
	' calling Stored Procedure 
								
	Set adoCmd = Server.CreateObject("ADODB.Command")
	Set adoCmd.ActiveConnection = con
	
	If sGetBookCode = "01" Then
		adoCmd.CommandText = "Proc_Acc_CashVoucherNo_Generate"
	Else
		adoCmd.CommandText = "Proc_Acc_BankVoucherNo_Generate"
	End IF
	adoCmd.CommandType = 4 'adCmdStoredProc
	
	
	'Response.Write "<p> sunit = " & sUnit 
	'Response.Write "<p> sBookCode = " & sBookCode & "----" & sGetBookCode 
	'Response.Write "<p> sTransactionType = " & sTransactionType 
	'Response.Write "<p> dtFromDate = " & dtFromDate 
	'Response.Write "<p> sFinPeriod = " & sFinPeriod 
	'Response.Write "<p>DAta="&Session("ACTN")
	'Response.End 

 
	adoCmd.Parameters.Append adoCmd.CreateParameter("@Unit", 129  ,1  , 250 ,sUnit)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@BookCode", 129 ,1 ,5 ,sBookCode)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@TransactionType", 129 ,1 ,5 ,sTransactionType)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@FromDate", 129 ,1 ,10 ,dtFromDate) 
	adoCmd.Parameters.Append adoCmd.CreateParameter("@FinPeriod", 129 ,1 ,10 ,sFinPeriod)
									
		
	adoCmd.Execute()
	''----------------------------------
	'con.RollbackTrans
	'Response.End 
	
	Con.CommitTrans
	
	'Response.Write "<p>Over" 
	
	
	Response.Clear
	If sGetBookCode = "01" Then
		Response.Redirect "CASHVOUCHERS.ASP?ACTN="&Session("ACTN")
	Else
		Response.Redirect "BankVOUCHERS.ASP?ACTN="&Session("ACTN")
	End IF
	
%>
