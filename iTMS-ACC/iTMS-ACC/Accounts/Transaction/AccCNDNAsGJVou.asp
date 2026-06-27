<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AccCNGJVoucGenerate.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	S.MAHESWARI
	'Created On					:	JUL 21, 2008
	'Modified On				:

%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/NoSeries.asp"-->
<%

'XML DOM Variables

Dim objRs,iBookNo,iCrTransNo,iSeriesNo,iSeriesCode,sOrgID,sVoucDate,adoCmd
Dim sQuery,iVouNo,sPara,sFormVal,iBookCode
Dim dDebTotal,dCreTotal,iTransNo
Set objRs = Server.CreateObject("ADODB.RecordSet")
sPara = Request("Para")
sFormVal = "G"

IF CStr(sPara) = "" Then
	sPara = "C"
End IF

iBookNo = Request("BookNo")
iCrTransNo = Request("hTransNo")
iBookCode = Request("BookCode")
sQuery = "Select OUDefinitionID,Convert(Varchar,VoucherDate,103) VoucherDate from  "&_
		 "Acc_T_CreatedVoucherHeader where CreatedTransNo = "& iCrTransNo 

objRs.Open sQuery,Con
If not objRs.EOF then
	sOrgId		= objRs("OUDefinitionID")
	sVoucDate	= objRs("VoucherDate")  
end If
objRs.Close 

sQuery="select DrSeriesNo,DrSeriesCode from Acc_M_BookNumberSeries where "&_
		"OUDefinitionID='"&sOrgId&"' and BookCode='"&Trim(iBookCode)&"' and BookNumber= "&iBookNo
Response.Write sQuery
objRs.open sQuery,con
if not objRs.EOF then
	iSeriesNo=objRs(0)
	iSeriesCode=objRs(1)
end if	
objRs.close() 
Response.Write " iSeriesNo= "& iSeriesNo
Response.Write " iSeriesCode = "& iSeriesCode

Con.beginTrans
iVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)

Response.Write "<br>"
Response.Write "iCrTransNo="& iCrTransNo  & "<br>"
Response.Write "iVouNo="& iVouNo & "<br>"

sQuery = "AccountDnCNGJVou"
Set adoCmd = Server.CreateObject("ADODB.Command")
Set adoCmd.ActiveConnection =con
adoCmd.CommandText = sQuery
adoCmd.CommandType = 4 'adCmdStoredProc
adoCmd.Parameters.Append adoCmd.CreateParameter("@iCrTransNo",3,1,3,iCrTransNo)
adoCmd.Parameters.Append adoCmd.CreateParameter("@AccVouNo",201,1,30,iVouNo)
adoCmd.Execute()


sQuery = " Select TransactionNumber from ACC_T_VoucherHeader where CreatedTransNo = "& iCrTransNo 
Response.Write "<p>"& sQuery
objRs.Open sQuery,con
if not objRs.EOF then
    iTransNo = objRs(0)
end if
objrs.Close 

sQuery = "Select SUM(Amount) from Acc_T_VoucherDetails where TransactionNumber = "& iTransNo  &" and TransCRDRIndication = 'D'"
Response.Write "<p>"& sQuery
objRs.Open sQuery,con
if not objRs.EOF then
    dDebTotal = objRs(0)
end if
objRs.Close 

sQuery = " Select SUM(Amount) from Acc_T_VoucherDetails where TransactionNumber = "& iTransNo  &" and TransCRDRIndication = 'C'"
Response.Write "<p>"& sQuery
objRs.Open sQuery,con
if not objRs.EOF then
    dCreTotal = objRs(0)
end if
objRs.Close 
Response.Write "<p>DebTotal="& dDebTotal 
Response.Write "<p>CreTotal ="& dCreTotal

dDebTotal = FormatNumber(dDebTotal,2,0,0,0)
dCreTotal = FormatNumber(dCreTotal,2,0,0,0)
Response.Write "<p>DebTotal="& dDebTotal 
Response.Write "<p>CreTotal ="& dCreTotal
if CDbl(dDebTotal)<> CDbl(dCreTotal) then
    'Response.Clear
    Response.Write "<P><b>Debit Total = "& dDebTotal 
    Response.Write "<P><b>Credit Total = "& dCreTotal 
    Response.Write "<P><b>Debit and Credit Does Not Match"
    con.rollbacktrans
    Response.End 
end if

IF con.Errors.count <>0 THEN
	con.RollbackTrans
	FOR iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	NEXT
ELSE
   ' con.rollbacktrans
   ' Response.End 
    
    Response.Clear 
    con.CommitTrans
    
	
	IF trim(sPara) = "C" then
		Response.Redirect ("CREDITVOUCHERS.ASP?hFormVal="&sFormVal)
	ElseIF  trim(sPara) = "D" then
		Response.Redirect ("DEBITVOUCHERS.ASP?hFormVal="&sFormVal)
	End IF
End IF
%>

