<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	prn_Cheque.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	K.P.Sanjai Kumar
	'Created On					:	April 10,2004
	'Modified By				:	Manohar Prabhu .R
	'Modified On				:	April 22, 2004
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


	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/GetOrganization.asp"-->
<!--#include virtual="/include/PrintFunctions.asp"-->

<%
'------------------------Declaration Constants -----------------------------
dim sPrintHeader, aiHeaderColWidth(5,5), iVoucherNumber
'++++++++++++++++++ Declaration of variables used in page +++++++++++++++++'
dim isNo, i, j, temp, iPageNo
dim objRs, objRs1, objRs2
dim sQuery, sPartyName, sPartyAddress1, sPartyAddress2, sPartyCity, sPartyPin

'-----------------------------Declaration for Printing----------------------
dim objFSO,objTxt, iRecCount
dim iPgNo,iLineNo,stStr,iTotal

'---FILE-FILE-FILE------Intialization of file Object-----FILE-FILE-FILE----
set objFSO = Server.CreateObject("Scripting.FileSystemObject")
'set objTxt = objFSO.CreateTextFile(server.MapPath("../temp/Reports/"&Session.SessionID&"_Inventory_doc.txt"))
set objTxt = objFSO.CreateTextFile(server.MapPath("../temp/Transaction/cheque.txt"))
'----------------------End of file Object Intialization ------------------
iPgNo=1
iLineNo=0
iSno=1
set objRs  = server.CreateObject("adodb.recordset")
set objRs1 = server.CreateObject("adodb.recordset")
set objRs2 = server.CreateObject("adodb.recordset")
'++++++++++++++++++++++++ array declarations START +++++++++++++++++++++++++

'The address and cheque no. date
aiHeaderColWidth(0,0) = 2
aiHeaderColWidth(0,1) = 40
aiHeaderColWidth(0,2) = 5
aiHeaderColWidth(0,3) = 15
aiHeaderColWidth(0,4) = 3
aiHeaderColWidth(0,5) = 11

'Bill details
aiHeaderColWidth(1,0) = 3
aiHeaderColWidth(1,1) = 2
aiHeaderColWidth(1,2) = 8
aiHeaderColWidth(1,3) = 27
aiHeaderColWidth(1,4) = 11
aiHeaderColWidth(1,5) = 2

'Under rupees and date   ** checque leaf rows **
aiHeaderColWidth(2,0) = 37
aiHeaderColWidth(2,1) = 18
aiHeaderColWidth(2,2) = 15

'Pay 
aiHeaderColWidth(3,0) = 5
aiHeaderColWidth(3,1) = 70

'Rupees 
aiHeaderColWidth(4,0) = 10
aiHeaderColWidth(4,1) = 65

'Rupees in numbers
aiHeaderColWidth(5,0) = 54
aiHeaderColWidth(5,1) = 14
'++++++++++++++++++++++++ array declarations END +++++++++++++++++++++++++'
iVoucherNumber = Request.QueryString("Value")
'iTotal = Request.QueryString("Total")

'Response.Write iTotal &"<br><br><br>"

dim sTransactionNo, sOUdefinitionId, sChequeNo, sChequeDt, sVoucherNo, sVoucherDt
dim sPayablesNo, sPaidByTransactionNo, sPaidOn, sAmountPaid
dim sPartyBillNumber, sPartyBillDate, sTransactionNumber
dim sLinkRef
dim sTotalAmountPaid,sPayTo
sTotalAmountPaid = 0
set objRs  = server.CreateObject("adodb.recordset")

sQuery = "SELECT TransactionNumber, OUdefinitionId, PayToRecdFrom, BankInstrumentNo, convert(char,BankInstrumentDate,103), VoucherNumber, convert(char,VoucherDate,103), VoucherAmount  FROM ACC_T_VOUCHERHEADER WHERE CreatedTransNo= "&iVoucherNumber&" "

Response.Write sQuery
With objRs
	.CursorLocation=3
	.CursorType=3
	.ActiveConnection=con
	.Source=sQuery
	.Open
End With
Set objRs.ActiveConnection = Nothing
if not objRs.EOF Then
	sTransactionNo	=	objRs(0)
	sOUdefinitionId	=	right(objRs(1),2)
	sPartyName		=	objRs(2)
	sChequeNo		=	objRs(3)
	sChequeDt		=	objRs(4)
	sVoucherNo		=	objRs(5)
	sVoucherDt		=	objRs(6)
	iTotal			=	objRs(7)
end if
objRs.Close

iVoucherNumber = sTransactionNo
sPayTo = sPartyName

sQuery = "SELECT PayablesNumber, PaidByTransactionNo, convert(char, PaidOn, 103), AmountPaid FROM Acc_T_PybleAdjustmentDetails WHERE PaidByTransactionNo = " & sTransactionNo & " "

With objRs
	.CursorLocation=3
	.CursorType=3
	.ActiveConnection=con
	.Source=sQuery
	.Open
End With
if not objRs.EOF Then
	sPayablesNo	=	objRs(0)
	sPaidByTransactionNo	=	objRs(1)
	sPaidOn		=	objRs(2)
	sAmountPaid	=	cdbl( objRs(3) )
Else
	sPayablesNo	=	0
	sPaidByTransactionNo	=	0
	sPaidOn		=	0
	sAmountPaid	=	0
end if
objRs.Close


sQuery = "SELECT PartyBillNumber, convert(char,PartyBillDate,103), TransactionNumber FROM Acc_T_Payables where PayablesNumber = " & sPayablesNo &""

With objRs
	.CursorLocation=3
	.CursorType=3
	.ActiveConnection=con
	.Source=sQuery
	.Open
End With
if not objRs.EOF Then
	sPartyBillNumber	=	objRs(0)
	sPartyBillDate		=	objRs(1)
	sTransactionNumber	=	objRs(2)
end if

objRs.Close


sQuery = "SELECT VoucherNumber FROM ACC_T_VOUCHERHEADER WHERE TransactionNumber = " & iVoucherNumber &" "

Response.Write sQuery
With objRs
	.CursorLocation=3
	.CursorType=3
	.ActiveConnection=con
	.Source=sQuery
	.Open
End With
if not objRs.EOF Then
	sLinkRef	=	objRs(0)
end if

objRs.Close

'++++++++++++++++++ Definitions of variables used in page +++++++++++++++++'

'++++++++++++++++++ QUERIES FOR THE CONTENT +++++++++++++++++'
dim tempvar, xUnitAddress

'sQuery = "select PartyName, AddressLine1, AddressLine2, City, Pincode from vwsupplieraddress "
sQuery = "SELECT PayToRecdFrom, BankInstrumentNo, convert(char,BankInstrumentDate,103), VoucherNumber, convert(char,VoucherDate,103)  FROM ACC_T_VOUCHERHEADER WHERE TransactionType='BAP'"
objRs.Open sQuery, con
if not objRs.EOF Then
	tempvar = trim(objRs(0))
end if
		sPartyName = objRs(0)
		sPartyAddress1 = objRs(1)
		sPartyAddress2 = objRs(2)
		sPartyCity = objRs(3)
		sPartyPin = objRs(4)
objRs.Close

sQuery = "Select P.PartyName,P.AddressLine1,P.AddressLine2,P.City,P.PinCode From App_M_PartyMaster P,Acc_T_VoucherDetails V,Acc_T_VoucherHeader H Where H.VoucherNumber = '" & sVoucherNo & "' and H.TransactionNumber = V.TransactionNumber and V.AccUnitPartyCode = P.PartyCode"
'Response.Write sQuery
objRs.Open sQuery, con
if not objRs.EOF Then
	sPartyName		=	objRs(0)
	sPartyAddress1	=	objRs(1)
	sPartyAddress2	=	objRs(2)
	sPartyCity		=	objRs(3)
	sPartyPin		=	objRs(4)
end if

objRs.Close

'++++++++++++++++++ BEGIN Construction Of HTML +++++++++++++++++'
%>
<%
'++++++++++++++++++ END Construction Of HTML +++++++++++++++++'

'+++++++++++++++++++++++++++++++++++ BEGIN Construction Of Text File '
sPrintHeader = sPrintHeader & addressHeader()
'__________________ ITEM DETAILS population start _________________'
i=0
sQuery = "SELECT PayablesNumber, PaidByTransactionNo, convert(char, PaidOn, 103), convert(int, AmountPaid) FROM Acc_T_PybleAdjustmentDetails WHERE PaidByTransactionNo = " & iVoucherNumber & "  "
With objRs
	.CursorLocation=3
	.CursorType=3
	.ActiveConnection=con
	.Source=sQuery
	.Open
End With
Set objRs.ActiveConnection = Nothing
	
while not objRs.EOF
	sPayablesNo = objRs(0)
	sPaidByTransactionNo = objRs(1)
	sPaidOn	= objRs(2)
	sAmountPaid	= objRs(3)

	sQuery = "SELECT PartyBillNumber, convert(char,PartyBillDate,103), TransactionNumber FROM Acc_T_Payables where PayablesNumber = '" & sPayablesNo &"'"
		
	With objRs1
		.CursorLocation=3
		.CursorType=3
		.ActiveConnection=con
		.Source=sQuery
		.Open
	End With
	if not objRs.EOF Then
		sPartyBillNumber	=	objRs1(0)
		sPartyBillDate		=	objRs1(1)
		sTransactionNumber	=	objRs1(2)
	end if

	while not objRs1.EOF
		sQuery = "SELECT VoucherNumber FROM ACC_T_VOUCHERHEADER WHERE TransactionNumber = " & iVoucherNumber &""
		With objRs2
			.CursorLocation=3
			.CursorType=3
			.ActiveConnection=con
			.Source=sQuery
			.Open
		End With
		if not objRs.EOF Then
			sLinkRef	=	objRs2(0)
		end if
		i = i + 1
		j = j + 1

		sPrintHeader=sPrintHeader & " " & vbCrLf & myAlign( i ,aiHeaderColWidth(1,0),"R")
		sPrintHeader=sPrintHeader & "  " & myAlign( sOUdefinitionId,aiHeaderColWidth(1,1),"R")
		sPrintHeader=sPrintHeader & "  " & myAlign( sLinkRef,aiHeaderColWidth(1,2),"L")
		tempvar = sPartyBillNumber & " Dt." & sPartyBillDate
		sPrintHeader=sPrintHeader & "  " & myAlign( tempvar,aiHeaderColWidth(1,3),"L")
		sPrintHeader=sPrintHeader & "  " & myAlign( sAmountPaid,aiHeaderColWidth(1,4),"R")
		sPrintHeader=sPrintHeader & "      " & myAlign( "D",aiHeaderColWidth(1,5),"L")
		sTotalAmountPaid = sTotalAmountPaid + sAmountPaid

		objRs2.MoveNext
		objRs1.MoveNext
	wend
				objRs2.Close
		objRs1.Close
		objRs.MoveNext
	wend
objRs.Close
Dim k
k=18-j
	for i=1 to k
		sPrintHeader=sPrintHeader & " " & vbCrLf
	next
'============================================== under rs.
	sPrintHeader = sPrintHeader & " " & vbCrLf
	sPrintHeader=sPrintHeader & myAlign( "" ,aiHeaderColWidth(1,0),"R")
	sPrintHeader=sPrintHeader & "  " & myAlign( "",aiHeaderColWidth(1,1),"R")
	sPrintHeader=sPrintHeader & "  " & myAlign( "",aiHeaderColWidth(1,2),"L")
	sPrintHeader=sPrintHeader & "  " & myAlign( "TOTAL",aiHeaderColWidth(1,3),"L")
	sPrintHeader=sPrintHeader & "  " & myAlign( FormatNumber(iTotal,2,,,0),aiHeaderColWidth(1,4),"R")
 	sPrintHeader=sPrintHeader & "      " & myAlign( "D",aiHeaderColWidth(1,5),"L")

for i=1 to 8
	sPrintHeader = sPrintHeader & " " & printblankline()
next

iTotal = CDbl(iTotal)
'__________________ item details population end _________________'
'============================================== under rs.
sPrintHeader = sPrintHeader &  " " & vbCrLf
sPrintHeader = sPrintHeader & myAlign( "",aiHeaderColWidth(2,0),"L")
tempvar = Cdbl(iTotal + 1)
sPrintHeader = sPrintHeader & myAlign( tempvar,aiHeaderColWidth(2,1),"L")
'============================================== date
sPrintHeader = sPrintHeader &  " " & vbCrLf &  " " & vbCrLf
sPrintHeader = sPrintHeader & myAlign( "",aiHeaderColWidth(2,0),"L")
sPrintHeader = sPrintHeader & myAlign( "",aiHeaderColWidth(2,1),"L")
sPrintHeader = sPrintHeader & myAlign( sVoucherDt,aiHeaderColWidth(2,2),"L")

'============================================== pay
sPrintHeader = sPrintHeader &  " " & vbCrLf &  " " & vbCrLf
sPrintHeader = sPrintHeader & myAlign( "",aiHeaderColWidth(3,0),"L")
sPrintHeader = sPrintHeader & myAlign( sPayTo,aiHeaderColWidth(3,1),"L")

'============================================== amount
sPrintHeader = sPrintHeader &  " " & vbCrLf &  " " & vbCrLf &  " " & vbCrLf
sPrintHeader = sPrintHeader & myAlign( "",aiHeaderColWidth(4,0),"L")
tempvar = AmountWords(iTotal)
sPrintHeader = sPrintHeader & myAlign( tempvar,aiHeaderColWidth(4,1),"L")

'============================================== amount in numbers
sPrintHeader = sPrintHeader &  " " & vbCrLf &  " " & vbCrLf
sPrintHeader = sPrintHeader & myAlign( "",aiHeaderColWidth(5,0),"L")
sPrintHeader = sPrintHeader & myAlign( FormatNumber(iTotal,2,,,0),aiHeaderColWidth(5,1),"R")

sPrintHeader = sPrintHeader &  " " & vbCrLf

objTxt.Write sPrintHeader &  " " & vbCrLf
'++++++++++++++++++ END Construction Of File +++++++++++++++++'

'----------- To Get The Values From the Selection Page ----------------
'saTemp=split(sGetVal,"|")

'================================= USER DEFINED FUNCTIONS ================================='
'++++++++++++++++++ This aligns the string passed either to right or left +++++++++++++++++'
	function myAlign(val1,alen,str1)
	dim vlen,k,str2,val
		val=val1
		vlen = CInt(len(val))
		if (vlen > alen) then
			val = Mid(val,1,alen-1)
			vlen = CInt(len(val))
		end if
		k = (alen - vlen)
		if alen = vlen then
		   str2 = val
		     myAlign = str2
		else if (str1="L") then
			str2 = val & String(k," ")
			myAlign = str2
		    else if (str1 = "R") then
			         str2 = String(k," ") & val
			         myAlign = str2
		          end if
		    end if
		end if
	end function
	'------------------------End OF myAlign Function----------------------------

'++++++++++++++++++ This aligns the string passed either to right or left +++++++++++++++++'
	function centerAlign(str1,width)
		dim diff,strlen,val, i, str, newstr, blank
			str = str1
			strlen = len(str)
			diff = width - strlen
			for i=0 to (diff-1)/2
				blank = blank & " "
			next
			newstr = blank & str & blank
		centerAlign = newstr
	end function
	'------------------------End OF myAlign Function----------------------------
'++++++++++++++++++ This prints the address details of the header region +++++++++++++++++'
	function addressHeader()
		dim stStr

		stStr=stStr & myAlign("",aiHeaderColWidth(0,0),"L")
		stStr=stStr & myAlign(sPartyName,aiHeaderColWidth(0,1),"L")

		stStr=stStr & vbCrLf
		stStr=stStr & myAlign("",aiHeaderColWidth(0,0),"L")
		stStr=stStr & myAlign(sPartyAddress1,aiHeaderColWidth(0,1),"L")

		stStr=stStr & vbCrLf
		stStr=stStr & myAlign("",aiHeaderColWidth(0,0),"L")
		stStr=stStr & myAlign(sPartyAddress2,aiHeaderColWidth(0,1),"L")
		stStr=stStr & myAlign("",aiHeaderColWidth(0,2),"L")
		stStr=stStr & myAlign(sChequeNo,aiHeaderColWidth(0,3),"L")
		stStr=stStr & myAlign("",aiHeaderColWidth(0,4),"L")
		stStr=stStr & myAlign(sChequeDt,aiHeaderColWidth(0,5),"L")
		stStr=stStr & vbCrLf
		stStr=stStr & myAlign("",aiHeaderColWidth(0,0),"L")
		tempvar = sPartyCity & " - " & sPartyPin
		stStr=stStr & myAlign(tempvar,aiHeaderColWidth(0,1),"L")
		stStr=stStr & vbCrLf
		stStr=stStr & myAlign("",aiHeaderColWidth(0,0),"L")
		stStr=stStr & myAlign("",aiHeaderColWidth(0,1),"L")
		stStr=stStr & myAlign("",aiHeaderColWidth(0,2),"L")
		stStr=stStr & myAlign(sVoucherNo,aiHeaderColWidth(0,3),"L")
		stStr=stStr & myAlign("",aiHeaderColWidth(0,4),"L")
		stStr=stStr & myAlign(sVoucherDt,aiHeaderColWidth(0,5),"L")
		stStr=stStr & " " & vbCrLf & " " & vbCrLf & " " & vbCrLf & " "
		addressHeader = stStr
	end function
	
	function printblankline()
		dim stStr
		stStr=stStr & " " & vbCrLf
		printblankline = stStr
	end function

iRecCount = 2
if cint(iRecCount)>=0 then
	Response.Redirect("../../Components/FormattPrint.asp?server=server&filepath=/accounts/temp/Transaction/cheque.txt&exitpath=/accounts/Transaction/VouBAView.asp&frame=_parent")
else
	Response.Clear
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 FINAL//EN">
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<TITLE>No Records </TITLE>
</HEAD>
<BODY BGCOLOR="#CCCCCC" LINK="#0000FF" VLINK="#800080" TEXT="#000000" TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>
<TABLE BORDER=0 ALIGN=CENTER CELLSPACING=0 CELLPADDING=0 NOF=LY>
    <TR >
		<TD height="20">&nbsp;</TD>
		<TD WIDTH=549 ><P ALIGN=CENTER><B>
		<FONT SIZE="-1" FACE="Arial,Helvetica,Univers,Zurich BT">No Records Found</FONT></B></TD>
	</TR>
	<TR>
        <TD height="20">&nbsp;</TD>
        <TD WIDTH=549 ><P ALIGN=CENTER><B>
        <FONT SIZE="-1" FACE="Arial,Helvetica,Univers,Zurich BT"><a href="#" onclick="window.history.back(1); return false;">Back</a></FONT></B></TD>
    </TR>
</Table>
</HTML>

<%end if%>


