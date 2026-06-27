<%@ Language=VBScript %>
<%	option explicit	%>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name					:	PRNBankVoucherEntry.asp
	'Module Name					:
	'Author Name					:	S.MAHESWARI
	'Modified By					:
	'Created On						:	24-Oct-2008 
	'Modified On					:	 
	'Tables Used					:
	'Temporary Tables				:
	'Temporary Files				:
	'Input Parameter				:	None
	'Connects To					:	BankVouchView_San.asp(Print Voucher Entries)
	'Procedures/Functions Used		:
	'Internal Variables				:
	'Database						:	iTMS_KSS_Test
	'Queries Used					:
	'Counters						:
	'String							:
	'Boolean						:
	'Object Holders					:
	'Description					:
%>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/PrintFunctions.asp"-->
<!--#include File="../../include/GetOrganization.asp"-->

			<%
'------------------------Declaration Constants -----------------------------
dim aiHeaderColWidth(5,9),objFSO,objTxt
dim sTextOut,sTempStr,sPaidTo,iAmount
dim sDetails, sReceivedPayment,sPreparedBy,sCheckedBy,sPassedBy,dcrs,iPageNo,sRetVal,i
dim sTemp,sHeadOfAcc,sAccFlag,sBInsDtFlag
dim sFlagTotAmt,sum,sBInsNoFlag

set		dcrs		= server.CreateObject("adodb.recordset")
set		objFSO	= Server.CreateObject("Scripting.FileSystemObject")
set		objTxt	= objFSO.CreateTextFile(server.MapPath("../temp/Transaction/"&Session.SessionID&"_BankVoucherEntry_View.txt"))
		sTempStr	= ""

			'XML DOM Variables
			Dim oDOM,Root,objRs,objRs1,sQuery,objRsTemp
			dim sNarration,sAccount,sAddtional,iSno,sNarr
			dim dTotal,sOrgId,sAccHead,sType
			dim EntryNode,HeaderNode,dAmount,sAccNo,sPartyName
			dim iVouNo,sOrgName,sBookName,sVouType,sApprove,sVoucDate,iBookCode,sPayTo
			dim iTransNo,iBkHeadCode,bOtherUnit,iTdsAmount,sExp,TempNode,sBankInsDet
			Dim sInstrType,sBankInsNo, sBankInsName,sBankInsDrawnOn,sBankInsDt
			Dim sAddress1, sAddress2,sCity,sState,sPostcode,sTranIndication,sTranEntryIndication
			Dim iCreatedBy,sCreatedOn,sVouStatus,sEmpName,iPartyCtrlAcc,sAdjType,sDetFlag
			Dim iEntryNo,iHeadOfAcc,sHeadOfAccName,iHeadOfAccAmt,iPartyCode,iEntryAmt,iCtr
			Dim iNetAmtPaid,iTotRecovered,sNarrFlag,sAddnFlag,iTotAddnlAmt,sAdjFlag,sAdjOn,iTotAdj,sBankInsAmt
			' Create our DOM Document Objects
			Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
			Dim iOne,iThree,iNoOfLinesCtr,iFlag12,iFlag34,sDrwOn,sPayAt

			iTransNo=Request("Value")
			'Response.Write iTransNo
			iFlag12 = false
			iFlag34 = false
			iNoOfLinesCtr = 0

			'oDOM.Load server.MapPath("../xmldata/Voucher/"&iTransNo&".xml")
			sRetVal = GetVouchXML(iTransNo)
			oDOM.Load server.MapPath(sRetVal)

			set Root=oDOM.documentElement
			sOrgId=Root.Attributes.Item(0).nodeValue
			sOrgName=Root.Attributes.Item(1).nodeValue
			iBookCode =Root.Attributes.Item(2).nodeValue
			sBookName =Root.Attributes.Item(3).nodeValue
			sVouType=Root.Attributes.Item(4).nodeValue
			sVoucDate=Root.Attributes.Item(5).nodeValue
			iBkHeadCode=Root.Attributes.Item(6).nodeValue

			iVouNo=Root.Attributes.Item(9).nodeValue
			sApprove=Root.Attributes.Item(7).nodeValue

			set objRs = Server.CreateObject("ADODB.Recordset")
			set objRs1 = Server.CreateObject("ADODB.Recordset")
			set objRsTemp = Server.CreateObject("ADODB.Recordset")
					
					
'------------------------Start of Declaration Constants ----------------------
'No and Date
aiHeaderColWidth(0,0)=67
aiHeaderColWidth(0,1)=16

'Head of Account, Amount
aiHeaderColWidth(1,0)=14
aiHeaderColWidth(1,1)=70

'Paid To, Rupees
aiHeaderColWidth(2,0)=14
aiHeaderColWidth(2,1)=46
aiHeaderColWidth(2,2)=7
aiHeaderColWidth(2,3)=16

'NameOfBank, Details
aiHeaderColWidth(3,0)=3
aiHeaderColWidth(3,1)=20
aiHeaderColWidth(3,2)=3
aiHeaderColWidth(3,3)=21
aiHeaderColWidth(3,4)=80
aiHeaderColWidth(3,5)=10
aiHeaderColWidth(3,6)=1
aiHeaderColWidth(3,7)=15
aiHeaderColWidth(3,8)=53
aiHeaderColWidth(3,9)=10


'Prepared by, Checked by, Passed by
aiHeaderColWidth(4,0)=3
aiHeaderColWidth(4,1)=16
aiHeaderColWidth(4,2)=3
aiHeaderColWidth(4,3)=18
aiHeaderColWidth(4,4)=2
aiHeaderColWidth(4,5)=18

'------------------------End of Declaration Constants ----------------------					
sQuery = "Select Distinct CreatedVoucherNo,TransactionType,OUDefinitionID,PayToRecdFrom," &_
					"PartyCode,CreatedBy,Convert(varchar,CreatedOn,103),CreatedVouchStatus," &_
					"BankInstrumentType,BankInstrumentNo,Convert(varchar,BankInstrumentDate,103),PayableAt,DrawnOnBank" &_
					" from VW_Created_BankVoucherView where CreatedTransNo="& iTransNo
				'Response.Write sQuery
			with objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			set objRs.ActiveConnection = nothing
			if not 	objRs.EOF then
				iVouNo = objRs(0)
				sVouType = objRs(1)
				sOrgId = objRs(2)
				sPayTo = objRs(3)
				iPartyCode = objRs(4)
				iCreatedBy = objRs(5)
				sCreatedOn = objRs(6)
				sVouStatus = objRs(7)
				sInstrType = objRs(8)
			 
			end if
			objRs.Close

			'Newly added on July 2nd 2008 by S.Maheswari to fetch bank details from Acc_T_CreatedVoucherInstrumentDet table instead of taking from CreatedVoucherHeader table
			sQuery = "Select BankInstrumentNo,convert(Varchar,BankInstrumentDate,103),BankInstrumentType,InstrumentAmount,PayableAt,DrawnOnBank "&_
					" from Acc_T_CreatedVoucherInstrumentDet where CreatedTransNo = "&iTransNo&" Order by InstrumentEntryNo "
			' Response.Write sQuery
			' Response.ENd

			objRs.Open sQuery,con
			If Not objRs.EOF then
				Do while not objRs.EOF
					sBankInsNo		= sBankInsNo&","&objRs(0)
					sBankInsDt		= sBankInsDt&","&objRs(1)
					sBankInsName	= sBankInsName&","&objRs(2)
					sBankInsAmt	    = sBankInsAmt&","&objRs(3)
					sDrwOn			= sDrwOn &","&objRs(5)
					sPayAt			= sPayAt &","&objRs(4) 
					objRs.MoveNext
				loop
			End If
			'Response.Write sDrwOn &"===="& sPayAt 
			
			'Response.end 
			objRs.Close
			
			
			sBankInsNo		= mid(sBankInsNo,2)
			sBankInsDt		= mid(sBankInsDt,2)
			sBankInsName	= mid(sBankInsName,2)
			sBankInsAmt	    = mid(sBankInsAmt,2)
			sDrwOn			= mid(sDrwOn,2)
			sPayAt			= mid(sPayAt,2)
		'	sBankInsDrawnOn = mid(sBankInsDrawnOn,2)
			sBankInsDrawnOn = sDrwOn&","&sPayAt
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		
		If trim(sVouStatus) = "010104" Then
				sQuery = "Select VoucherNumber from ACC_T_VoucherHeader where CreatedTransNo="&iTransNo
				with objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
				end with
				if not 	objRs.EOF then
					iVouNo = objRs(0)
				End If
				objRs.Close
			End If
iPageNo = 1

sQuery = "Select VoucherEntryNumber,AccUnitAccountHead,Amount,VoucherNarration,AccUnitPartyCode,"&_
		 "TransCrDrIndication from Acc_T_CreatedVoucherdetails where CreatedTransNo= "& iTransNo
objRs.Open sQuery,con
Do while not objRs.EOF 
 
	iEntryNo = objRs(0)
	iHeadOfAcc = objRs(1)
	iHeadOfAccAmt = FormatNumber(objRs(2),2,,,0)
	sNarr = objRs(3)
	iPartyCtrlAcc = objRs(4)
  	'Response.Write iPartyCtrlAcc &"<BR>"
	IF iHeadOfAcc <> "" Then
		sQuery = "select AccountDescription from ACC_M_GLAccountHead where AccountHead ="&iHeadOfAcc
	Else
		If trim(iPartyCtrlAcc) <> "" then 	sQuery = "SELECT PARTYNAME FROM APP_M_PARTYMASTER WHERE PARTYCODE ="&iPartyCtrlAcc
	End If
	'Response.write sQuery
	objRs1.Open sQuery,con
	if not 	objRs1.EOF then
		sHeadOfAcc = objRs1(0)
	end if
	objRs1.Close
	 
'Function Header(iPageNo)
'iPageNo=1
	if sType="P" then
		sPaidTo			= sPartyName
	else
		sPaidTo			= sPayTo
	end if
	iNoOfLinesCtr = 0
	'Blank Lines
	if iPageNo > 1 then 
		iNoOfLinesCtr = 0
		sTextOut = sTextOut & " " & vbcrlf & vbcrlf & vbcrlf & vbcrlf & vbcrlf & vbcrlf & vbcrlf  
	End if
	sTextOut = sTextOut & " " & vbcrlf & Vbcrlf
	iNoOfLinesCtr = iNoOfLinesCtr + 2
	'Number and date

	 sTextOut =  sTextOut & myAlign("",aiHeaderColWidth(0,0)+6,"L") 

	'Response.Write  sHeadOfAcc
	IF Cstr(sVouStatus) = "010104" Then
		 sTextOut =  sTextOut & myAlign(iVouNo,aiHeaderColWidth(0,1),"L")
	Else
		 sTextOut =  sTextOut & myAlign(" ",aiHeaderColWidth(0,1),"L")
	End IF 
	sTextOut = sTextOut & vbCrLf
	iNoOfLinesCtr = iNoOfLinesCtr + 1
	
	sTextOut = sTextOut & " " & vbcrlf
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(0,0)+6,"L")
	sTextOut = sTextOut  & myAlign(sVoucDate,aiHeaderColWidth(0,1),"L")
'	
	sTextOut = sTextOut & vbCrLf
	iNoOfLinesCtr = iNoOfLinesCtr + 2
	'Response.Write iNoOfLinesCtr
	'Response.End 
	'Blank Lines
	sTextOut = sTextOut & " " & vbcrlf & " " & vbcrlf
	'Head of Account
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(1,0)+7,"L")
	'sTextOut = sTextOut  & myAlign(sHeadOfAccName,aiHeaderColWidth(1,1),"L")
	sTextOut = sTextOut  & myAlign(sHeadOfAcc,aiHeaderColWidth(1,1),"L")
'	
	sTextOut = sTextOut & vbCrLf & vbCrLf
	iNoOfLinesCtr = iNoOfLinesCtr + 3
	
'End Function	 
	
	Dim sDesc2,sDesc,iWidth
	sDesc	= AmountWords(replace(iHeadOfAccAmt,",",""))
	If Len(sDesc) > 50 Then
		For i = 1 to 50
			If Mid(sDesc,50-i,1) = " " Then
				iWidth = 50-i
			Exit For
			End if
		Next
		sDesc2 = Mid(sDesc,iWidth+1,Len(sDesc))
		sDesc =  Mid(sDesc,1,50-i)
	End If

	sTextOut = sTextOut & vbCrLf
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(1,0)+7,"L")
	sTextOut = sTextOut  & myAlign(sDesc,aiHeaderColWidth(1,1),"L")
	'
	sTextOut = sTextOut & vbCrLf
	
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(1,0)+7,"L")
	sTextOut = sTextOut  & myAlign(sDesc2,aiHeaderColWidth(1,1),"L")
'	
	'sTextOut = sTextOut & vbCrLf
	
	iNoOfLinesCtr = iNoOfLinesCtr + 3 
	'PaidTo,Rs
	sTextOut = sTextOut & vbCrLf & Vbcrlf
	sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(2,0)+7,"L")
	sTextOut = sTextOut  & myAlign(sPayTo,aiHeaderColWidth(2,1),"L")
	sTextOut = sTextOut  & myAlign("",5,"R")
	sTextOut = sTextOut  & myAlign(iHeadOfAccAmt,aiHeaderColWidth(2,2)+2,"R")
	'
	sTextOut = sTextOut & vbCrLf
	sTextOut = sTextOut & vbCrLf
	iNoOfLinesCtr = iNoOfLinesCtr + 4
	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	sTextOut = sTextOut  & formattprint("CONDENSESTART","") 'Condensed formatt
'sTextOut = sTextOut  & " " & vbCrLf
'iNoOfLinesCtr = iNoOfLinesCtr + 1

		 

'------------------------End of Declaration Constants ----------------------

%>  
 
<%
'Fetch the ADDITIONAL PAYMENT / RECEIPT ENTRIES
'check from this
'Response.Write sDrwOn  &"*************"&vbcrlf

'
	IF sDrwOn <> "" then
		IF trim(sDrwOn) = "AB" then
			sDrwOn = "ANDHARA BANK"
			sPayAt = "TRIPUR" 
		' Else			
		'	'Response.write sBankInsDrawnOn
		'	sTemp = split(sBankInsDrawnOn,",")
		'	sDrwOn = trim(sTemp(0))
		'	sPayAt = trim(sTemp(1))
		 End IF
		
		sTextOut = sTextOut  & vbCrLf
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5),"L") '31 Space  
		sTextOut = sTextOut  & myAlign(trim(sDrwOn),aiHeaderColWidth(3,3),"L") & vbCrLf '21
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5),"L") '31 Space  
		sTextOut = sTextOut  & myAlign(trim(sPayAt),aiHeaderColWidth(3,3),"L") 
	Else
		sTextOut = sTextOut  & vbCrLf
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5),"L") '31 Space  
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3),"L") & vbCrLf '21
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5),"L") '31 Space  
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3),"L") 
	End If
	sTextOut = sTextOut  & myAlign("",4,"L") 
	iNoOfLinesCtr =  iNoOfLinesCtr + 3
	sum = len(sNarr) / 80
	sum = Round(sum)
	
	IF sNarr <> "" then
		IF len(sNarr) > 80 then			
			sNarr = BreakString(sNarr,80)	
			'for i = 1 to len(sNarr) Step 80
			for i = 0 to UBOUND(sNarr) 
				
				'sTextOut = sTextOut  & myAlign(UCASE(trim(mid(sNarr,i,80))),aiHeaderColWidth(3,4),"L")& vbCrLf  '80
				 'sTextOut = sTextOut & myAlign(UCASE(trim(sNarr(i))),80,"L")& vbCrLf  '80
				 sTextOut = sTextOut & myAlign("",1,"L") & myAlign(UCASE(trim(sNarr(i))),80,"L")& vbCrLf  '80
					
				iNoOfLinesCtr = iNoOfLinesCtr + 1 
				
				IF sum = i then		
					sNarr = ""
					Exit for
				End If
				IF iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22 or iNoOfLinesCtr = 24 then 
					BankDetails(iNoOfLinesCtr)
				else
					sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5),"L") '31 Space  
					sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+4,"L")  '21
				End IF
			next
		else
			'sTextOut = sTextOut & myAlign(UCASE(trim(sNarr)),aiHeaderColWidth(3,4),"L") & vbCrLf   
			sTextOut = sTextOut & myAlign("",1,"L") & myAlign(UCASE(trim(sNarr)),aiHeaderColWidth(3,4),"L") & vbCrLf
			IF iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22 or iNoOfLinesCtr = 24 then 
			   	BankDetails(iNoOfLinesCtr)
			else
				sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5),"L") '31 Space  
				sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+4,"L")  '21
			End IF
			iNoOfLinesCtr = iNoOfLinesCtr + 1
		End IF
			sNarrFlag = True
	End IF

	sTextOut = sTextOut  & vbCrLf
	iNoOfLinesCtr = iNoOfLinesCtr + 1

 
	
	For i = iNoOfLinesCtr to 29
		IF iNoOfLinesCtr = 20 or iNoOfLinesCtr = 22 or iNoOfLinesCtr = 24 then 
			BankDetails(iNoOfLinesCtr)
			iNoOfLinesCtr = iNoOfLinesCtr +1	
		End IF
		
	Next  
	for i = iNoOfLinesCtr to 30
		sTextOut = sTextOut & Vbcrlf	
		iNoOfLinesCtr = iNoOfLinesCtr + 1
	Next
	
	sTextOut = sTextOut & formattprint("CONDENSEEND","") 'Condensed formatt
	sTextOut = sTextOut & myAlign("",aiHeaderColWidth(3,0)+7,"L")
	
	sQuery = "SELECT LoginId FROM DCS_User WHERE InternalUserID ="&iCreatedBy
	objRs1.Open sQuery,con
	IF not objRs1.EOF  Then
		sEmpName = Trim(Objrs1(0))
	End IF
	objRs1.Close	
	IF Cstr(sVouStatus) <> "010104" Then
		sTextOut = sTextOut & myAlign(sEmpName&"/" & iVouNo & "-"&sCreatedOn,aiHeaderColWidth(3,7)+4,"L")
	Else
		sTextOut = sTextOut & myAlign(sEmpName&"-"&sCreatedOn,aiHeaderColWidth(3,7),"L")
	End IF
	'sTextOut = sTextOut & chr(12)
	iPageNo = iPageNo + 1
	 
objRs.MoveNext 
loop
objRs.Close

' Response.End 
	objTxt.write sTextOut
	Response.Redirect("../../Components/FormattPrintNew.asp?server=server&filepath=/accounts/temp/Transaction/"&Session.SessionID&"_BankVoucherEntry_View.txt&exitpath=/accounts/reports/VouBAView.asp&frame=_parent")
	'Response.Redirect("../../Components/FormattPrint.asp?server=server&filepath=/accounts/temp/Transaction/"&Session.SessionID&"_BankVoucher_View.txt&exitpath=/accounts/reports/VouBAView.asp&frame=_parent")
%>
<%
Function BankDetails(iLineNo)
dim Flag 

if iPageNo > 1 and Flag = "" then 
	sAccFlag = ""
	sBInsNoFlag = ""
	sBInsDtFlag = ""
end if
if iPageNo > 1 then  Flag = "Y"
	IF iLineNo = 20 and sAccFlag <> True then   
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5),"L") '31 Space  
		sTextOut = sTextOut  & myAlign("OCC90116",aiHeaderColWidth(3,3),"L")
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)+1,"L")'2 Space
		sAccFlag = True  
		sTextOut = sTextOut  & vbCrLf
		iNoOfLinesCtr = iNoOfLinesCtr +1
	End If  
 
	IF iLineNo = 22  and sBInsNoFlag <> true then
		sTextOut = sTextOut & Vbcrlf
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5),"L") '31 Space
		sTextOut = sTextOut  & myAlign(sBankInsNo,aiHeaderColWidth(3,3),"L") '14
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)+1,"L")'2 Space
		sBInsNoFlag = True
		sTextOut = sTextOut  & vbCrLf
		iNoOfLinesCtr = iNoOfLinesCtr +1
	End If 
	
	If iNoOfLinesCtr = 24 and sBInsDtFlag <> True then 
		sTextOut = sTextOut & Vbcrlf
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,3)+aiHeaderColWidth(3,5),"L") '31 Space
		sTextOut = sTextOut  & myAlign(sBankInsDt,aiHeaderColWidth(3,3),"L") '22
		sTextOut = sTextOut  & myAlign("",aiHeaderColWidth(3,0)+1,"L")'2 Space
		sBInsDtFlag = True
		sTextOut = sTextOut  & vbCrLf
		iNoOfLinesCtr = iNoOfLinesCtr +1
	End If 
	
End Function

'================================= USER DEFINED FUNCTIONS ================================='
'++++++++++++++++++ This aligns the string passed either to right or left +++++++++++++++++'
	function myAlign(val1,alen,str1)
	dim vlen,k,str2,val
		val=val1
		IF len(val) then vlen = CInt(len(val))
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
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 FINAL//EN">
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<TITLE>iTMS -
		  <%
			if sVouType="BAP" then
				Response.Write "Cheque Payment Voucher"
			else
				Response.Write "Cheque Receipt Voucher"
			end if
		%>
 </TITLE>
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
        <FONT SIZE="-1" FACE="Arial,Helvetica,Univers,Zurich BT"><a href="javascript:window.history.back(1)">Back</a></FONT></B></TD>
    </TR>
</Table>
</HTML>

