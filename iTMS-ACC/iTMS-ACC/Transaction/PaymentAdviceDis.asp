 <%@ Language=VBScript %>
<%	option explicit	%>
<%
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PaymentAdviceDis.asp
	'Module Name				:
	'Author Name				:	Ragavendran R
	'Created On					:	May 07,2013
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/PrintFunctions.asp"-->
<!--#include File="../../include/GetOrganization.asp"-->

<%
dim aiHeaderColWidth(5,9)
dim objFSO,dcrs,dcrs1
dim sTextOut,sTempStr, nCashVoucherNo,dCashVoucherDate,sHeadOfAccount,sAmount,sPaidTo,iAmount,iAccAmount
dim sDetails, sReceivedPayment,sPreparedBy,sCheckedBy,sPassedBy,iPageNo,sRetVal,iTotVouAmt,i,iHeadAmount
dim iTransNo,sQuery,iPayNo,sParBillNo,dtParBillDate,pCrTransNo,sPayDocNo,iCreatedBy 
dim iVouNo,dtVouDate,iVouAmt,dtCrVouDate,sOrgId,iCheqNo,dtCheqdate,iAmtPaid,iLineNo
dim iPartyCode,sPartyName,sSuppName,sAddress1,sAddress2,sAddress3 
dim sUser,dtPrnDate,tmPrnTime,sPara,iTotAmtPaid
set		dcrs		= server.CreateObject("adodb.recordset")
set		dcrs1		= server.CreateObject("adodb.recordset")
set		objFSO	= Server.CreateObject("Scripting.FileSystemObject")

iPageNo=1
iLineNo = 0
sTempStr	= ""
sSuppName	= ""
sAddress1	= ""
sAddress2	= ""
sAddress3	= ""
iTransNo=Request("Value")

sQuery = "Select isNull(CreatedVoucherNo,''),Convert(Varchar,VoucherDate,103),OUDefinitionID,isnull(BankInstrumentNo,''),Convert(Varchar,isnull(BankInstrumentDate,''),103),CreatedBy,convert(Varchar,CreatedOn,103),convert(Varchar,CreatedOn,108),VoucherAmount from Acc_T_CreatedVoucherHeader where CreatedTransNo = " & iTransNo

dcrs.Open sQuery,con
if not dcrs.EOF then	
	sPayDocNo	= dcrs(0)
	dtCrVouDate	=  dcrs(1)
	sOrgId		= dcrs(2)
	iCheqNo		= dcrs(3)
	dtCheqdate	= dcrs(4)
	iCreatedBy	= dcrs(5)
	dtPrnDate   = dcrs(6)
	tmPrnTime	= dcrs(7)
	iHeadAmount = dcrs(8)
end if
dcrs.Close 

'Newly added on Dec 16th 2008

If trim(iCheqNo) = "" then 
	sQuery = "Select BankInstrumentNo,Convert(Varchar,BankInstrumentDate,103) from Acc_T_CreatedVoucherInstrumentDet where CreatedTransNo = " & iTransNo
	dcrs.Open sQuery,con
	if not dcrs.EOF then	
		iCheqNo = dcrs(0)
		dtCheqdate = dcrs(1)
	end if
	dcrs.Close 
End If
If trim(dtCheqdate) = "01/01/1900" then dtCheqdate = "" 

sUser = Trim(session("username"))

sQuery = "Select PartyName,isNull(AddressLine1,''),isNull(AddressLine2,'') ,isNull(City,''),isNull(State,''),isNull(PinCode,'') from App_M_PartyMaster Where Partycode in "&_
		" (Select isnull(AccUnitPartyCode,0) from Acc_T_CreatedVoucherDetails where CreatedTransno = "& iTransNo &" )"
dcrs.Open sQuery,con
if not dcrs.EOF then	
	sSuppName = dcrs(0)
	IF trim(dcrs(1)) <> "" then sAddress1 =  dcrs(1) &","&  dcrs(2)
	sAddress2 =  dcrs(3)
	IF trim(dcrs(4)) <> "" then  sAddress3 =  dcrs(4) &"-"&  dcrs(5)
	
end if
dcrs.Close 

iTotVouAmt = 0
iTotAmtPaid = 0
'-------------Function Start------------------
Header()
InvDetails()


sPara = ""
If iLineNo > 60 then 
	NextPageDetails(sPara)	
End If

sTempStr = sTempStr & myAlign(" ",7,"L") & myAlign("Kindly Acknowledge",20,"L") & vbCrLf
for i = 0 to 4
	sTempStr = sTempStr  & vbCrLf
	iLineNo = iLineNo + 1
	If iLineNo > 60 then 
		NextPageDetails(sPara)	
	End If
next
 
sTempStr = sTempStr & myalign(" ",7,"L") & myalign("For K.Sivasubramaniam Spinners (P) Ltd",40,"L") & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
stempstr = stempstr & myalign(" ",7,"L") & myalign("V.Venkatesh",38,"R")  & vbCrLf 
stempstr = stempstr & myalign(" ",7,"L") & myalign("Authorised Signatory",38,"R")  & vbCrLf & vbCrLf


sTempStr = sTempStr & myalign(" ",7,"L") & myalign("Note : We request you to send the cheque For clearing to Tirupur branch",75,"L") & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf

'-------------Function End--------------------

sTextOut = sTextOut & sTempStr  

%>
<%
Function Header()
	sTempStr = sTempStr & vbCrLf & vbCrLf & vbCrLf & vbCrLf  & vbCrLf & vbCrLf '& vbCrLf & vbCrLf 
	sTempStr = sTempStr & myAlign(" ",30,"L") &  centerAlign("PAYMENT ADVICE",20) & myAlign(" ",30,"L") & vbCrLf
	sTempStr = sTempStr & vbCrLf & vbCrLf
	'sTempStr = sTempStr & myAlign(" ",7,"L") &  myAlign("Supplier Name & Location :",30,"L") & myAlign(" ",6,"L")  & myAlign("Reference No:",14,"L") & myAlign(sPayDocNo&" ("&dtCrVouDate &")" ,20,"L")& vbCrLf
	sTempStr = sTempStr & myAlign(" ",7,"L") &  myAlign(" ",30,"L") & myAlign(" ",6,"L")  & myAlign("Reference No:",14,"L") & myAlign(sPayDocNo&" ("&dtCrVouDate &")" ,20,"L")& vbCrLf
	sTempStr = sTempStr & myAlign(" ",7,"L") &  myAlign(sSuppName,45,"L") & vbCrLf
	sTempStr = sTempStr & myAlign(" ",7,"L") &  myAlign(sAddress1,70,"L") & vbCrLf
	sTempStr = sTempStr & myAlign(" ",7,"L") &  myAlign(sAddress2,70,"L") & vbCrLf
	sTempStr = sTempStr & myAlign(" ",7,"L") &  myAlign(sAddress3,70,"L") & vbCrLf
	sTempStr = sTempStr & vbCrLf & vbCrLf
	sTempStr = sTempStr & myAlign(" ",7,"L") & formattprint("CONDENSESTART","") &  myAlign("Enclosed please find the details of bills against which the payment is made.",76,"L") & formattprint("CONDENSEEND","")& vbCrLf
	sTempStr = sTempStr & vbCrLf & vbCrLf
	sTempStr = sTempStr & myAlign(" ",7,"L") &  myAlign("Vide Cheque No:",15,"L") & myAlign(iCheqNo,25,"L") & myAlign("",1,"L") & myAlign("Dated : ",10,"L") & myAlign(dtCheqdate,15,"L") & vbCrLf
	iLineNo = iLineNo + 19	
	  
End Function 
Function InvDetails()
dim sPara
	sTempStr = sTempStr & formattprint("CONDENSESTART","")& vbCrLf
	sTempStr = sTempStr & myAlign(" ",12,"L") & string(103,"-")& vbCrLf
	sTempStr = sTempStr & myAlign(" ",12,"L") & centerAlign("Invoice No ",16) & myAlign(" ",1,"L") & centerAlign("Invoice",11) & myAlign(" ",1,"L")  & myAlign(" ",1,"L") & centerAlign("Invoice",13) & myAlign(" ",1,"L") & centerAlign("Deduction",28)                                              & myAlign(" ",1,"L") & centerAlign("Paid",13) & myAlign(" ",1,"L") & centerAlign("Remarks",12) & vbCrLf
	sTempStr = sTempStr & myAlign(" ",12,"L") & centerAlign(" ",16)             & centerAlign("Date",11)    & myAlign(" ",1,"L") & myAlign(" ",1,"L") & centerAlign("Amount",13)  & myAlign(" ",1,"L") & centerAlign("T.D.S",8)& centerAlign("Advance",8)& centerAlign("D.Note",12) & myAlign(" ",1,"L") & centerAlign("Amount",12) & myAlign(" ",1,"L") & centerAlign("",12) & vbCrLf   
	sTempStr = sTempStr & myAlign(" ",12,"L") & string(103,"-")& vbCrLf
	iLineNo = iLineNo +  4

	sQuery = "Select PayablesNumber from Acc_T_CreatedPybleAdjdet where CreatedTransNo = " & iTransNo &" order by 1 "
	dcrs.Open sQuery,con
	Do while not dcrs.EOF 
		iPayNo = dcrs(0)
		 
		sQuery = "Select isNull(PartyBillNumber,''),convert(VarChar,isnull(PartyBillDate,''),106),CreatedTransNo from Acc_T_CreatedPayables where PayablesNumber = " & iPayNo
		dcrs1.Open sQuery,con
		If not dcrs1.EOF then	
			sParBillNo		= dcrs1(0)
			dtParBillDate	= replace(dcrs1(1)," ","-")
			pCrTransNo		= dcrs1(2)
		End If
		dcrs1.Close
		
		sQuery = "Select VoucherNumber,convert(VarChar,VoucherDate,106),VoucherAmount from Acc_T_VoucherHeader where CreatedTransNo = " & pCrTransNo
		dcrs1.Open sQuery,con
		If not dcrs1.EOF then	
			iVouNo		= dcrs1(0)
			dtVouDate	= replace(dcrs1(1)," ","-")
			iVouAmt		= dcrs1(2)
		End If
		dcrs1.Close
		iTotVouAmt = CDbl(iTotVouAmt) + CDbl(iVouAmt)
		'Debit note column
		sQuery = "Select AmountPaid from Acc_T_CreatedPybleAdjdet where PayablesNumber =  " & iPayNo & " and CreatedTransNo = " & iTransNo
	 
		dcrs1.Open sQuery,con
		If not dcrs1.EOF then	
			iAmtPaid = dcrs1(0)
		End If
		dcrs1.Close 
		iTotAmtPaid  = cdbl(iTotAmtPaid) + cdbl(iAmtPaid)
		sTempStr = sTempStr & myAlign(" ",12,"L") & centerAlign(sParBillNo,16) & myAlign(" ",1,"L") & centerAlign(dtParBillDate,11) & myAlign(" ",1,"L")  & myAlign(formatNumber(iVouAmt,2,,,0),13,"R") & myAlign(" ",1,"L") & myAlign("0.00",8,"R") & myAlign("0.00",8,"R") & myalign(formatNumber(iAmount,2,,,0),12,"R") & myAlign(" ",1,"L") & myAlign(formatNumber(iAmtPaid,2,,,0),13,"R") & myAlign(" ",1,"L") & centerAlign(" ",12) & vbCrLf
		iLineNo = iLineNo +  1
		sPara = "D"
		If iLineNo > 60 then 
			NextPageDetails(sPara)	
		End If
		dcrs.MoveNext
	loop
	dcrs.Close


	sTempStr = sTempStr & myAlign(" ",12,"L") & string(103,"-")& vbCrLf
	sTempStr = sTempStr & myAlign(" ",12,"L") & centerAlign(" ",16) & myAlign(" ",1,"L") & centerAlign(" ",11) & myAlign(" ",1,"L")    & centerAlign(" ",13)   & myAlign(" ",1,"L") & myAlign("0.00",8,"R") & myAlign("0.00",8,"R") & myAlign(FormatNumber(iAmount,2,,0),12,"R") & myAlign(" ",1,"L") & myAlign(FormatNumber(iTotAmtPaid,2,,0),13,"R") & myAlign(" ",1,"L") & centerAlign(" ",12) & vbCrLf	  
	sTempStr = sTempStr & formattprint("CONDENSEEND","")
	sTempStr = sTempStr & vbCrLf
	sTempStr = sTempStr & myAlign(" ",7,"L") & myAlign("Total Paid Amt  : ",20,"L") & myAlign(FormatNumber(iHeadAmount,2,,0),15,"L")& vbCrLf & vbCrLf
	sTempStr = sTempStr & myAlign(" ",7,"L") & myAlign("Amount in Words : ",20,"L") & formattprint("CONDENSESTART","") & myAlign(amountwords(iHeadAmount),70,"L") & formattprint("CONDENSEEND","")& vbCrLf & vbCrLf 
	iLineNo = iLineNo + 7
End Function


Function NextPageDetails(sPara)
dim i
	iLineNo = 0
	iPageNo = iPageNo + 1
	sTempStr = sTempStr & myAlign("Page No: "&iPageNo &"(Contd....)",80,"R") &vbCrLf 
	iLineNo = iLineNo + 1
	For i = iLineNo to 7
		sTempStr = sTempStr & vbCrLf
	Next
	IF trim(sPara) = "D" then 
		sTempStr = sTempStr & myAlign(" ",7,"L") & string(107,"-")& vbCrLf
		sTempStr = sTempStr & myAlign(" ",7,"L") & centerAlign("Invoice No ",16) & myAlign(" ",1,"L") & centerAlign("Invoice",11) & myAlign(" ",1,"L")  & centerAlign("Invoice",13) & myAlign(" ",1,"L") & centerAlign("Deduction",28)                                              & myAlign(" ",1,"L") & centerAlign("Paid",13) & myAlign(" ",1,"L") & centerAlign("Remarks",12) & vbCrLf
		sTempStr = sTempStr & myAlign(" ",7,"L") & centerAlign(" ",16)             & centerAlign("Date",11)    & myAlign(" ",1,"L")  & centerAlign("Amount",13)  & myAlign(" ",1,"L") & centerAlign("T.D.S",8)& centerAlign("Advance",8)& centerAlign("D.Note",12) & myAlign(" ",1,"L") & centerAlign("Amount",12) & myAlign(" ",1,"L") & centerAlign("",12) & vbCrLf   
		sTempStr = sTempStr & myAlign(" ",7,"L") & string(107,"-")& vbCrLf
		iLineNo = iLineNo +  4
	End IF
End Function

'================================= USER DEFINED FUNCTIONS ================================='
'++++++++++++++++++ This aligns the string passed either to right or left +++++++++++++++++'
	function myAlign(val1,alen,str1)
	dim vlen,k,str2,val
		val=val1
		IF val <> "" then vlen = CInt(len(val))
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
'	'------------------------End OF myAlign Function----------------------------

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Payment Advice
</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script language="vbscript">
</script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<form method="POST" name="formname" action="">
	<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">
				  PAYMENT ADVICE
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>
								<tr>
								    <td align="center" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								    <td align="center">
								        <table cellpadding="0" cellspacing="0" width="100%">
								            <tr>
								                <td style="width:25%" class="FieldCellSub">
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                </td>
								                <td style="width:25%" class="FieldCellSub" align="right">
								                    <b>Reference No :</b>
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                    <b><%=sPayDocNo &" ("&dtCrVouDate &")"%></b>
								                </td>
								            </tr>
								            <tr>
								                <td style="width:25%" class="FieldCellSub">
								                    <b><%=sSuppName%></b>
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                </td>
								            </tr>
								            <tr>
								                <td style="width:25%" class="FieldCellSub">
								                    <b><%=sAddress1%></b>
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                </td>
								            </tr>
								            <tr>
								                <td style="width:25%" class="FieldCellSub">
								                    <b><%=sAddress2%></b>
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                </td>
								            </tr>
								            <tr>
								                <td style="width:25%" class="FieldCellSub">
								                    <b><%=sAddress3%></b>
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                </td>
								            </tr>
								            <tr>
									            <td align="center" colspan="4" class="MiddlePack">
									            </td>
								            </tr>
								            <tr>
								                <td class="FieldCellSub" colspan="4" align="center">
								                    <b>
								                        Enclosed please find the details of bills against which the payment is made.
								                    </b>
								                </td>
								            </tr>
								            <tr>
									            <td align="center" colspan="4" class="MiddlePack">
									            </td>
								            </tr>
								             <tr>
								                <td style="width:25%" class="FieldCellSub">
								                    <b>Vide Cheque No :</b>
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                    <%=iCheqNo%>
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                    <b>Dated : </b>
								                </td>
								                <td style="width:25%" class="FieldCellSub">
								                    <%=dtCheqdate%>
								                </td>
								            </tr>
								            <tr>
									            <td align="center" colspan="4" class="MiddlePack">
									            </td>
								            </tr>
								            <tr>
								                <td align="center" colspan="4" style="width:100%">
								                    <table cellpadding="0" cellspacing="1" width="100%" class="ExcelTable">
								                        <tr>
								                            <td class="ExcelHeaderCell" align="center">
								                                Invoice No
								                            </td>
								                            <td class="ExcelHeaderCell" align="center">
								                                Invoice Date
								                            </td>
								                            <td class="ExcelHeaderCell" align="center">
								                                Invoice Amount 
								                            </td>
								                            <td class="ExcelHeaderCell" align="center">
								                                T.D.S
								                            </td>
								                            <td class="ExcelHeaderCell" align="center">
								                                Deduction Advance 
								                            </td>
								                            <td class="ExcelHeaderCell" align="center">
								                                D.Note
								                            </td>
								                            <td class="ExcelHeaderCell" align="center">
								                                Paid Amount 
								                            </td>
								                            <td class="ExcelHeaderCell" align="center">
								                                Remarks
								                            </td>
								                        </tr>
								                        <%
								                            sQuery = "Select PayablesNumber from Acc_T_CreatedPybleAdjdet where CreatedTransNo = " & iTransNo &" order by 1 "
	                                                            dcrs.Open sQuery,con
	                                                            Do while not dcrs.EOF 
		                                                            iPayNo = dcrs(0)
                                                            		 
		                                                            sQuery = "Select isNull(PartyBillNumber,''),convert(VarChar,isnull(PartyBillDate,''),106),CreatedTransNo from Acc_T_CreatedPayables where PayablesNumber = " & iPayNo
		                                                            dcrs1.Open sQuery,con
		                                                            If not dcrs1.EOF then	
			                                                            sParBillNo		= dcrs1(0)
			                                                            dtParBillDate	= replace(dcrs1(1)," ","-")
			                                                            pCrTransNo		= dcrs1(2)
		                                                            End If
		                                                            dcrs1.Close
                                                            		
		                                                            sQuery = "Select VoucherNumber,convert(VarChar,VoucherDate,106),VoucherAmount from Acc_T_VoucherHeader where CreatedTransNo = " & pCrTransNo
		                                                            dcrs1.Open sQuery,con
		                                                            If not dcrs1.EOF then	
			                                                            iVouNo		= dcrs1(0)
			                                                            dtVouDate	= replace(dcrs1(1)," ","-")
			                                                            iVouAmt		= dcrs1(2)
		                                                            End If
		                                                            dcrs1.Close
		                                                            iTotVouAmt = CDbl(iTotVouAmt) + CDbl(iVouAmt)
		                                                            'Debit note column
		                                                            sQuery = "Select AmountPaid from Acc_T_CreatedPybleAdjdet where PayablesNumber =  " & iPayNo & " and CreatedTransNo = " & iTransNo
                                                            	 
		                                                            dcrs1.Open sQuery,con
		                                                            If not dcrs1.EOF then	
			                                                            iAmtPaid = dcrs1(0)
		                                                            End If
		                                                            dcrs1.Close 
		                                                            iTotAmtPaid  = cdbl(iTotAmtPaid) + cdbl(iAmtPaid)
		                                                            %>
		                                                                <tr>
		                                                                    <td class="ExcelDisplayCell">
		                                                                        <%=sParBillNo%>
		                                                                    </td>
		                                                                    <td class="ExcelDisplayCell">
		                                                                        <%=dtParBillDate%>
		                                                                    </td>
		                                                                    <td class="ExcelDisplayCell">
		                                                                        <%=formatNumber(iVouAmt,2,,,0)%>
		                                                                    </td>
		                                                                    <td class="ExcelDisplayCell">
		                                                                        <%="0.00"%>
		                                                                    </td>
		                                                                    <td class="ExcelDisplayCell">
		                                                                        <%="0.00"%>
		                                                                    </td>
		                                                                    <td class="ExcelDisplayCell">
		                                                                        <%=formatNumber(iAmount,2,,,0)%>
		                                                                    </td>
		                                                                    <td class="ExcelDisplayCell">
		                                                                        <%=formatNumber(iAmtPaid,2,,,0)%>
		                                                                    </td>
		                                                                    <td class="ExceldisplayCell">
		                                                                    </td>
		                                                                </tr>    
		                                                            <%
		                                                            sPara = "D"
		                                                            dcrs.MoveNext
	                                                            loop
	                                                            dcrs.Close
								                        
								                        %>
								                    </table>
								                </td>
								            </tr>
								            <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                         <tr>
					                            <td style="width:25%" class="FieldCellSub">
					                                <b>Total Paid Amt  : </b>
					                            </td>
					                            <td style="width:25%" class="FieldCellSub" colspan="3">
					                                <b><%=FormatNumber(iHeadAmount,2,,0)%></b>
					                            </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                         <tr>
					                            <td style="width:25%" class="FieldCellSub">
					                                <b>Amount in Words : </b>
					                            </td>
					                            <td style="width:25%" class="FieldCellSub" colspan="3">
					                                <b><%=amountwords(iHeadAmount)%></b>
					                            </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                        <tr>
					                            <td style="width:25%" class="FieldCellSub">
					                                <b>Kindly Acknowledge</b>
					                            </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                        <tr>
					                            <td style="width:25%" class="FieldCellSub" colspan="2">
					                                <b>For K.Sivasubramaniam Spinners (P) Ltd</b>
					                            </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                        <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                        <tr>
					                            <td style="width:25%" class="FieldCellSub" colspan="2" align="right">
					                                <b>V.Venkatesh</b>
					                            </td>
					                        </tr>
					                        <tr>
					                            <td style="width:25%" class="FieldCellSub" colspan="2" align="right">
					                                <b>Authorised Signatory</b>
					                            </td>
					                        </tr>
					                         <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                         <tr>
						                        <td align="center" colspan="4" class="MiddlePack">
						                        </td>
					                        </tr>
					                         <tr>
					                            <td class="FieldCellSub" colspan="4">
					                                <b>Note : We request you to send the cheque For clearing to Tirupur branch</b>
					                            </td>
					                        </tr>
								        </table>
								    </td>
								    <td align="center" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>
								<tr>
									<td align="center" width="5">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
													<input type="button" value="Close" name="B3" class="ActionButton" onclick="window.close()">
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="3" class="BottomPack">
									</td>
								</tr>

							</table>
						</td>
					</tr>

				</table>
			</td>
		</tr>

	</table>
	</form>
</body>
