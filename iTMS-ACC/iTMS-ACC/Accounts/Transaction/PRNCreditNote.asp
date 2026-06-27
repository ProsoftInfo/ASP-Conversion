<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PRNCreditNote.asp
	'Module Name				:	Accounts (Reports)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	29 March 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/PrintFunctions.asp"-->

<%
'------------------------Declaration Constants -----------------------------
dim sPgPitch,sPrFooter,sPgMargin,sPgBreak,iPgLineNo,iRecCount,sDisplayHead
dim aiHeadColWidth(0,7),aiUtilColWidth(3,8),sPagetitle2,sTstr,objFSO,objTxt
dim oDOM,RootNode,VoucherNode,PartyNode,sOrgID,sVouDate,sVouNo,sPartyCode,sExp
dim iTraNo,sBlankLine
dim iPageLen,iActualpgLen,iLineNo,sFinalText

iTraNo = Request("iTraNo")



' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set RootNode = oDOM.documentElement

if oDOM.Load(server.MapPath("../xmldata/Voucher/"&iTraNo&".xml")) then
	Set RootNode = oDOM.documentElement

	sExp="//voucher"
	set VoucherNode = RootNode.Selectnodes(sExp)
	sOrgID = trim(VoucherNode.Item(0).Attributes.getNamedItem("UnitNo").Value)
	sVouDate = trim(VoucherNode.Item(0).Attributes.getNamedItem("VouDate").Value)
	sVouNo = trim(VoucherNode.Item(0).Attributes.getNamedItem("VoucherNo").Value)
	'iTraNo = trim(VoucherNode.Item(0).Attributes.getNamedItem("TransNo").Value)
	
	sExp="//voucher/Party"
	set PartyNode = RootNode.Selectnodes(sExp)
	sPartyCode = trim(PartyNode.Item(0).Attributes.getNamedItem("ParCode").Value)
	
end if

iPageLen = formattprint("PAGELEN6LN","")
iActualpgLen = formattprint("PAGELEN12LN","")

sPgPitch = formattprint("10PITCH","")
sPgBreak = formattprint("PAGESKIP","")
sBlankLine = formattprint("BLANKLINE","")
iPgLineNo = 64

'WIDTH SPECIFICATION FOR PAGE TITLE LINE 1 AND 2
aiHeadColWidth(0,0)=5
aiHeadColWidth(0,1)=10
aiHeadColWidth(0,2)=25
aiHeadColWidth(0,3)=10
aiHeadColWidth(0,4)=10
aiHeadColWidth(0,5)=4
aiHeadColWidth(0,6)=8
aiHeadColWidth(0,7)=11

'WIDTH SPECIFICATION FOR OPENING/CLOSING  LINE 
aiUtilColWidth(0,0)=5
aiUtilColWidth(0,1)=10
aiUtilColWidth(0,2)=25
aiUtilColWidth(0,3)=10
aiUtilColWidth(0,4)=11
aiUtilColWidth(0,5)=4
aiUtilColWidth(0,6)=8
aiUtilColWidth(0,7)=13

'------------------------End of Declaration Constants ----------------------
%>
<%
dim dcrs,dcrs1,sQuery,sGetVal,saTemp,sAcHead,sAcCode,isNo
dim iPageNo,sAcDesc,dIncSum,dExpSum
dim iTransNo,sAccDescription,sOrgName,sAccHeadDesc
dim iVocEntryNo,sAccUnitPartyCode,iAccHead
dim sVovNarration,sTransCrDrId,dAmount,sUnitAcCode,sTransType,iVouEntNo
dim iCtr,arrTemp,sInvDate,sInvNo,arrTemp1,sUnitName
isNo=0
iPageNo=1

set dcrs  = server.CreateObject("adodb.recordset")
set dcrs1  = server.CreateObject("adodb.recordset")

set objFSO = Server.CreateObject("Scripting.FileSystemObject")
set objTxt = objFSO.CreateTextFile(server.MapPath("../temp/Reports/"&iTraNo&"_CreditNote.txt"))

sTStr = ConsPgHeader()

with dcrs1
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = con
	.Source = "SELECT ACCOUNTINGUNIT,ISNULL(ACCUNITACCOUNTHEAD,0),VOUCHERNARRATION,AMOUNT,TRANSCRDRINDICATION FROM ACC_T_CreatedVOUCHERDETAILS WHERE CreatedTransNo = " & iTraNo &""
	.Open 
End with
iRecCount = dcrs1.RecordCount

set dcrs1.ActiveConnection =nothing

dExpSum=0
dIncsum=0

do while not dcrs1.EOF 
	sUnitAcCode	= left(trim(dcrs1(0)),2)
	iAccHead = trim(dcrs1(1))
	sVovNarration = trim(dcrs1(2))
	dAmount	= trim(dcrs1(3))
	sTransCrDrId = trim(dcrs1(4))
	'if sVovNarration <> "" then
	'	arrTemp = split(sVovNarration,":")
	'	arrTemp1 = split(arrTemp(1),"-")
	'	sInvDate = arrTemp1(1)
	'	sInvNo = arrTemp1(0)
	'else
	'	sInvDate = " "
	'	sInvNo = " "
	'end if

	With dcrs
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = "SELECT ACCOUNTHEADCODE,ACCOUNTDESCRIPTION FROM ACC_M_GLACCOUNTHEAD WHERE ACCOUNTHEAD = " & iAccHead
		.Open 
	End with
	set dcrs.ActiveConnection =nothing
	If not dcrs.EOF Then
		sAcCode = left(trim(dcrs(0)),8)
		sAcDesc = trim(dcrs(1))
	End if
	dcrs.Close 

	If cdbl(dAmount) <> 0 then
		sTStr = sTStr &myalign(sUnitAcCode,aiUtilColWidth(0,0),"L")	
		sTStr = sTStr &myalign(sAcCode,aiUtilColWidth(0,1),"L")	
		sTStr = sTStr &myalign(sAcDesc,aiUtilColWidth(0,2),"L")	
		sTStr = sTStr &myalign(sInvNo,aiUtilColWidth(0,3),"L")	
		sTStr = sTStr &myalign(sInvDate,aiUtilColWidth(0,4),"L")	
		sTStr = sTStr &myalign(" ",aiUtilColWidth(0,5),"L")	
		sTStr = sTStr &myalign(sTransCrDrId,aiUtilColWidth(0,6),"C")	
		If sTransCrDrId="C" Then
			dExpSum=cdbl(dExpSum)+cdbl(dAmount)
			sTStr = sTStr &myalign(FormatNumber(dAmount,2,,,0),aiUtilColWidth(0,7),"R")	
		Else
			dIncSum=cdbl(dIncSum)+cdbl(dAmount)
			sTStr = sTStr &myalign(FormatNumber(dAmount,2,,,0),aiUtilColWidth(0,7),"R")	
		End IF
	End if
	sTstr = sTstr &vbCrLf 
	iLineNo = iLineNo + 1
	CheckNew
dcrs1.MoveNext 
loop 
dcrs1.Close 

for iCtr = iLineNo to 20 
	sTStr=sTStr & string(83," ") & vbCrLf
	iLineNo = iLineNo + 1
next

sTStr=sTStr & string(83,"-") & vbCrLf
sTstr = sTstr &myalign("NET AMOUNT : ",70,"R") & myalign(FormatNumber((dIncSum - dExpSum),2,,,0),aiUtilColWidth(0,7),"R") & vbCrLf
sTStr=sTStr & string(83,"-") & vbCrLf

sTStr = sTStr & "AMOUNT : (" & AmountWords((dIncSum - dExpSum)) & ")" & vbCrLf
sTstr = sTstr & sBlankLine

sTStr = sTStr & myalign("For ",55,"R") & sUnitName & vbCrLf
sTstr = sTstr & sBlankLine
sTstr = sTstr & sBlankLine

sTStr = sTStr & myalign("Authorised Signatory",75,"R") & vbCrLf

sFinalText = sFinalText & sTstr		
objTxt.write sFinalText			

objTxt.Close

function ConsPgHeader()
	dim sTStr,i,sPartAddr,sAccountOf,sUnitAddr

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ORGUNITDESCRIPTION,RTRIM(ADDRESS1)+','+RTRIM(ADDRESS2)+','+RTRIM(CITY)+','+RTRIM(STATE)+' - '+RTRIM(STR(POSTCODE,6,0)) FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID = " & Pack(sOrgID) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		sUnitName = trim(dcrs(0))
		sUnitAddr = myalign(trim(dcrs(1)),83,"C")
	end if
	dcrs.Close

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT PARTYNAME,ADDRESSLINE1,ADDRESSLINE2,RTRIM(CITY)+','+RTRIM(STATE)+','+RTRIM(COUNTRY)+' - '+RTRIM(STR(PINCODE,6,0)) FROM APP_M_PARTYMASTER WHERE PARTYCODE = " & Pack(sPartyCode) & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		sPartAddr = string(5," ") & trim(dcrs(0)) & vbCrLf
		sPartAddr = sPartAddr & string(5," ") & trim(dcrs(1)) & vbCrLf
		sPartAddr = sPartAddr & string(5," ") & trim(dcrs(2)) & vbCrLf
		sPartAddr = sPartAddr & string(5," ") & trim(dcrs(3)) 
	end if
	dcrs.Close

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT PAYTORECDFROM FROM Acc_T_CreatedVoucherHeader WHERE CreatedTransNo = " & iTraNo & ""
		.ActiveConnection = con
		.Open
	end with
	'Response.Write dcrs.Source
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		sAccountOf = trim(dcrs(0))
	end if
	dcrs.Close

	sTStr=sTStr & formattprint("RESET","")
	sTStr=sTStr & sPgPitch
	sTStr=sTStr & iPageLen
	sTStr=sTStr & formattprint("BLANKLINE","")
	
	sTStr= myalign(sUnitName,83,"C") & vbCrLf
	sTStr=sTStr & sUnitAddr & vbCrLf

	sTStr=sTStr & myalign("CREDIT NOTE",45,"R") & vbCrLf

	sTStr=sTStr & myalign("VOUCHER NO   : ",65,"R") & sVouNo & vbCrLf
	sTStr=sTStr & myalign("VOUCHER DATE : ",65,"R") & sVouDate & vbCrLf

	sTStr=sTStr & "To" & vbCrLf
	sTStr=sTStr & sPartAddr & vbCrLf
		
	sTStr = sTStr & "BEING CREDITED TOWARDS YOU ON ACCOUNT OF :"& vbCrLf
	sTStr=sTStr & ucase(sAccountOf) & vbCrLf

	sTStr=sTStr & string(83,"-") & vbCrLf
	
	sTStr=sTStr & myalign("AU",aiHeadColWidth(0,0),"L")
	sTStr=sTStr & myalign("A/C CODE",aiHeadColWidth(0,1),"L")
	sTStr=sTStr & myalign("ACCOUNT NAME",aiHeadColWidth(0,2),"L")
	sTStr=sTStr & myalign("BILL NO.",aiHeadColWidth(0,3),"L")
	sTStr=sTStr & myalign("BILL DATE",aiHeadColWidth(0,4),"L")
	sTStr=sTStr & myalign("CC",aiHeadColWidth(0,5),"R")
	sTStr=sTStr & myalign("DR/CR",aiHeadColWidth(0,6),"R")
	sTStr=sTStr & myalign("AMOUNT",aiHeadColWidth(0,7),"R")&"  "
	sTStr=sTStr & vbCrLf
	
	sTStr=sTStr & string(83,"-") & vbCrLf
	iLineNo = 16

	'----------------------End of Page Header Construction----------------------
	ConsPgHeader=sTStr

end function

Function CheckNew
	if iLineNo>=21 then
		sTStr=sTStr & string(83,"-") & vbCrLf
		sTstr = sTstr &myalign("NET AMOUNT : ",70,"R") & myalign(FormatNumber((dIncSum - dExpSum),2,,,0),aiUtilColWidth(0,7),"R") & vbCrLf
		sTStr=sTStr & string(83,"-") & vbCrLf

		sTStr = sTStr & "AMOUNT : (" & AmountWords((dIncSum - dExpSum)) & ")" & vbCrLf
		sTstr = sTstr & sBlankLine

		sTStr = sTStr & myalign("For ",55,"R") & sUnitName & vbCrLf
		sTstr = sTstr & sBlankLine
		sTstr = sTstr & sBlankLine

		sTStr = sTStr & myalign("Authorised Signatory",75,"R") & vbCrLf
		sFinalText = sFinalText & sTstr & chr(12)
		sTstr = ConsPgHeader()
	end if	
end Function

'------------------------End Of User Define Function------------------------

if cint(iRecCount)>=0 then
	Response.Write iRecCount &" ======= "
	'Response.Redirect("../../Components/FormattPrint.asp?server=server&filepath=/accounts/temp/Reports/"&iTraNo&"_CreditNote.txt&exitpath=/accounts/reports/VouCNBookSelection.asp&frame=_parent")
	Response.Redirect("../../Components/FormattPrint.asp?server=server&filepath=/accounts/temp/Reports/"&iTraNo&"_CreditNote.txt&exitpath=/accounts/reports/VouCNBookSelection.asp&frame=_parent")
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
        <FONT SIZE="-1" FACE="Arial,Helvetica,Univers,Zurich BT"><a href="javascript:window.history.back(1)">Back</a></FONT></B></TD>
    </TR>
</Table>
</HTML>

<%end if%>

