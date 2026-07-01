<%@ Language=VBScript %>
<%	option explicit	%>
<%'	on Error Resume Next%>
<%
	'Program Name				:	BooksEditPopupUpdate.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 10,2010
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
	<base target="_self">
<script language="javascript">
window.__itmsPopupCompat = {
	type: "autoClose",
	message: "Data Updated Successfully",
	returnValue: "Done"
};
</script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</head>

<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/NoSeries.asp"-->
<%
dim iUnitNo,sUnitName,iBookNo,iBookId,sBookName,iCounter,Temparr
dim sQuery,objRs,iExistBookNo,sOtherUnitEligible
dim iSeries,iSeriesType,bPayRecNo,iLength,sCallTy,sEditTy,iRowCnt
dim bEntryFlag

'XML DOM Variables
Dim oDOM,nodUnit,Root,nodBook

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = Server.CreateObject("ADODB.RecordSet")
oDOM.Load server.MapPath("../xmldata/UnitBookDetails.xml")

Set Root = oDOM.documentElement

' iUnitNo=trim(Request.QueryString("hOrgCode"))
' iBookId=trim(Request.QueryString("hBookCode"))
' iSeries=trim(Request.QueryString("selNoSeries"))
' iSeriesType=trim(Request.QueryString("hSeriesType"))
' bPayRecNo=trim(Request.QueryString("selPayRecNo"))
' iLength=trim(Request.QueryString("hSeriesLen"))
' sBookName=Replace(trim(Request.QueryString("txtName")),"'","''")
' sOtherUnitEligible=trim(Request.QueryString("optEligible"))
' iBookNo = Request.QueryString("hBookNo")

' sCallTy = Request.QueryString("hCallType")
' sEditTy = Request.QueryString("hEditType")

' iRowCnt = Request.QueryString("hRowCnt")
Response.write "<font color=red>"

iUnitNo=trim(Request.Form("hOrgCode"))
iBookId=trim(Request.Form("hBookCode"))
iSeries=trim(Request.Form("selNoSeries"))
iSeriesType=trim(Request.Form("hSeriesType"))
bPayRecNo=trim(Request.Form("selPayRecNo"))
iLength=trim(Request.Form("hSeriesLen"))
sBookName=Replace(trim(Request.Form("txtName")),"'","''")
sOtherUnitEligible=trim(Request.Form("optEligible"))
iBookNo = Request.Form("hBookNo")

sCallTy = Request.Form("hCallType")
sEditTy = Request.Form("hEditType")

iRowCnt = Request.Form("hRowCnt")


IF CInt(iRowCnt) > 10 Then
	iSeriesType = "M"
Elseif CInt(iRowCnt) >= 1 and CInt(iRowCnt) <= 2 Then
	iSeriesType = "Y"
Elseif CInt(iRowCnt) >= 4 and CInt(iRowCnt) <= 5 Then
	iSeriesType = "Q"
ENd IF

'Response.Write iSeriesType
'Response.End

'Response.Write sCallTy &" " & sEditTy


'Response.Write Request.Form("txtPrefix1") &"<br>"

Con.BeginTrans
bEntryFlag = "N"
IF CStr(sCallTy) = "E" Then


	sQuery = "UPDATE Acc_R_ApplicableAccountHeads SET BookName = '"&sBookName&"' , OtherUnitTransaction = '"&sOtherUnitEligible&"' "&_
			 "WHERE OUDefinitionID = '"&iUnitNo&"' AND BookCode = '"&iBookId&"' AND BookNumber = "&iBookNo&" "

	con.Execute (sQuery)

	Dim iDrNo,iDrCode,iCrNo,iCrCode,iCDrNo,iCDrCode,iCCrNo,iCCrCode

	sQuery = "Select DrSeriesNo,DrSeriesCode,CrSeriesNo,CrSeriesCode,CreatedDrSeriesNo, "&_
			 "CreatedDrSeriesCode,CreatedCrSeriesNo,CreatedCrSeriesCode From "&_
			 "Acc_M_BookNumberSeries Where OUDefinitionID = '"&iUnitNo&"' and BookCode = '"&iBookId&"' "&_
			 "and BookNumber = "&iBookNo&" "

	'Response.Write sQuery

	Objrs.Open sQuery,Con
	IF Not Objrs.Eof Then
		iDrNo = Objrs(0)
		iDrCode = Objrs(1)
		iCrNo = Objrs(2)
		iCrCode = Objrs(3)
		iCDrNo = Objrs(4)
		iCDrCode = Objrs(5)
		iCCrNo = Objrs(6)
		iCCrCode = Objrs(7)
		bEntryFlag ="Y"
	else
	    bEntryFlag ="N"

	End IF
	Objrs.close
if trim(bEntryFlag)="Y" then
	sQuery = "Delete APP_R_NoSeriesModuleEntry Where SeriesNo = "&iDrNo&" and SeriesCode = "&iDrCode&" "&_
			 "and OUDefinitionID = '"&iUnitNo&"' "
	'		 Response.Write sQuery

	Con.Execute sQuery


	sQuery = "Delete APP_R_NoSeriesModuleEntry Where SeriesNo = "&iCrNo&" and SeriesCode = "&iCrCode&" "&_
			 "and OUDefinitionID = '"&iUnitNo&"' "
	'	 Response.Write sQuery
	Con.Execute sQuery


	sQuery = "Delete APP_R_NoSeriesModuleEntry Where SeriesNo = "&iCDrNo&" and SeriesCode = "&iCDrCode&" "&_
			 "and OUDefinitionID = '"&iUnitNo&"' "

	Con.Execute sQuery

	sQuery = "Delete APP_R_NoSeriesModuleEntry Where SeriesNo = "&iCCrNo&" and SeriesCode = "&iCCrCode&" "&_
			 "and OUDefinitionID = '"&iUnitNo&"' "

	Con.Execute sQuery

	'=========================================================================================
	sQuery = "Delete APP_R_NoSeriesModules Where SeriesNo = "&iDrNo&" and SeriesCode = "&iDrCode&" "&_
			 "and OUDefinitionID = '"&iUnitNo&"' "

	Con.Execute sQuery

	sQuery = "Delete APP_R_NoSeriesModules Where SeriesNo = "&iCrNo&" and SeriesCode = "&iCrCode&" "&_
			 "and OUDefinitionID = '"&iUnitNo&"' "

	Con.Execute sQuery

	sQuery = "Delete APP_R_NoSeriesModules Where SeriesNo = "&iCDrNo&" and SeriesCode = "&iCDrCode&" "&_
			 "and OUDefinitionID = '"&iUnitNo&"' "

	Con.Execute sQuery

	sQuery = "Delete APP_R_NoSeriesModules Where SeriesNo = "&iCCrNo&" and SeriesCode = "&iCCrCode&" "&_
			 "and OUDefinitionID = '"&iUnitNo&"' "

	Con.Execute sQuery
	'=========================================================================================
	sQuery = "Delete Acc_M_BookNumberSeries Where BookCode = '"&iBookId&"' and BookNumber = "&iBookNo&" "&_
			 "and OUDefinitionID = '"&iUnitNo&"' "

	Con.Execute sQuery
	'=========================================================================================
end if 'if trim(bEntryFlag)="Y" then

	dim iCreatedCRCode,iCreatedDRCode

	if bPayRecNo="Y" then
		iCRCode= GenSeriesCode(iUnitNo,"1","1",iSeries,iSeriesType,"Cr","Payment Voucher No-"&sBookName,iLength)
		iDRCode=GenSeriesCode(iUnitNo,"1","1",iSeries,iSeriesType,"Dr","Receipt Voucher No-"&sBookName,iLength)
		iCreatedCRCode=GenSeriesCode(iUnitNo,"1","1",iSeries,iSeriesType,"Cr","Created Payment Voucher No-"&sBookName,iLength)
		iCreatedDRCode=GenSeriesCode (iUnitNo,"1","1",iSeries,iSeriesType,"Dr","Created Receipt Voucher No-"&sBookName,iLength)
	else
		iCRCode=GenSeriesCode(iUnitNo,"1","1",iSeries,iSeriesType,"","Payment/Receipt Voucher No-"&sBookName,iLength)
		iCreatedCRCode=GenSeriesCode(iUnitNo,"1","1",iSeries,iSeriesType,"","Created Payment/Receipt Voucher No-"&sBookName,iLength)
		iDRCode=iCRCode
		iCreatedDRCode=iCreatedCRCode
	end if

	sQuery = "INSERT INTO Acc_M_BookNumberSeries(OUDefinitionID, BookCode, BookNumber, DrSeriesNo, "&_
			 "DrSeriesCode, CrSeriesNo, CrSeriesCode,CreatedDrSeriesNo,CreatedDrSeriesCode, CreatedCrSeriesNo,"&_
			 "CreatedCrSeriesCode, LastChequeNo, ReceiptNo,FinPeriod) "&_
			 "VALUES('"&iUnitNo&"','"&iBookId&"',"&iBookNo&","&iSeries&","&_
			 ""&iDRCode&","&iSeries&","&iCRCode&","&iSeries&","&iCreatedDRCode&","&iSeries&","&_
			 ""&iCreatedCRCode&",1,NULL,'"&Session("FinPeriod")&"'	)"

'	Response.write sQuery

	con.Execute (sQuery)

End IF

IF CStr(sCallTy) = "D" Then
	sQuery = "UPDATE Acc_R_ApplicableAccountHeads SET  Useable = '1' "&_
			 "WHERE BookCode = '"&iBookId&"' AND BookNumber = "&iBookNo&" AND OUDefinitionID = '"&iUnitNo&"' "

'	Response.Write sQuery

	Con.Execute sQuery
End IF
Con.CommitTrans
'Response.Redirect "BooksEditEntryPopup.asp?OrgCode="&iUnitNo&"&BookCode="&iBookId&"&BookNumber="&iBookNo
'Con.RollbackTrans
'Response.end
'Response.Redirect "DAYBOOKGRID.ASP"
%>
<body onLoad="init()">
</body>
</html>



