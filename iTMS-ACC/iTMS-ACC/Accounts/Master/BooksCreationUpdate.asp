<%@ Language=VBScript %>
<%	option explicit	%>
<%'	on Error Resume Next%>
<%
	'Program Name				:	BooksCreationUpdate.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	SENTHIL E
	'Created On					:	December 06, 2002
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
<!--#include file="../../include/NoSeries.asp"-->
<%
dim iUnitNo,sUnitName,iBookNo,iBookId,sBookName,iCounter
dim sQuery,objRs,iExistBookNo,sOtherUnitEligible
dim iSeries,iSeriesType,bPayRecNo,iLength,sFinPer,objFSO
'XML DOM Variables
Dim oDOM,nodUnit,Root,nodBook
sFinPer = Session("FinPeriod")

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = Server.CreateObject("ADODB.RecordSet")
set objFSO = Server.CreateObject("Scripting.FileSystemObject")
'Response.Write server.MapPath("../../accounts/temp/master/Unit_Book_"& Session.SessionID &".xml")
oDOM.Load server.MapPath("../../accounts/temp/master/Unit_Book_"& Session.SessionID &".xml")

Set Root = oDOM.documentElement

iUnitNo=trim(Request.Form("hUnitId"))
iBookId=trim(Request.Form("selDayBook"))
iSeries=trim(Request.Form("selNoSeries"))
iSeriesType=trim(Request.Form("hSeriesType"))
bPayRecNo=trim(Request.Form("selPayRecNo"))
iLength=trim(Request.Form("hSeriesLen"))
sBookName=Replace(trim(Request.Form("txtName")),"'","''")
sOtherUnitEligible=trim(Request.Form("optEligible"))

sQuery="select count(0)from  Acc_R_ApplicableAccountHeads where OUDefinitionID='"&iUnitNo&"' and "&_
"BookCode='"&iBookId&"' and upper(BookName)='"&UCase(sBookName)&"' and Useable = '0' "

with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with

if objRs(0)>0 then
%>

<HTML>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr)
	{
			alert(strr);
			window.history.back();
	}
//-->
</SCRIPT>
<BODY onLoad = "msgbox('Book Name already Exist')">
</BODY>
<HTML>
<%
Response.End

end if
objRs.Close

con.BeginTrans

For Each nodUnit In Root.childNodes
	if nodUnit.Attributes.Item(0).nodeValue = iUnitNo then
		For Each nodBook In nodUnit.childNodes
			if nodBook.Attributes.Item(0).nodeValue=iBookId then
				iExistBookNo=nodBook.Attributes.Item(1).nodeValue
				
				sQuery = "Select BookCode from Acc_M_OrgBooks where BookCode = '"& iBookId &"' and OUDefinitionID ='"& iUnitNo &"'"
				objRs.Open sQuery,con
				if not objRs.EOF then
				    sQuery="Update Acc_M_OrgBooks set NumberOfBooks=+1 where OUDefinitionID='"&iUnitNo&"'"&_
						" and BookCode='"&iBookId&"'"
				else
				    sQuery="Insert into Acc_M_OrgBooks (OUDefinitionID,BookCode,NumberOfBooks) values('"& iUnitNo &"','" &iBookId & "',1)"
				end if
				objrs.Close 
				nodBook.Attributes.Item(1).nodeValue=cint(iExistBookNo)+1
				iBookNo	=cint(iExistBookNo)+1
				Response.Write "<p>"& sQuery
				con.Execute (sQuery)
			end if
		next
	end if
next

sQuery = "Select isNull(Max(BookNumber),0)+1 From Acc_R_ApplicableAccountHeads Where " & _
		 "OUDefinitionID = '"&iUnitNo&"' and BookCode = '"&iBookId&"' "
Response.Write "<p>"& sQuery
objRs.OPen sQUery,COn
IF Not objRs.Eof Then
	iBookNo = objRs(0)
End IF
objRs.close


sQuery="insert into Acc_R_ApplicableAccountHeads (OUDefinitionID,BookCode,BookNumber,"&_
		"BookName,OtherUnitTransaction,BookAccountHead) values ('"&iUnitNo&"','"&iBookId&"',"&_
		""&iBookNo&",'"&sBookName&"','"&sOtherUnitEligible&"',NULL)"

Response.write "<p>"& sQuery
	con.Execute (sQuery)

dim iCRCode,iDRCode,iCreatedCRCode,iCreatedDRCode

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

sQuery="INSERT INTO Acc_M_BookNumberSeries(OUDefinitionID, BookCode, BookNumber, DrSeriesNo, "&_
	"DrSeriesCode, CrSeriesNo, CrSeriesCode,CreatedDrSeriesNo,CreatedDrSeriesCode, CreatedCrSeriesNo,"&_
	"CreatedCrSeriesCode, LastChequeNo, ReceiptNo,FinPeriod) "&_
	"VALUES('"&iUnitNo&"','"&iBookId&"',"&iBookNo&","&iSeries&","&_
	""&iDRCode&","&iSeries&","&iCRCode&","&iSeries&","&iCreatedDRCode&","&iSeries&","&_
	""&iCreatedCRCode&",1,NULL,'"&sFinPer&"'	)"

Response.Write "<p>"& sQuery &"<br><br>"
con.Execute (sQuery)



if con.Errors.count <>0 then
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else
   ' con.rollbacktrans
    'Response.end 
	con.CommitTrans
	if objFSO.FileExists(server.MapPath("../temp/master/Unit_Book_"& Session.SessionID &".xml")) then
	    objFSO.DeleteFile(server.MapPath("../temp/master/Unit_Book_"& Session.SessionID &".xml"))
	end if
	
	Response.Clear
%>
<HTML>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			if (confirm("Do You want to Create Book"))
				window.location.href = "BooksCreationEntry.asp";
			else
				window.location.href = "DayBookGrid.asp";
		}
		else {
			alert(strr);
			window.location.href = "BooksCreationEntry.asp";
		}
	}
//-->
</SCRIPT>
	<BODY onLoad = "msgbox('Book Created Successfully','Y')">

</BODY>
<HTML>
<%end if%>

