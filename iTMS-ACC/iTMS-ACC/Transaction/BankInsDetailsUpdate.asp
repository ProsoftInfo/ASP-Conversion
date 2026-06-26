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
	'Program Name				:	BankInsDetailsUpdate.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	S.Maheswari
	'Created On					:	June 13, 2008
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<%

dim sQuery,sorgID,iBookNo,bActionFlag
dim svalue
'sValue = Request.QueryString("Value")

Con.BeginTrans
'*********************************Added by Maheswari on June 13th 2008 *******************************************

dim oDOM,Root,InsNode,objRs,iEntNo,objFS
dim sUnitId,iBookId,sDrawnOn,sPayAt,iStartNo,iEndNo,dtIssueDate,sStatus
Dim sInsEntNo,i,iCtr
set objRs  = server.CreateObject("ADODB.recordset")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objFS = Server.CreateObject("Scripting.FileSystemObject")
iCtr = 1
oDOM.Load server.MapPath("../temp/transaction/BankInsDet_BA_"&Session.SessionID&".xml")	
set Root=oDOM.documentElement
IF Root.haschildnodes then 
	For Each InsNode in Root.childnodes 
		IF trim(InsNode.NodeName) = "InstrumentDetails" then 
			iEntNo		= InsNode.getAttribute("EntryNo")
			sUnitId 	= InsNode.getAttribute("UnitId")
			iBookId 	= InsNode.getAttribute("BookId")
			sDrawnOn	= InsNode.getAttribute("DrawnOn")
			sPayAt		= InsNode.getAttribute("PayAt")
			iStartNo	= InsNode.getAttribute("StartNo")
			iEndNo		= InsNode.getAttribute("EndNo")
			dtIssueDate = InsNode.getAttribute("IssuedOn")
			sStatus		= InsNode.getAttribute("Status")
		End IF
	Next
End IF

IF trim(iEntNo) <>  ""  then 
	'Update Into Acc_R_BankInstrumentDetails -ITMS-Sangeeth_2008

	sQuery = " Update Acc_R_BankInstrumentDetails Set DrawnOn = '"&sDrawnOn&"',PayableAt='"&sPayAt&"', "&_
			 " StartNo = '"&iStartNo&"',EndNo = '"&iEndNo&"',DateOfIssue=Convert(DateTime,'"&dtIssueDate&"',103),"&_
			 " Status = '"&sStatus&"' where EntryNo= "&iEntNo&" and OUDefinitionID = '"&sUnitId &"' and BookNumber = "&iBookId&" "
			 
	Response.Write sQuery	&"<BR><BR>"
	con.execute sQuery 		
	
	sQuery = "Delete from Acc_R_BankInstrumentUsage where EntryNo = "&iEntNo&" "
	Response.Write sQuery	&"<BR><BR>"
	con.execute sQuery 		
Else
	sQuery = "Select isNull(Max(EntryNo),0)+1 from Acc_R_BankInstrumentDetails"
	objRs.Open sQuery,con
	IF not objRs.EOF then 
		iEntNo = objRs(0)
	Else
		iEntNo = 1
	End If
	objRs.close
	IF trim(iEntNo) = "" then iEntNo = 1
	Response.Write "ientryno="&iEntNo 
	'Insert Into Acc_R_BankInstrumentDetails - ITMS-Sangeeth_2008

	sQuery = " Insert into Acc_R_BankInstrumentDetails(EntryNo,OUDefinitionID,BookNumber,DrawnOn,PayableAt,"&_
			 " StartNo,EndNo,DateOfIssue,Status) Values("&iEntNo&",'"&sUnitId &"',"&iBookId&",'"&sDrawnOn&"', "&_
			 " '"&sPayAt&"','"&iStartNo&"','"&iEndNo&"',Convert(DateTime,'"&dtIssueDate&"',103),'"&sStatus&"')"
	Response.Write sQuery	&"<BR><BR>"
	con.execute sQuery 		
	
End IF
'*****************************************End*****************************************************
Response.Clear
'Newly Added on June 16 th to insert values into Acc_R_BankInstrumentUsage table	
 dim iLen,sChk,sAdd,sLen,j
IF IsNumeric(iStartNo) = True then 
sLen = Len(iStartNo)

	For i = cint(iStartNo) to cint(iEndNo)
'		Response.Write "<p>"&cint(iStartNo) & vbCrLf 	
		'IF Left(iStartNo,1)= 0 then 
		'	If iCtr <> 1 then iStartNo = 0&iStartNo
		'End IF		
		If iCtr <> 1 then
			iLen = Len(cint(iStartNo))
			
			IF sLen <> ilen then 
				For j = 1 to (sLen - ilen)
					sAdd = sAdd & 0
				Next
			End IF
			'Response.Write sAdd &vbCrLf 
			'Response.Write sLen &"=="&iLen&vbCrLf 
			iStartNo = sAdd & iStartNo		
		End IF		
		sQuery = " Insert into Acc_R_BankInstrumentUsage(EntryNo,InstrumentEntryNo,InstrumentNo) Values("& iEntNo &","& iCtr &",'"&iStartNo&"') "
		Response.Write sQuery &"<BR><BR>"
		con.execute sQuery 
		sAdd = ""	
		iStartNo = CInt(iStartNo)  + 1 
		
		iCtr = iCtr  + 1
	Next 
End IF
 
'Response.End 
Response.clear
Con.CommitTrans	
if objFS.fileExists(Server.MapPath("../temp/transaction/BankInsDet_BA_"&Session.SessionID&".xml")) then
	objFS.deleteFile(Server.MapPath("../temp/transaction/BankInsDet_BA_"&Session.SessionID&".xml"))
end if

%>
 