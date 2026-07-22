<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	sendToAccountsInsert.asp
	'Module Name				:	Inventory (Send Closing Stock to Accounts)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	September 24,2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	sendToAccounts.asp
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

'XML DOM Variables

Dim oDOM,nodHeader,Root,objRs,sQuery,objFSO,oDOMBook
Dim ndBookRoot,ndBookChild
dim EntryNode,HeaderNode,nodANL,newElem

dim sNarration,sAccount,sAddtional,iSno,sAmount,sTemp
dim sOrgId,sBookNo,sVouType,sVouNo,sVouName
dim sVouCode,sApprove,sVoucDate,sAccUnit
dim dTotal,sTransType,dCRAmt,dDRAmt
dim sAccType,sAccCode,sEntryType,sEntryno,iTransNo
dim sDocType,sVouStatus,sPayTo
dim iSeriesNo,iSeriesCode

dim sAccHeadCode,sFinFrom

sVouStatus = "010101" 'Crearted For Accounting to be Approved
sVouCode = "08"
sFinFrom = trim(Request.Form("hFinFrom"))

set objRs  = server.CreateObject("adodb.recordset")
set objFSO = Server.CreateObject("Scripting.FileSystemObject")

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
set oDOMBook = Server.CreateObject("Microsoft.XMLDOM")
oDOM.Load server.MapPath("../temp/transaction/Creation_GJ_"&Session.SessionID&".xml")	

set Root=oDOM.documentElement

sOrgId = trim(Root.Attributes.Item(0).nodeValue)
sBookNo = trim(Root.Attributes.Item(2).nodeValue)
sVouType = trim(Root.Attributes.Item(4).nodeValue)
sVoucDate = trim(Root.Attributes.Item(5).nodeValue)
sAccHeadCode = trim(Root.Attributes.Item(6).nodeValue)

Root.Attributes.Item(7).nodeValue = sApprove

sPayTo = trim(Root.childNodes(0).Attributes.Item(2).nodeValue)

dTotal = 0
sTransType = "GJR"

if objFSO.FileExists(server.MapPath("../XMLData/BookSetup.xml")) then
oDOMBook.load server.MapPath("../XMLData/BookSetup.xml")
set ndBookRoot = oDOMBook.documentElement
    if ndBookRoot.hasChildNodes() then
        for each ndBookChild in ndBookRoot.childNodes
            if trim(ndBookChild.nodeName)="Book" then
                sBookNo = ndBookChild.getAttribute("No")
                exit for
            end if 
        next
    end if 
else
%>
 <HTML>
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript>
	alert("Setup is not done please setup the books and continue.");
	window.location.href = "sendToAccountsDetails.asp";
</SCRIPT>	
</html>
<%
end if 


con.BeginTrans



    IF strcomp(sVouType,"D") = 0 THEN
		sQuery = "select CreatedDrSeriesNo,CreatedDrSeriesCode from Acc_M_BookNumberSeries where "&_
				 "OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
	ELSE
		sQuery = "select CreatedCrSeriesNo,CreatedCrSeriesCode from Acc_M_BookNumberSeries where "&_
				 "OUDefinitionID='"&sOrgId&"' and BookCode='"&sVouCode&"' and BookNumber= "&sBookNo
	END IF
    Response.write sQuery
	
	with objRs
		.ActiveConnection = con
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.Open 
	end with
	
	If Not objRs.EOF then
		iSeriesNo=objRs(0)
		iSeriesCode=objRs(1)
	End IF
	objRs.close()
	'con.BeginTrans
	
	if trim(iSeriesNo) ="" or IsNull(iSeriesNo) then    
	    Response.clear
	    Dim sBookName
	    sBookName = GetAccBookName(sVouCode,sBookNo)
	    Response.write "<h1>Number Series is not created for <font color=red>"& sBookName &"</font></h1>" 
	    Response.end
	end if
	IF CStr(sVouNo) = "" Then
		sVouNo=GenSeriesNumber(sOrgId,iSeriesNo,iSeriesCode,sVoucDate)
	End IF


with objRs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT ISNULL(MAX(CREATEDTRANSNO),0)+1 FROM ACC_T_CREATEDVOUCHERHEADER"
	.ActiveConnection = con
	.Open
end with
set objRs.ActiveConnection = nothing

if not objRs.EOF then
	iTransNo = trim(objRs(0))
end if
objRs.Close	

sQuery = "INSERT INTO ACC_T_CREATEDVOUCHERHEADER (CREATEDTRANSNO,OUDEFINITIONID,BOOKCODE,"&_
		"BOOKNUMBER,TRANSACTIONTYPE,PARTYTYPE,ACCOUNTHEAD,CREATEDVOUCHERNO,VOUCHERDATE,"&_
		"VOUCHERAMOUNT,CRDRINDICATION,CREATEDBY,CREATEDON,APPROVEDBY,CREATEDVOUCHSTATUS,"&_
		"FROMAPPLICATION) values(" & iTransNo & ",'" & sOrgId & "','" & sVouCode & "',"&_
		""& sBookNo &",'" & sTransType & "',NULL,NULL,'"& sVouNo &"',convert(datetime,'" & sVoucDate & "',103),"&_
		"0,NULL," & getUserid & ",convert(datetime,'" & sVoucDate & "',103),NULL,'" & sVouStatus & "','4')"

Response.Write sQuery& "<BR>"	
con.execute(sQuery)

'-----------------------PROCESS ENTRY NODES-------------------------------
FOR EACH EntryNode IN Root.childNodes
	sEntryno=EntryNode.Attributes.Item(0).nodeValue
	sAmount=EntryNode.Attributes.Item(3).nodeValue 
	sEntryType=EntryNode.Attributes.Item(1).nodeValue
	sAccUnit=EntryNode.Attributes.Item(4).nodeValue 
	if CDbl(sAmount)>0 then
    '---------PROCESS THE CHILD NODES OF ENTRIES FOR DETAIL TABLE UPDATION----
	    FOR EACH HeaderNode IN EntryNode.childNodes
		    IF HeaderNode.nodeName="AccHead" THEN
			    sAccCode=HeaderNode.Attributes.Item(0).nodeValue 
			    sAccType=HeaderNode.Attributes.Item(4).nodeValue 

			    sQuery = "UPDATE INV_T_ITEMLEDGER SET SENTTOACCOUNTS = 'S' WHERE ORGANISATIONCODE = '" & sOrgId & "' " &_
					    "AND CONVERT(DATETIME,TRANSACTIONDATE,103) >= CONVERT(DATETIME,'" & sFinFrom & "',103) " &_
					    "AND CONVERT(DATETIME,TRANSACTIONDATE,103) <= CONVERT(DATETIME,'" & sVoucDate & "',103) " &_
					    "AND (STR(ITEMCODE)+STR(CLASSIFICATIONCODE)) IN (SELECT (STR(ITEMCODE)+STR(CLASSIFICATIONCODE)) " &_
					    "FROM INV_M_ITEMORGACCOUNTHEAD WHERE ACCOUNTHEAD = " & trim(sAccCode) & " " &_
					    "AND ORGANISATIONCODE = '" & sOrgId & "')"

			    Response.Write "<br>"&sQuery& "<BR>"	
			    con.execute(sQuery)

		    END IF 'End of Check for Account head Node
		    IF 	HeaderNode.nodeName="Narration" THEN
			    sNarration=HeaderNode.text
		    END IF 'End of Check for Narration Node
	    NEXT 
    '-------------END OF PROCESSING CHILD NODES OF ENTRIES---------------------	
    '----------------------------DETAIL TABLE UPDATION-------------------------
	    IF StrComp(sAccType,"G")=0 THEN
		    sQuery="INSERT INTO ACC_T_CREATEDVOUCHERDETAILS (CREATEDTRANSNO,ACCOUNTINGUNIT," &_
			    "VOUCHERENTRYNUMBER,ACCUNITACCOUNTHEAD,ACCUNITPARTYTYPE,ACCUNITPARTYCODE,ACCUNITPARTYSUBTYPE," &_
			    "VOUCHERNARRATION, AMOUNT,TRANSCRDRINDICATION) VALUES (" &_
			    "" & iTransNo & ",'" & sAccUnit & "'," &_
			    "" & sEntryno & "," & sAccCode & ",NULL,NULL,NULL," &_
			    "'" & sNarration & "'," & sAmount & ",'" & sEntryType & "')"
	    ELSE
		    sTemp=Split(sAccCode,"?")
		    sQuery="INSERT INTO ACC_T_CREATEDVOUCHERDETAILS (CREATEDTRANSNO,ACCOUNTINGUNIT," &_
			    "VOUCHERENTRYNUMBER,ACCUNITACCOUNTHEAD,ACCUNITPARTYTYPE,ACCUNITPARTYSUBTYPE,ACCUNITPARTYCODE," &_
			    "VOUCHERNARRATION, AMOUNT,TRANSCRDRINDICATION) VALUES (" &_
			    "" & iTransNo & ",'" & sAccUnit & "'," &_
			    "" & sEntryno & ",NULL,'" &sTemp(0) & "','," & sTemp(1) & "," & sTemp(3) & "," &_
			    "'" & sNarration & "'," & sAmount & ",'" & sEntryType & "')"
	    END IF
        Response.Write "<br>"&sQuery& "<BR>"	
        con.execute(sQuery)
    end if 'if CDbl(sAmount)>0 then
'-----------------------END OF DETAIL TABLE UPDATION----------------------
DIM sCCGroup,sAddCode,sAddRatio,sAddAmount,dAddTotal
dAddTotal=0	
'--------PROCESS CHILD NODES OF ENTRIES FOR ADDTIONAL DETAILS UPDATION----
	FOR EACH HeaderNode IN EntryNode.childNodes
'----------------------PROCESS COST CENTER NODES -------------------------
		IF 	HeaderNode.nodeName="CostCenter" THEN
			FOR EACH  nodANL IN HeaderNode.childNodes
				sAddCode=nodANL.Attributes.Item(0).nodeValue
				sAddRatio=nodANL.Attributes.Item(3).nodeValue 
				sAddAmount=nodANL.Attributes.Item(4).nodeValue
				sQuery="INSERT INTO ACC_T_CREATEDVOUCHERCCDET(CREATEDTRANSNO, VOUCHERENTRYNUMBER, ACCOUNTINGUNIT," &_
					"ACCUNITACCOUNTHEAD,ACCUNITCCHEAD," &_
					"CCRATIOPERCENT, CCRATIOAMOUNT)" &_
					" VALUES(" & iTransNo & "," & sEntryno & ",'" & sAccUnit & "'," & sAccCode & "," &_
					" " & sAddCode & "," & sAddRatio & "," & sAddAmount & ")"
					Response.Write sQuery& "<BR>"	
					con.execute(sQuery)
			NEXT
		END IF 
'-------------END OF PROCESSING COST CENTER NODES ------------------------
'----------------------PROCESS ANALYTICAL NODES --------------------------
		IF 	HeaderNode.nodeName="Analytical" THEN
			FOR EACH  nodANL IN HeaderNode.childNodes
				sAddCode=nodANL.Attributes.Item(0).nodeValue
				sAddRatio=nodANL.Attributes.Item(3).nodeValue 
				sAddAmount=nodANL.Attributes.Item(4).nodeValue

				sQuery="INSERT INTO ACC_T_CRETEDVOUCHERAHDET(CREATEDTRANSNO,VOUCHERENTRYNUMBER,ACCOUNTINGUNIT,"&_
					"ACCUNITACCOUNTHEAD, ACCUNITANALYTICALCODE,"&_
					"RATIOPERCENTAGE, RATIOAMOUNT)"&_
					" VALUES("& iTransNo & "," & sEntryno & ",'" & sAccUnit & "'," & sAccCode & ","&_
					"" & sAddCode & "," & sAddRatio & "," & sAddAmount & ")"
					Response.Write sQuery& "<BR>"	
					con.execute(sQuery)
			NEXT
		END IF 
'-------------END OF PROCESSING ANALYTICAL NODES -------------------------
	NEXT 
NEXT
'------------------END OF PROCESSING ENTRY NODES--------------------------
					
if con.Errors.count <>0 then
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) &"<br>"
	next
	'Redirect to Error Handling System
else
'	con.rollbacktrans
'	Response.End 
	Response.Clear 
	con.CommitTrans
	Root.setAttribute "TransNo",iTransNo
	'oDOM.Save server.MapPath("../../Accounts/xmldata/Voucher/"&iTransNo&".xml")	
%>
<HTML>
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript>
	alert("Closing Stock Details to Accounts has been updated Successfully");
	window.location.href = "sendToAccounts.asp";
</SCRIPT>	
</html>
<%
end if
%>

