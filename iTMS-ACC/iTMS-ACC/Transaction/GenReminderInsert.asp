<%@ Language=VBScript %>
<%	option explicit	%>
<%
		
	'Program Name				:	GenReminderInsert.asp
	'Module Name				:	Accounts (Transaction)
	'Author Name				:	UMAMAHESWARI S	
	'Created On					:	April 11, 2011
	'Modified On				:   
	'Modified by				:   
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/MatPopulate.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->
<!-- #include File="../../include/NoSeries.asp" -->

<%
    Dim objDom,Root,Node,SubNode
	Dim sSql,nReminderNo,iPartyCode,nTransNo,nPartyInvNo,sPartyInvDate,nInvAmt,nAmtPaid
	Dim nAmtOutStanding,nOverDues,sSendBy,sCorComName,sCorTransID,sCorAddress
	Dim nRemTypeId,nRemReason,sRemContent,sPartyType,sSentOrReceipt,sSelect
	
	Dim objRs1,objRs2
	
	Set objRs1 = Server.CreateObject("ADODB.RecordSet")
	Set objRs2 = Server.CreateObject("ADODB.RecordSet")
	
	set objDom = Server.CreateObject("Microsoft.XMLDOM")
	objDom.async = False
	objDom.load(Request)
		
	set Root = objDom.documentElement 
	
	'objdom.save server.MapPath("../Temp/Transaction/ReminderXML.xml")
	con.beginTrans
	
	sSql = "Select ReminderTypeID,ApplicationCode,ReminderLetterContent,ReminderTypeDesc From APP_M_ReminderTypes"
	objRs1.Open sSql,con
	If Not objRs1.EOF Then
		nRemTypeId = objrs1(0)
		sRemContent = objrs1(2)
		nRemReason = ""
	End IF
	objRs1.Close 
	
	If Root.hasChildNodes Then
			
		sSendBy = Root.getAttribute("SENDBY")
		sCorComName = Root.getAttribute("NAME")
		sCorTransID = Root.getAttribute("ID")
		sCorAddress = Root.getAttribute("ADDRESS")

		For Each Node in Root.childNodes
			
			If Node.nodeName = "Party" Then
			
				iPartyCode = Node.getAttribute("CODE")
				sPartyType = Node.getAttribute("TYPE")
				if trim(sPartyType)="CR" then
				    sSentOrReceipt = "R"
				else
				    sSentOrReceipt = "S"
				end if
				
				For Each SubNode in Node.childNodes
				    sSelect = SubNode.getAttribute("SELECT")
				    if sSelect = "Y" then
				
					    sSql = "Select ISNULL(MAX(isNULL(ReminderNo,0)),0)+1 From ACC_T_OverDueReminderDet"   
					    objRs1.Open sSql,con
					    If not objRs1.EOF Then
						    nReminderNo = objRs1(0)
					    End IF
					    objRs1.Close 
    					
    					
					    nTransNo = SubNode.getAttribute("TRANSACTIONNO")
					    nPartyInvNo = SubNode.getAttribute("INVOICENO")
					    sPartyInvDate = SubNode.getAttribute("DATE")
					    nInvAmt = SubNode.getAttribute("AMOUNT")
					    nAmtPaid = SubNode.getAttribute("AMOUNTPAIDTILLDATE")
					    nAmtOutStanding =  SubNode.getAttribute("BALANCE")
					    nOverDues = SubNode.getAttribute("NOOFDAYSOVER")
    					
    					if Trim(nOverDues)="" or IsNull(nOverDues) then nOverDues = "0"
    					
					    sSql = "Select ReminderTypeDesc from APP_M_ReminderTypes where ReminderTypeID = "& nRemTypeId 
					    objrs2.Open sSql,con
					    if not objRs2.EOF then
					        nRemReason = objRs2(0)
					    end if
					    objrs2.Close 
    					
					    If sSendBy = "" Then sSendBy = "E"	
    					
					    sSql = "select ReminderNo From ACC_T_OverDueReminderDet Where PartyCode="&iPartyCode &" and TransactionNo ="& nTransNo&" and PartyInvoiceNo = '"& nPartyInvNo &"'"
					    'Response.Write sSql
    					
					    objRs2.Open ssql,con
    					
					    If objRs2.EOF Then
    					
						    sSql =  " INSERT INTO APP_R_ApplicationReminders (ReminderNo,ReminderDate,ReminderTypeID,ApplicationCode, "&_
								    " ReminderToPartyCode,ReminderReason,CreatedBy,CreatedOn,SentThrough,CourierCompanyName,CourierTransactionID,CourierAddress,"&_
								    " ReminderLetterContent,ActionTaken,ActionTakenOn,SentOrReceipt,PartyType) VALUES ("& nReminderNo &",Convert(DateTime,getDate(),103),"&nRemTypeId&",1,"&_
								    " "& iPartyCode &",'"&nRemReason &"',"& Session("userid")&",Convert(DateTime,getDate(),103),'"& sSendBy &"','"& sCorComName &"',"&_
								    " '"& sCorTransID &"','"& sCorAddress &"','"& sRemContent &"','Created',Convert(datetime,getDate(),103),'"& sSentOrReceipt &"','"&sPartyType&"') "
						    'Response.Write "<p>"&sSql
						    con.Execute sSql

						    sSql  = " INSERT INTO ACC_T_OverDueReminderDet(ReminderNo,PartyCode,TransactionNo,PartyInvoiceNo,PartyInvoiceDate,"&_
								    " InvoiceAmount,AmountPaid,AmountOutStanding,OverDueDays,PartyType) VALUES ("& nReminderNo&","& iPartyCode&","&_
								    " "& nTransNo & ",'"& nPartyInvNo &"',Convert(Datetime,'"& sPartyInvDate &"',103),"& nInvAmt &","&_
								    " "& nAmtPaid &","& nAmtOutStanding &","& nOverDues &",'"& sPartyType &"')"
						    'Response.Write "<p>"&sSql
						    con.Execute sSql
						else
						    Response.Write "Reminder Already Exist For the Invoice"
					    End IF
					    objRs2.Close 
					end if 'if sSelect = "Y" then
				Next
				
			End IF	'If Node.nodeName = "Party" Then
			
		Next 
	End IF
	
'	con.rollbackTrans
'	Response.End 
	' Response.Clear 
	 con.commitTrans
	Response.Write "@RemNo="& nReminderNo 
%>