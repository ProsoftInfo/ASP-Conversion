<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	NarrationUpdate.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	SENTHIL E
	'Created On					:	December 17, 2002
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
<%
Dim sQuery,objRs,newxml,Root,Node,sOrgID,iBookCode,iBookNo
dim sName,sShortName,Temp,arrBook,iCounter,iCode,iRecCount
dim sFlag,bUpdateFlag,sMessage,iNarrID,sType
Dim iBookCodeArr,iBookNoArr,iCtr
Set objRs = Server.CreateObject("ADODB.RecordSet")

sOrgID = trim(Session("organizationcode"))
	
set newxml = Server.CreateObject("Microsoft.XMLDOM")
newxml.async = False
newxml.Load(Request)

Con.Begintrans

set Root = newxml.documentElement

sFlag =False

For Each Node in Root.childNodes
	
	If Node.nodeName="Desc" then
	
		sType = trim(Node.getAttribute("Type"))
		iNarrID = Node.getAttribute("NarrNo")
		sName = Node.getAttribute("Desc")
		sShortName = Node.getAttribute("ShortDesc")
		iBookCode  = Node.getAttribute("BookCode")
		iBookNo    = Node.getAttribute("BookNo")
		iBookCodeArr = Split(iBookCode,",")
		iBookNoArr = Split(iBookNo,",")
		
		If sType = "N" Then	
			
			sQuery="select count(1)  from  Acc_M_FrequentDescriptions where NarrationShortDesc='"&sShortName&"'"
			
			with objRs	
				.CursorLocation = 3
				.CursorType = 3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			iRecCount=objRs(0)
			objRs.Close
			
			if iRecCount =0 then 
				sFlag = True
					
				sQuery="select isnull(max(NarrationNumber)+1,1)  from Acc_M_FrequentDescriptions"
					
				with objRs	
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with
				iCode=objRs(0)

				sQuery="INSERT Acc_M_FrequentDescriptions(NarrationNumber, NarrationDesc,NarrationShortDesc) "&_
						"VALUES("&iCode&",'"&sName&"','"&sShortName&"')"
				'Response.Write "<p>"&sQuery 
				con.Execute(sQuery)
			
				For iCtr= LBound(iBookCodeArr) to UBound(iBookCodeArr)
					
					sQuery="INSERT INTO Acc_R_BookFreqDesc(OUDefinitionID, BookCode, BookNumber, NarrationNumber)"&_
							"VALUES('"&sOrgID&"','"&iBookCodeArr(iCtr)&"',"&iBookNoArr(iCtr) &","&iCode&")"
				'	Response.Write "<p>"&sQuery 
					con.Execute(sQuery)
				Next
			End IF
		Else
			
			sQuery="update Acc_M_FrequentDescriptions set NarrationDesc='"&sName&"',NarrationShortDesc='"&sShortName&"' "&_
					" where NarrationNumber="&iNarrID	
			con.Execute(sQuery)
		
			sQuery="delete from  Acc_R_BookFreqDesc where NarrationNumber="&iNarrID
			con.Execute(sQuery)
		
			For iCtr= LBound(iBookCodeArr) to UBound(iBookCodeArr)
					
				sQuery="INSERT INTO Acc_R_BookFreqDesc(OUDefinitionID, BookCode, BookNumber, NarrationNumber)"&_
						"VALUES('"&sOrgID&"','"&iBookCodeArr(iCtr)&"',"&iBookNoArr(iCtr) &","&iNarrID &")"
				con.Execute(sQuery)
					
			next	
				
			
		End IF	'If sType = "N" Then	
		
	End If	'If Node.nodeName="Desc" then
	
Next


'con.RollbackTrans
'Response.End 
Con.CommitTrans
con.close
Set con = Nothing

%>
