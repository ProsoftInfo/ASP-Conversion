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
	'Program Name				:	PRNGJView.asp
	'Module Name				:	Accounts (Reports)
	'Author Name				:	Tajudeen.S
	'Created On					:	29 March 2004
	'Modified By				:	S.Maheswari
	'Modified On				:	5th Sep 2008
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
	'CounteRs					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/PrintFunctions.asp"-->
<%
'------------------------Declaration Constants -----------------------------
Dim aiHeadColWidth(6,8)

'WIDTH SPECIFICATION FOR PAGE TITLE 1
aiHeadColWidth(0,0)=80
aiHeadColWidth(0,1)=12
aiHeadColWidth(0,2)=14
aiHeadColWidth(0,3)=13
aiHeadColWidth(0,4)=9
aiHeadColWidth(0,5)=10

'WIDTH SPECIFICATION FOR PAGE TITLE 2
aiHeadColWidth(1,0)=30
aiHeadColWidth(1,1)=25
aiHeadColWidth(1,2)=20
aiHeadColWidth(1,3)=10
aiHeadColWidth(1,4)=15

'WIDTH SPECIFICATION FOR OPENING/CLOSING  LINE 
aiHeadColWidth(2,0)=5
aiHeadColWidth(2,1)=80
aiHeadColWidth(2,2)=5

'WIDTH SPECIFICATION FOR HEADING DETAIL LINE 
aiHeadColWidth(3,0)=15 '3
aiHeadColWidth(3,1)=15
aiHeadColWidth(3,2)=20
aiHeadColWidth(3,3)=15
aiHeadColWidth(3,4)=15
aiHeadColWidth(3,5)=10
aiHeadColWidth(3,6)=3

'WIDTH SPECIFICATION FOR HEADING DETAIL LINE 
aiHeadColWidth(4,0)=0 '3
aiHeadColWidth(4,1)=35
aiHeadColWidth(4,2)=14
aiHeadColWidth(4,3)=14
aiHeadColWidth(4,4)=14
aiHeadColWidth(4,5)=7
aiHeadColWidth(4,6)=15
aiHeadColWidth(4,7)=3

dim objFSO,objTxt
dim iPageNo,iLineNo,sTStr,i,sFinalText,sVoucherDate,iVoucherNo,AccVoucherNo
set objFSO = Server.CreateObject("Scripting.FileSystemObject")
set objTxt = objFSO.CreateTextFile(server.MapPath("../temp/transaction/"&Session.SessionID&"_GJ_View.txt"))

dim iTransNo,strText,XmlData,Root,Node,subNode,Narration,ObjRs,ObjRs1,TempDate,sOrgId
dim Au,Accode,AcName,Drcr,Amount,TotDr,TotCr,sQuery,sType,sCode,z,iEnNo,sRetVal
dim sEmpName,sTotFlag,sFlag,sTempDrCr,j
set ObjRs=server.CreateObject("ADODB.Recordset")
set ObjRs1=server.CreateObject("ADODB.Recordset")
set XmlData=server.CreateObject("Microsoft.XMLDOM")

iTransNo=Request.QueryString("iTransNo")
'XmlData.load server.MapPath("../XmlData/Voucher/" & iTransNo & ".xml")
sRetVal = GetVouchXML(iTransNo)
XmlData.Load server.MapPath(sRetVal)
iLineNo=1
set Root=XmlData.documentElement
sOrgId=Root.Attributes.Item(0).nodeValue
sVoucherDate=Root.Attributes.Item(5).nodeValue
sEmpName =  Session("UserName")
totcr=0
totdr=0
ConsPgHeader1(1)

objTxt.write strText 
strText=""
'iLineNo=14
iPageno = 1
iEnNo=0
iLineNo = iLineNo + 2

for each node in root.childNodes	
	iEnNo=iEnNo+1

	au=node.attributes.getNamedItem("AccUnit").value
	drcr=node.attributes.getNamedItem("CRDR").value
	Amount=node.attributes.getNamedItem("Amount").value 
 
	IF trim(drcr) = "D" then  'Added newly
		TotDr=totdr+amount
		for each SubNode in node.ChildNodes		
			if subnode.nodeName="AccHead" then
				accode=	subNode.attributes.getNamedItem("No").value
				sType = subNode.attributes.getNamedItem("Type").value
				IF CStr(sType) = "G" Then
					sQuery = "Select AccountHeadCode from Acc_M_GLAccountHead where AccountHead=" & Accode
					'Response.Write sQuery
					with ObjRs
						.CuRsorLocation=3
						.CuRsorType=3
						.Source = sQuery
						.ActiveConnection=con
						.Open 
					end with
					set ObjRs.ActiveConnection =nothing
					if not ObjRs.EOF then Accode=ObjRs(0)
					ObjRs.Close 
				Else
					Accode = " "
				End IF	
	
	
				strText=strText & string(14," ") & myalign(subNode.attributes.getNamedItem("Name").value,aiHeadColWidth(4,1)+1,"L") & myalign(" ",aiHeadColWidth(4,5)+1,"L")  
				strText = strText & string(4," ")& myalign(FormatNumber(Amount),aiHeadColWidth(4,3),"R") & string(aiHeadColWidth(4,4)+2," ") & vbCrLf
			
			elseif subNode.nodeName="Narration" then
				if IsNull(subnode.Text) then
					Narration=""
				else
					Narration = subnode.Text
				end if			
			end if
			sTempDrCr = drcr
			
			objTxt.write strText
			strText=""	
		next
	End IF
Next	

for each node in root.childNodes	
	iEnNo=iEnNo+1

	au=node.attributes.getNamedItem("AccUnit").value
	drcr=node.attributes.getNamedItem("CRDR").value
	Amount=node.attributes.getNamedItem("Amount").value  
 
	IF trim(drcr) = "C" then  'Added newly	
		TotCr=TotCr+amount
		for each SubNode in node.ChildNodes		
			if subnode.nodeName="AccHead" then
				accode=	subNode.attributes.getNamedItem("No").value
				sType = subNode.attributes.getNamedItem("Type").value
				IF CStr(sType) = "G" Then
					sQuery = "Select AccountHeadCode from Acc_M_GLAccountHead where AccountHead=" & Accode
					'Response.Write sQuery
					with ObjRs
						.CuRsorLocation=3
						.CuRsorType=3
						.Source = sQuery
						.ActiveConnection=con
						.Open 
					end with
					set ObjRs.ActiveConnection =nothing
					if not ObjRs.EOF then Accode=ObjRs(0)
					ObjRs.Close 
				Else
					Accode = " "
				End IF	
					
				IF sFlag <> True then
					strText=strText & vbCrLf
					iLineNo = iLineNo + 1
					strText=strText & string(14," ") & myalign("To",4,"L") & myalign(subNode.attributes.getNamedItem("Name").value,aiHeadColWidth(4,1)+9,"L") '& myalign(" ",aiHeadColWidth(4,5)+1,"L")
				  	sFlag = True
				Else			
					strText=strText & string(14," ") & myalign("",4,"L")  & myalign(subNode.attributes.getNamedItem("Name").value,aiHeadColWidth(4,1)+1,"L") & myalign(" ",aiHeadColWidth(4,5)+1,"L")  
				End If
			
				strText = strText & string(4," ")& myalign(" ",aiHeadColWidth(4,3)-2,"R") & myalign(FormatNumber(Amount),aiHeadColWidth(4,4),"R") & vbCrLf
							
			elseif subNode.nodeName="Narration" then
				if IsNull(subnode.Text) then
					Narration=""
				else
					Narration = subnode.Text
				end if			
			end if
			sTempDrCr = drcr
			
			objTxt.write strText
			strText=""			
		
		next	
	End If 
	iLineNo=iLineNo+1  
	if iLineNo > 25 then								
		for j = iLineNo to 29  
			'Response.Write iLineNo		
			strText = strText & string(80," ") &  vbCrLf 
			iLineNo = iLineNo + 1
		Next
		
		footer()
		
		'iLineNo = 2
		
		ConsPgHeader1 (iPageno)	
		'iLineNo=iLineNo + 2			
	end if
 next 

strText = strText & string(45," ") & myalign("",10,"R") & string(6," ") &   myalign("---------------",15,"R") & string(1," ")& myalign("---------------",15,"R")&vbCrLf
strText = strText & string(45," ") & myalign("TOTAL",10,"R") & string(6," ") &   myalign(FormatNumber(TotDr),15,"R") & string(1," ")& myalign(FormatNumber(TotCr),15,"R")&vbCrLf
strText = strText & string(45," ") & myalign("",10,"R") & string(6," ") &     myalign("---------------",15,"R") & string(1," ")& myalign("---------------",15,"R")&vbCrLf
iLineNo = iLineNo + 3


Dim sNarr,sNarr1,sNarr2,sTemp,Cnt 
sNarr = trim(Narration)
sTemp = Len(Narration)
Cnt = cint(sTemp)/25
'strText = strText  & string(14," ")&  myalign("NARRATION : ",aiHeadColWidth(0,1),"L") 
sNarr  = "NARRATION : "&sNarr
iLineNo = iLineNo + 1
sNarr = BreakString(sNarr,45)

For i = 0 to UBOUND(sNarr)
	IF trim(sNarr(i)) <> "" then
		strText = strText  & string(14," ")&  myalign(sNarr(i),aiHeadColWidth(0,0),"L") & vbCrLf		
		iLineNo = iLineNo + 1
		'Response.Write sNarr(i)&":::"&iLineNo&"<BR>"	 
	End IF
Next 

strText = strText & string(80," ") & vbCrLf	
ilineNo = ilineNo + 1
For i =  iLineNo to 29
	strText = strText & string(80," ") &  vbCrLf	 
    ilineNo = ilineNo + 1
Next
 
footer()
sFinalText = sFinalText & strText
'sFinalText = sFinalText & FormattPrint("PAGESKIP","")
objTxt.write sFinalText
if strText<>"" then
	Response.Redirect("../../Components/FormattPrint.asp?server=server&filepath=/accounts/temp/transaction/"&Session.SessionID&"_GJ_View.txt&exitpath=/accounts/reports/CreditNoteSelection.asp&frame=_parent")
else
	Response.Clear
end if

%>              
<%

Function footer()
	strText = strText & string(10," ")  & myAlign(sEmpName&"-"&sVoucherDate,aiHeadColWidth(1,2),"L")  & vbCrLf  
	strText = strText & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf  & vbCrLf    	
End function

Function centerAlign(str1,width)
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

Function ConsPgHeader1(iPageNo)
	 
iPageno = iPageNo + 1
iEnNo=0
iLineNo = 1
sQuery = "SELECT ORGUNITDESCRIPTION,ADDRESS1,CITY,STATE,POSTCODE FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID='"&sOrgId&"'"

With ObjRs
	.ActiveConnection = con
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.Open
End With
Set ObjRs.ActiveConnection = nothing

IF Not ObjRs.EOF Then	
	strText = strText &  string(10," ") & myalign(Trim(Objrs(0)),aiHeadColWidth(0,0),"L") & vbCrLf
	'objTxt.write myalign(Trim(Objrs(1))&","&Trim(Objrs(2)),aiHeadColWidth(0,0)-32,"L")&" - "& Trim(Objrs(3))& vbCrLf
	strText = strText &  string(10," ") & myalign(Trim(Objrs(1))&","&Trim(Objrs(2)),aiHeadColWidth(0,0)-32,"L")& vbCrLf		
	strText = strText &  string(10," ") & myalign(Trim(Objrs(3)),aiHeadColWidth(0,0)-45,"L")
	'strText = strText & " " & vbCrLf
	'objTxt.write  strText & " " & vbCrLf
	iLineNo= iLineNo + 3
	
End IF
ObjRs.close

sQuery="Select H.CreatedVouchStatus,V.VoucherNumber,V.CreatedVoucherNo from Acc_T_CreatedVoucherHeader H , Acc_T_VoucherHeader v where H.CreatedTransNo=v.CreatedTransNo and " _
& "right(H.CreatedVouchStatus,2)=04  and H.CreatedTransNo="&iTransNo
With ObjRs
	.ActiveConnection = con
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.Open
End With
Set ObjRs.ActiveConnection = nothing
if not ObjRs.EOF then 
	AccVoucherNo=ObjRs(1)
	iVoucherNo=ObjRs(2)
else
	iVoucherNo=Root.attributes.Item(9).nodevalue	
end if
ObjRs.Close 

IF trim(AccVoucherNo) <> "" then 
	strText = strText &  myalign("",aiHeadColWidth(3,0)+6,"L") & myalign(AccVoucherNo,aiHeadColWidth(3,1),"L")& string(2," ")& myalign(sVoucherDate,aiHeadColWidth(3,1),"R") & vbCrLf 
Else
	strText = strText &  myalign("",aiHeadColWidth(3,0)+6,"L") & myalign(iVoucherNo,aiHeadColWidth(3,1),"L")& string(2," ")& myalign(sVoucherDate,aiHeadColWidth(3,1),"R") & vbCrLf 
End IF
'objTxt.write myalign(" ",aiHeadColWidth(3,0)+aiHeadColWidth(3,1),"L")& string(24," ")& myalign("Voucher No : ",aiHeadColWidth(3,0),"R") & myalign(AccVoucherNo,aiHeadColWidth(3,1),"L") & vbCrLf 

'iLineNo=iLineNo + 1
	'if not iPageNo =1 then		
		'objTxt.write string(aiHeadColWidth(0,0),"-") & vbCrLf		
		'objTxt.write myalign("A/c Head",aiHeadColWidth(0,1),"L") & string(41," ") & myalign("Debit",aiHeadColWidth(0,2),"L")& myalign("Credit",aiHeadColWidth(0,2),"L") & vbCrLf
		'objTxt.write string(aiHeadColWidth(0,0),"-") & vbCrLf
		strText = strText & " " &  vbCrLf & vbCrLf & vbCrLf
		'objTxt.write   strText
													
	'end if
	
end function

%>
