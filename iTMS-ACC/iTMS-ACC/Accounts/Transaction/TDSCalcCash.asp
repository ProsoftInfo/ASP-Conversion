<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	TDSCalcCash.ASP
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	MAHESHWARI S
	'Created On					:	Feb 26,2007
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			:  
	'Input Parameter			:	
	'							:
	'Connects To				:	BankVoucher.asp
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->

<%				
Dim objFS,sQry,rs,sTDSAmt,sGrpId,sGrpName,sCompMode,sAcHeadCode,sCompForm,sEntNo,iGrpId,sNewAmt 
Dim sTemp,Arr,Arr1,sVal,Value,iCtr,n,i,j,sTemp1,sArrVal,sTempVal,sVal1,sVal2,sTotValue 
Dim iVal1,iVal2,iVal3,iVal4,newElem,RootElem,subnode,sNewPer,sNewCtr,sArrGrp,ArrGrpId,iGrpHeadId
Dim TDSnode,TFormula,Tamt,iCnt,sArrTemp,sArrTemp1,sArrTempVal,k,sPer,sNewArr1,sNewArr2
Dim Root,docobj

Set objFs = server.CreateObject("Scripting.FileSystemObject")
Set rs = server.CreateObject("ADODB.Recordset")
set docobj=server.CreateObject("Microsoft.XMLDOM")
set Root=docobj.DocumentElement

set Root=docobj.CreateElement("Root")
docobj.appendchild Root

sEntNo = Request("EntNo") 						
sTDSAmt = Request("Amount")
iGrpId = Request("GrpId")
sNewAmt = Request("NewAmt")
'Response.Write "iGrpId="& iGrpId &"<BR><BR>"
'Response.Write "sNewAmt="& sNewAmt
													
If trim(sTDSAmt) <> "" then
sQry = " Select GroupHeadID,GroupHeadName,ComputeMode,AcHeadCode,isNull(ComputeFormula,'') "&_
	   " from ACC_M_TDSHeadComputation where GroupId = "&iGrpId&""
	  ' Response.Write sQry 
with rs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = con
	.Source = sQry 
	.Open
end with
iCtr = 1	
																		
Do while not rs.EOF
	sGrpId = rs(0)
	sGrpName = rs(1)
	sCompMode = rs(2)
	sAcHeadCode = rs(3)
	sCompForm = rs(4)
	iGrpHeadId = iGrpId &"-"& sGrpId
	'Response.Write "sCompForm=" &sCompForm						 
	If trim(sCompForm) = "" then
			sTemp1 = "0"
			Value = "0"
	Else
		IF trim(sCompMode) <> "F" then
			Arr = split(sCompForm,",")
			sTemp =Arr(0)
			'Response.Write "***"&sTemp&"***"
			IF sTemp <> "0" then
				Arr1 = split(sTemp,"-")
			'	Response.Write "###"&Arr1(0)
				sTemp1 = Arr1(1)
			End IF
		End IF ' IF trim(sCompMode) <> "F" then
	End If	
	Set newElem = docobj.CreateElement("TDS")														
	newElem.SetAttribute "Ctr",iCtr
	newElem.SetAttribute "AccHeadCode",sAcHeadCode
	newElem.SetAttribute "TDSAmount",Round(sTDSAmt,2)
	newElem.SetAttribute "TdsPercentage",sTemp1
	newElem.SetAttribute "PayRecAmount","0"  
	newElem.SetAttribute "Formula",sCompForm 
	newElem.SetAttribute "GroupHeadId",iGrpHeadId  
	newElem.SetAttribute "TdsRndOff","N"
	
	Root.appendchild newElem													
														
	
	If trim(sCompMode)= "F" then
		sCompMode = ""
		sCompForm = ""	
		Value = "0"
		sTemp1 = "0"
	End If
	If trim(sCompForm) = "" then
		sTemp1 = "0"
		Value = "0"
	End If
	If trim(sCompMode) = "P" then
		sCompMode = "P"
		
		Value=CalculatePer(Root,sCompForm,sTDSAmt)
																	
			For each TDSnode in Root.childnodes
				If trim(iCtr) = TDSnode.getAttribute("Ctr") then
					TDSnode.setAttribute "PayRecAmount",Value																		
				End If
			next
														
	End If 'If trim(sCompMode) = "P" then
	'Response.Write Value
																			
	sTotValue = sTotValue + Value 
	'Response.Write sTotValue
														 
	docobj.save server.MapPath("../temp/transaction/TDS_Bank_"&Session.SessionID&".xml")       
	iCtr = iCtr + 1
	rs.MoveNext 
		loop
	rs.Close
	Response.ContentType="text/xml"
	Response.Write docobj.xml

	End If 'If trim(sTDSAmt) <> "" then
%>
	
<%
Function CalculatePer(Root,sFormula,sTDSAmt)
	dim saTemp,iCounter,iTemp,dPercentage,sTotAmt
	dim oNodTemp,sGrpCtr,sVal,sTempGrp
	dim saTemp1,iCtr

	saTemp=Split(sFormula,",")
	'Response.Write sFormula
	'sTotAmt = 0
	iTemp = 0
	iCtr = 1
	sTempGrp = split(saTemp(0),"#")
	sGrpId = sTempGrp(0)
	dPercentage = Split(trim(saTemp(0)),"-")
	
	If trim(sGrpId) = "0" then
		sTotAmt=sTDSAmt*(cdbl(dPercentage(1))/100)
	End If
	
	For iCounter=iTemp to UBound(saTemp)

		dPercentage = Split(trim(saTemp(iCounter)),"-")
		saTemp1=Split(trim(saTemp(iCounter)),"#")
		sGrpId = saTemp1(0)
		sGrpCtr = left(saTemp1(1),1)

		For Each oNodTemp in Root.childNodes
			If trim(sGrpCtr) = oNodTemp.getAttribute("Ctr") then
				sVal = oNodTemp.getAttribute("PayRecAmount")
				sTotAmt = sTotAmt + sVal *(dPercentage(1) / 100)
			End If
		Next
	next
	CalculatePer = sTotAmt

End function
%>