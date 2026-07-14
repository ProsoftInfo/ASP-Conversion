<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	PurNoSeriesInsert.asp
	'Module Name				:	Purchase (Master Creation)
	'Author Name				:	Malathi N
	'Created On					:	Sep 29,2004
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/NoSeries.asp"-->
<!--#include virtual="/include/Accpopulate.asp"-->

<%
dim iUnitNo,iActivity,iCounter,iInvSeriesCode,sType
dim sSql,iExistBookNo,sItmType,sActName
dim iSeries,iSeriesType,bPayRecNo,iLength,sAgentcode,objrs,objrs1
Dim sFromFin,sToFin,iOldSrCode,iOldSrNo,iTotalEntNo,iCtr,sQuery
Dim sNumFor,sItemValue,sSalType,sSalValue,sInvType,sInvValue,sAgTy,iTransNo,iEntryNo
Dim arrTemp,objDOM,objfs,Root,sExp,TempNode,iCount
Dim sSuffix,sPrefix,iNumber,sTotEntNo,sClassCode,sArrClassCode,iCnt,sCatCode,sArrCateCode

Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")
Set objrs = Server.CreateObject("ADODB.RecordSet")
Set objrs1 = Server.CreateObject("ADODB.RecordSet")

con.BeginTrans

If objfs.FileExists(Server.MapPath("./temp/master/NoSeries_PUR_"&Session.SessionID&".xml")) then
	objDOM.Load server.MapPath("./Temp/master/NoSeries_PUR_"&Session.SessionID&".xml")
	Set Root = objDOM.documentElement
	
	'Response.Write Request.Form +"<br>"

	iUnitNo=trim(Request.Form("selUnit"))
	iActivity=trim(Request.Form("selActType"))
	iSeries=trim(Request.Form("selNoSeries"))
	iSeriesType=trim(Request.Form("hSeriesType"))
	iLength=trim(Request.Form("hSeriesLen"))
	sActName = trim(Request.Form("hActivityName"))
	sTotEntNo = Request.Form("hTotEntNo")
	sClassCode = (Request.Form("hClassCode"))
	sCatCode = Trim(Request.Form("hCatCode"))
	
	Response.Write "<p> Unit Code = "& iUnitNo
	
	'Response.Write iSeries & "  " & iSeriesType +"<br>"
	
	if Trim(sClassCode)="" or IsNull(sClassCode) then sClassCode = "NULL"
	if Trim(sCatCode)="" or IsNull(sCatCode) then sCatCode = "NULL"
	
	sArrClassCode = Split(sClassCode,",")
	sArrCateCode = Split(sCatCode,",")
	
	
	sExp = "//NumSeriesList[@EditCheck!=""N""]"
	'Response.Write sExp
	Set TempNode = Root.selectNodes(sExp)
	
	For iCount = 0 To TempNode.length - 1
		'To check whether the entry is new entry "E" For New entry
		IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("EditCheck").value) = "E" Then
		
			IF CStr(Left(TempNode.Item(iCount).Attributes.getNamedItem("NumFor").Value,1)) = "B" Then
				sNumFor = "B"
			ElseIF CStr(Left(TempNode.Item(iCount).Attributes.getNamedItem("NumFor").Value,1)) = "I" Then
				sNumFor = "I"
			Else
				sNumFor = "D"
			End IF
			
			IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("ItemTy").Value) = "Specific" Then
				sItmType = "1"
			Else
				sItmType = "0"
			End IF
			
			
			sItemValue = TempNode.Item(iCount).Attributes.getNamedItem("ItemValue").Value
			sFromFin = Trim(Request.Form("hFinFrom"))
			sToFin = Trim(Request.Form("hFinTo"))

			iOldSrNo = Trim(Request.Form("hSeriesNo"))
			iOldSrCode = Trim(Request.Form("hSeriesCode"))
			'Response.Write iOldSrCode&iOldSrNo
			iTotalEntNo = Trim(Request.Form("hEntryNo"))

			
			if CStr(sItmType) = "1" or CStr(sItmType) = "0" Then
				'Response.Write "GenSeriesCode"
				iInvSeriesCode = GenSeriesCode(iUnitNo,"2","2",iSeries,iSeriesType,"",sActName,iLength)
				sQuery = "Select isNull(Max(NoSeriesTransactionNo),0) + 1 From Pur_M_Noseries "
				With objrs
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = Con
					.Open
				End With
				Set objrs.ActiveConnection = Nothing
				If not objrs.EOF Then
					iTransNo = objrs(0)
				End IF
				objrs.Close
				iEntryNo = 1
			'	sQuery = "INSERT INTO Pur_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
			'			 "NumberFor,SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
			'			 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"',"&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode &") "
						 
                sQuery = "INSERT INTO Pur_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
						 "NumberFor,SeriesNo, SeriesCode, NoSeriesStatus) "&_
						 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"',"&iSeries&", "&iInvSeriesCode&", '0') "
				Response.Write sQuery &"<br><br>"
				con.Execute sQuery
				
				if UBound(sArrCateCode)=UBound(sArrClassCode) then
				    For iCnt = 0 to UBound(sArrClassCode)
				        sQuery = "Insert into PUR_M_NoSeriesClass (SeriesNo,SeriesCode,ClassCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","& sArrClassCode(iCnt) & ","& sArrCateCode(iCnt) &")"
				        Response.Write "<p>"& sQuery
				        con.execute sQuery
				    Next
				elseif UBound(sArrCateCode)>0 then
				    For iCnt = 0 to UBound(sArrCateCode)
				        sQuery = "Insert into PUR_M_NoSeriesClass (SeriesNo,SeriesCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","& sArrCateCode(iCnt) &")"
				        Response.Write "<p>"& sQuery
				        con.execute sQuery
				    Next
				end if'if UBound(sArrCateCode)=UBound(sArrClassCode) then
				
			'	sQuery = "INSERT INTO Pur_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType, "&_
			'			"SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
			'			"VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', "&_
			'			" "&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode & ") "
						
				sQuery = "INSERT INTO Pur_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType, "&_
						"SeriesNo, SeriesCode, NoSeriesStatus) "&_
						"VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', "&_
						" "&iSeries&", "&iInvSeriesCode&", '0') "
				Response.Write "<p>"& sQuery &"<br><br>"
				Con.Execute sQuery
				
				IF CStr(sItmType) = "0" Then
					sItemValue = "0"
				End IF
							
			'	sQuery = "INSERT INTO Pur_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, ItemType,"&_
			'			 "SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
			'			 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"',  "&_
			'			 " "&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode &") "
						 
				sQuery = "INSERT INTO Pur_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, ItemType,"&_
						 "SeriesNo, SeriesCode, NoSeriesStatus) "&_
						 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"',  "&_
						 " "&iSeries&", "&iInvSeriesCode&", '0') "
				Response.Write sQuery &"<br><br>"
				Con.Execute sQuery
			
						
			''Response.Write "End of Num Series Entry "&"<br><br>"
			End if			
		'********************************* New Insertion is Over **************************************************
	'To Check the Consered entry is for Amendment "Y" For Amendment 
		ElseIF CStr(TempNode.Item(iCount).Attributes.getNamedItem("EditCheck").value) = "Y" Then 
			Response.Write "Inside Amendment " &"<br>"
			
			iTransNo = TempNode.Item(iCount).Attributes.getNamedItem("TransNo").Value
			iSeries = TempNode.Item(iCount).Attributes.getNamedItem("SeriesNo").Value
			iInvSeriesCode = TempNode.Item(iCount).Attributes.getNamedItem("SeriesCode").Value
			
		'Helps to Update the Sratrt No Prefix and Suffix 
			
			
			sQuery = "Select ActivityType From Pur_M_Noseries Where NoSeriesTransactionNo = "&iTransNo&" "
			
			
			objrs.Open sQuery,Con
			IF not objrs.EOF Then
				iActivity = objrs(0)
			End IF
			objrs.Close
			
'************************************ Deletion of Old Values From the Table ************************************
			sQuery = "Delete from PUR_M_NoSeriesClass where SeriesNo="& iSeries &" and SeriesCode ="& iInvSeriesCode
			Response.Write "<p>"& sQuery
			con.execute sQuery
			
			sQuery = "Delete From Pur_M_NoSeriesAddDet Where NoSeriesTransactionNo = "&iTransNo&" "
			Response.Write "<p>"&sQuery
			Con.Execute sQuery
			sQuery = "Delete From Pur_M_NoSeriesDetails Where NoSeriesTransactionNo = "&iTransNo&" "
			Response.Write "<p>"&sQuery
			Con.Execute sQuery
			sQuery = "Delete From Pur_M_Noseries Where NoSeriesTransactionNo = "&iTransNo&" "
			Response.Write "<p>"&sQuery
			Con.Execute sQuery
			
			
'************************************ Deletion of Old Values From the Table Ends *******************************
			IF CStr(Left(TempNode.Item(iCount).Attributes.getNamedItem("NumFor").Value,1)) = "B" Then
				sNumFor = "B"
			ElseIF CStr(Left(TempNode.Item(iCount).Attributes.getNamedItem("NumFor").Value,1)) = "I" Then
				sNumFor = "I"
			Else
				sNumFor = "D"
			End IF
			
			IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("ItemTy").Value) = "Specific" Then
				sItmType = "1"
			Else
				sItmType = "0"
			End IF
			
			
			sItemValue = TempNode.Item(iCount).Attributes.getNamedItem("ItemValue").Value
			iEntryNo = 1
			
			UpdateNoSerValue iSeries,iInvSeriesCode,sTotEntNo,TempNode.Item(iCount)
			
'************************** Amendment Insertion Starts Here **************************************************
				IF CStr(sNumFor) = "B" and CStr(sItmType) = "0"  Then
					Response.Write "Inside All "
					sQuery = "INSERT INTO Pur_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
							 "NumberFor, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
							 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', "&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode &") "
									 
					Response.Write sQuery &"<br><br>"
					Con.Execute sQuery
					
					if UBound(sArrCateCode)=UBound(sArrClassCode) then
					    For iCnt = 0 to UBound(sArrClassCode)
					        sQuery = "Insert into PUR_M_NoSeriesClass (SeriesNo,SeriesCode,ClassCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","& sArrClassCode(iCnt) & ","& sArrCateCode(iCnt) &")"
					        Response.Write "<p>"& sQuery
					        con.execute sQuery
					    Next
					elseif UBound(sArrCateCode)>0 then
					    For iCnt = 0 to UBound(sArrCateCode)
					        sQuery = "Insert into PUR_M_NoSeriesClass (SeriesNo,SeriesCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","& sArrCateCode(iCnt) &")"
					        Response.Write "<p>"& sQuery
					        con.execute sQuery
					    Next
					end if'if UBound(sArrCateCode)=UBound(sArrClassCode) then
					
					sQuery = "INSERT INTO Pur_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType,"&_
						 " SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
						 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', "&_
						 " "&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode  &") "
											 
					Response.Write sQuery &"<br><br>"
					Con.Execute sQuery
				Elseif CStr(sItmType) = "1" or CStr(sItmType) = "0" Then
					Response.Write "Inside Other " &"<br>"
					sQuery = "INSERT INTO Pur_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
							 "NumberFor, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
							 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', "&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode &") "
													 
							Response.Write sQuery &"<br><br>"
							Con.Execute sQuery
							
							
							if UBound(sArrCateCode)=UBound(sArrClassCode) then
							    For iCnt = 0 to UBound(sArrClassCode)
							        sQuery = "Insert into PUR_M_NoSeriesClass (SeriesNo,SeriesCode,ClassCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","& sArrClassCode(iCnt) & ","& sArrCateCode(iCnt) &")"
							        Response.Write "<p>"& sQuery
							        con.execute sQuery
							    Next
							elseif UBound(sArrCateCode)>0 then
							    For iCnt = 0 to UBound(sArrCateCode)
							        sQuery = "Insert into PUR_M_NoSeriesClass (SeriesNo,SeriesCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","& sArrCateCode(iCnt) &")"
							        Response.Write "<p>"& sQuery
							        con.execute sQuery
							    Next
							end if'if UBound(sArrCateCode)=UBound(sArrClassCode) then
				
							
							sQuery = "INSERT INTO Pur_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType,"&_
								 " SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
								 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', "&_
								 " "&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode &") "
											 
							Response.Write sQuery &"<br><br>"
							Con.Execute sQuery
								
							IF CStr(sItmType) = "0" Then
								sItemValue = "0"
							End IF
							sQuery = "INSERT INTO Pur_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, ItemType, "&_
								 " SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
								 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"',  "&_
								 ""&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode &") "
											 
							Response.Write sQuery &"<br><br>"
							Con.Execute sQuery
	'					Next
						Response.Write "End of Num Series Entry "&"<br><br>"
				
				End IF 'Item Series Check 
			'End IF 'Activity Check 
'********************************************* Amendment Insertion Ends **************************************************
	
		End IF 'New Entry Check 
	Next 'Tempnode Loop
	'objfs.DeleteFile(server.MapPath("../Temp/master/NoSeries_PUR_"&Session.SessionID&".xml"))
End IF 'If objfs.FileExists(Server.MapPath("../Temp/master/NoSeries_PUR_"&Session.SessionID&".xml")) then

if con.Errors.count <> 0 then
	dim iErrCounter
	con.RollbackTrans
	for iErrCounter=0 to con.Errors.count
		'Response.Write con.Errors(iErrCounter) & vbCrLf
	next
	'Redirect to Error Handling System
else
'	con.RollbackTrans
'	Response.End 
	Response.Clear
	con.CommitTrans
end if

'con.close
set con = nothing
Response.Redirect "PurNoSeriesEntry.asp"
%>      

<%
	Function UpdateNoSerValue(iSerNo,iSerCode,iLoop,sRoot)
		Dim sStartNo,sPreVal,sSufVal,iCtloop,sStr,NoSerNode
		sStr = "//NumSeriesList[@SeriesNo="&iSerNo&" and @SeriesCode="&iSerCode&"]/NoList"
		Set NoSerNode = sRoot.selectNodes(sStr)
		IF NoSerNode.length <> 0 Then
			For iCtloop = 0 To NoSerNode.length - 1
				
				sStartNo = NoSerNode.Item(iCtloop).Attributes.Item(0).value
				sPreVal = NoSerNode.Item(iCtloop).Attributes.Item(1).value
				sSufVal = NoSerNode.Item(iCtloop).Attributes.Item(2).value
				
				
				sQuery = "UPDATE APP_R_NoSeriesModuleEntry SET  Number = "&sStartNo&", Prefix = '"&sPreVal&"', Suffix = '"&sSufVal&"' "&_
						 "WHERE OUDefinitionID = '"&iUnitNo&"' AND SeriesNo = "&iSerNo&" AND SeriesCode = "&iSerCode&" AND EntryNo = "&iCtloop+1&" "
						 
				Response.Write sQuery &"<br><br><br>"
						 
				Con.Execute sQuery	
			Next
		End IF
	End Function
	
%>
