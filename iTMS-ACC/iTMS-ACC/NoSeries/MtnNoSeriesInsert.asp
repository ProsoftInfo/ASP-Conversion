<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	MtnNoSeriesInsert.asp
	'Module Name				:	Maintenance (Master Creation)
	'Author Name				:	Kalaiselvi R
	'Created On					:	Jan 2,2009
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
Dim sItemValue,sSalType,sSalValue,sInvType,sInvValue,sAgTy,iTransNo,iEntryNo
Dim arrTemp,objDOM,objfs,Root,sExp,TempNode,iCount
Dim sSuffix,sPrefix,iNumber,sTotEntNo

Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")
Set objrs = Server.CreateObject("ADODB.RecordSet")
Set objrs1 = Server.CreateObject("ADODB.RecordSet")

con.BeginTrans

If objfs.FileExists(Server.MapPath("../Temp/master/NoSeries_MTN_"&Session.SessionID&".xml")) then
	objDOM.Load server.MapPath("../Temp/master/NoSeries_MTN_"&Session.SessionID&".xml")
	Set Root = objDOM.documentElement
	
	'Response.Write Request.Form +"<br>"
	
	iUnitNo=trim(Request.Form("selUnit"))
	iActivity=trim(Request.Form("selActType"))
	iSeries=trim(Request.Form("selNoSeries"))
	iSeriesType=trim(Request.Form("hSeriesType"))
	iLength=trim(Request.Form("hSeriesLen"))
	sActName = trim(Request.Form("hActivityName"))
	sTotEntNo = Request.Form("hTotEntNo")
	
	'Response.Write iSeries & "  " & iSeriesType +"<br>"
	
	
	sExp = "//NumSeriesList[@EditCheck!=""N""]"
	'Response.Write sExp
	Set TempNode = Root.selectNodes(sExp)
	
	For iCount = 0 To TempNode.length - 1
		'To check whether the entry is new entry "E" For New entry
		IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("EditCheck").value) = "E" Then
		
			
			IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("WorkGroupType").Value) = "Specific" Then
				sItmType = "1"
			Else
				sItmType = "0"
			End IF
			
			
			sItemValue = TempNode.Item(iCount).Attributes.getNamedItem("WorkGroupValue").Value
			sFromFin = Trim(Request.Form("hFinFrom"))
			sToFin = Trim(Request.Form("hFinTo"))

			iOldSrNo = Trim(Request.Form("hSeriesNo"))
			iOldSrCode = Trim(Request.Form("hSeriesCode"))
			'Response.Write iOldSrCode&iOldSrNo
			iTotalEntNo = Trim(Request.Form("hEntryNo"))

			
			if CStr(sItmType) = "1" or CStr(sItmType) = "0" Then
				'Response.Write "GenSeriesCode"
				iInvSeriesCode = GenSeriesCode(iUnitNo,"2","2",iSeries,iSeriesType,"",sActName,iLength)
				sQuery = "Select isNull(Max(NoSeriesTransactionNo),0) + 1 From MTN_M_Noseries "
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
				sQuery = "INSERT INTO MTN_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
						 "SeriesNo, SeriesCode, NoSeriesStatus) "&_
						 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', "&iSeries&", "&iInvSeriesCode&", '0') "
				Response.Write sQuery &"<br><br>"
				con.Execute sQuery
				sQuery = "INSERT INTO MTN_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, WorkGroup, "&_
						"SeriesNo, SeriesCode, NoSeriesStatus) "&_
						"VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', "&_
						" "&iSeries&", "&iInvSeriesCode&", '0') "
				Response.Write sQuery &"<br><br>"
				Con.Execute sQuery
				
				IF CStr(sItmType) = "0" Then
					sItemValue = "0"
				End IF
							
				sQuery = "INSERT INTO MTN_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, WorkGroupCode,"&_
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
			
			
			sQuery = "Select ActivityType From MTN_M_Noseries Where NoSeriesTransactionNo = "&iTransNo&" "
			
			
			objrs.Open sQuery,Con
			IF not objrs.EOF Then
				iActivity = objrs(0)
			End IF
			objrs.Close
			
'************************************ Deletion of Old Values From the Table ************************************
			sQuery = "Delete From MTN_M_NoSeriesAddDet Where NoSeriesTransactionNo = "&iTransNo&" "
			
			Con.Execute sQuery
			sQuery = "Delete From MTN_M_NoSeriesDetails Where NoSeriesTransactionNo = "&iTransNo&" "
			Con.Execute sQuery
			sQuery = "Delete From MTN_M_Noseries Where NoSeriesTransactionNo = "&iTransNo&" "
			Con.Execute sQuery
'************************************ Deletion of Old Values From the Table Ends *******************************
			
			
			IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("WorkGroupType").Value) = "Specific" Then
				sItmType = "1"
			Else
				sItmType = "0"
			End IF
			
			
			sItemValue = TempNode.Item(iCount).Attributes.getNamedItem("WorkGroupValue").Value
			iEntryNo = 1
			
			UpdateNoSerValue iSeries,iInvSeriesCode,sTotEntNo,TempNode.Item(iCount)
			
'************************** Amendment Insertion Starts Here **************************************************
				
				if CStr(sItmType) = "1" or CStr(sItmType) = "0" Then
					Response.Write "Inside Other " &"<br>"
					sQuery = "INSERT INTO MTN_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
							 "SeriesNo, SeriesCode, NoSeriesStatus) "&_
							 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', "&iSeries&", "&iInvSeriesCode&", '0') "
													 
							Response.Write sQuery &"<br><br>"
							Con.Execute sQuery
							sQuery = "INSERT INTO MTN_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, WorkGroup,"&_
								 " SeriesNo, SeriesCode, NoSeriesStatus) "&_
								 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', "&_
								 " "&iSeries&", "&iInvSeriesCode&", '0') "
											 
							Response.Write sQuery &"<br><br>"
							Con.Execute sQuery
								
							IF CStr(sItmType) = "0" Then
								sItemValue = "0"
							End IF
							sQuery = "INSERT INTO MTN_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, WorkGroupCode, "&_
								 " SeriesNo, SeriesCode, NoSeriesStatus) "&_
								 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"',  "&_
								 ""&iSeries&", "&iInvSeriesCode&", '0') "
											 
							Response.Write sQuery &"<br><br>"
							Con.Execute sQuery
	'					Next
						Response.Write "End of Num Series Entry "&"<br><br>"
				
				End IF 'Item Series Check 
			'End IF 'Activity Check 
'********************************************* Amendment Insertion Ends **************************************************
	
		End IF 'New Entry Check 
	Next 'Tempnode Loop
	'objfs.DeleteFile(server.MapPath("../Temp/master/NoSeries_MTN_"&Session.SessionID&".xml"))
End IF 'If objfs.FileExists(Server.MapPath("../Temp/master/NoSeries_MTN_"&Session.SessionID&".xml")) then

if con.Errors.count <> 0 then
	dim iErrCounter
	con.RollbackTrans
	for iErrCounter=0 to con.Errors.count
		'Response.Write con.Errors(iErrCounter) & vbCrLf
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	'Response.End
	
	con.CommitTrans
end if

set con = nothing
Response.Redirect "MtnNoSeriesEntry.asp"
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
