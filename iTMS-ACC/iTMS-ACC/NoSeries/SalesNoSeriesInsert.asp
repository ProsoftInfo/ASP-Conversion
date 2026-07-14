<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	SalesNoSeriesInsert.asp
	'Module Name				:	Sales (Master Creation)
	'Author Name				:	Subbiah
	'Created On					:	August 07,2003
	'Modified By				:	RAGAVENDRAN R
	'Modified On				:	MARCH 31,2010
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
Dim sSuffix,sPrefix,iNumber,sTotEntNo,sClassCode,sArrClassCode,iCnt,sCatCode,sArrCatCode

Set objDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")
Set objrs = Server.CreateObject("ADODB.RecordSet")
Set objrs1 = Server.CreateObject("ADODB.RecordSet")

con.BeginTrans


If objfs.FileExists(Server.MapPath("./Temp/master/NoSeries_SA_"&Session.SessionID&".xml")) then
	objDOM.Load server.MapPath("./Temp/master/NoSeries_SA_"&Session.SessionID&".xml")
	Set Root = objDOM.documentElement
	
	
	
	iUnitNo=trim(Request.Form("selUnit"))
	iActivity=trim(Request.Form("selActType"))
	iSeries=trim(Request.Form("selNoSeries"))
	iSeriesType=trim(Request.Form("hSeriesType"))
	iLength=trim(Request.Form("hSeriesLen"))
	sActName = trim(Request.Form("hActivityName"))
	sTotEntNo = Request.Form("hTotEntNo")
	sClassCode = Trim(Request.Form("hClassCode"))
	sCatCode = Trim(Request.Form("hCatCode"))
	
	
	
	if Trim(sClassCode)="" or IsNull(sClassCode) then sClassCode= "NULL"
	if Trim(sCatCode)="" or IsNull(sCatCode) then sCatCode = "NULL"
	
	sArrClassCode = Split(sClassCode,",")
	sArrCatCode = Split(sCatCode,",")
	'Response.Write iSeries & "  " & iSeriesType &"<br>"
	'Response.End 
	
	sExp = "//NumSeriesList[@EditCheck!=""N""]"
	
	Set TempNode = Root.selectNodes(sExp)
	
	For iCount = 0 To TempNode.length - 1
		'To check wheather the entry is new entry "E" For New entry
		IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("EditCheck").value) = "E" Then
		
			IF CStr(Left(TempNode.Item(iCount).Attributes.getNamedItem("NumFor").Value,1)) = "B" Then
				sNumFor = "B"
			ElseIF CStr(Left(TempNode.Item(iCount).Attributes.getNamedItem("NumFor").Value,1)) = "E" Then
				sNumFor = "E"
			Else
				sNumFor = "D"
			End IF
			
			IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("ItemTy").Value) = "Specific" Then
				sItmType = "1"
			Else
				sItmType = "0"
			End IF
			
			IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("InvTy").Value) = "Specific" Then
				sInvType = "1"
			Else
				sInvType = "0"
			End IF
			
			IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("SaleTy").Value) = "Specific" Then
				sSalType = "1"
			Else
				sSalType = "0"
			End IF
			
			IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("AgentTy").Value) = "Yes" Then
				sAgTy = "0"
			Else
				sAgTy = "1"
			End IF
			
			sItemValue = TempNode.Item(iCount).Attributes.getNamedItem("ItemValue").Value
			sInvValue = TempNode.Item(iCount).Attributes.getNamedItem("InvValue").Value
			sSalValue = TempNode.Item(iCount).Attributes.getNamedItem("SaleValue").Value
			sAgentcode = TempNode.Item(iCount).Attributes.getNamedItem("AgentCode").Value
			

			sFromFin = Trim(Request.Form("hFinFrom"))
			sToFin = Trim(Request.Form("hFinTo"))

			iOldSrNo = Trim(Request.Form("hSeriesNo"))
			iOldSrCode = Trim(Request.Form("hSeriesCode"))
			iTotalEntNo = Trim(Request.Form("hEntryNo"))

			

			''Response.Write "Calling "
			''Response.Write iTotalEntNo &" "& iOldSrNo &" "& iOldSrCode 

			
			''Response.Write sItmType &" >> " & sSalType &">> " & sInvType &" >> " & sAgTy &"<br>"
			''Response.Write iActivity &"<br>"

			IF CStr(iOldSrCode) = "" and CStr(iOldSrNo) = "" Then

				IF CStr(iActivity) <> "QUT" Then
					IF CStr(sNumFor) = "B" and CStr(sItmType) = "0" and CStr(sSalType) = "0" and CStr(sInvType) = "0" and CStr(sAgTy) = "1" Then
						''Response.Write "Inside All "
						iInvSeriesCode=GenSeriesCode(iUnitNo,"3","3",iSeries,iSeriesType,"",sActName,iLength)
						sQuery = "Select isNull(Max(NoSeriesTransactionNo),0) + 1 From Sal_M_Noseries "
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
					'	sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
					'			 "NumberFor, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
					'			 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', "&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode &") "
					
						sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
							     "NumberFor, SeriesNo, SeriesCode, NoSeriesStatus) "&_
							     "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', "&iSeries&", "&iInvSeriesCode&", '0') "
						Response.Write sQuery &"<br><br>"
						Con.Execute sQuery
						
						if UBound(sArrCatCode)=UBound(sArrClassCode) then
    						For iCnt = 0 to UBound(sArrClassCode)
						        sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,ClassCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrClassCode(iCnt)&","& sArrCatCode(iCnt) &") "
						        Response.Write "<p>"& sQuery
						        con.execute sQuery
						    Next
					    elseif UBound(sArrCatCode)>0 then
					        For iCnt = 0 to UBound(sArrCatCode)
						        sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrCatCode(iCnt)&") "
						        Response.Write "<p>"& sQuery
						        con.execute sQuery
						    Next
					    end if
						
					Elseif CStr(sItmType) = "1" or CStr(sSalType) = "1" or CStr(sInvType) = "1" or CStr(sItmType) = "0" or CStr(sSalType) = "0" or CStr(sInvType) = "0" Then
						IF CStr(sAgTy) = "0" Then
							'Response.Write "Inside Agent Insert "&"<br><br>"
							arrTemp = Split(sAgentcode,":")
							For iCtr = 0 To UBound(arrTemp)
								iInvSeriesCode=GenSeriesCode(iUnitNo,"3","3",iSeries,iSeriesType,arrTemp(iCtr),sActName,iLength)
								sQuery = "Select isNull(Max(NoSeriesTransactionNo),0) + 1 From Sal_M_Noseries "
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
							'	sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
							'			 "NumberFor, NoSeriesStatus,SeriesNo,SeriesCode,ClassificationCode) "&_
							'			 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', '0',"&iSeries&","& iInvSeriesCode &","& sClassCode &") "
										 
								sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
										 "NumberFor, NoSeriesStatus,SeriesNo,SeriesCode) "&_
										 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', '0',"&iSeries&","& iInvSeriesCode &") "
								Response.Write "<br><br>" & sQuery
								Con.Execute sQuery
								
								if UBound(sArrCatCode)=UBound(sArrClassCode) then
    						        For iCnt = 0 to UBound(sArrClassCode)
						                sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,ClassCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrClassCode(iCnt)&","& sArrCatCode(iCnt) &") "
						                Response.Write "<p>"& sQuery
						                con.execute sQuery
						            Next
					            elseif UBound(sArrCatCode)>0 then
					                For iCnt = 0 to UBound(sArrCatCode)
						                sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrCatCode(iCnt)&") "
						                Response.Write "<p>"& sQuery
						                con.execute sQuery
						            Next
					            end if
								
							'	sQuery = "INSERT INTO Sal_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
							'			 "SaleType, CommissionAgent, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
							'			 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', '"&sInvType&"', '"&sSalType&"', "&_
							'			 "'"&sAgTy&"', "&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode &") "
										 
								sQuery = "INSERT INTO Sal_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
										 "SaleType, CommissionAgent, SeriesNo, SeriesCode, NoSeriesStatus) "&_
										 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', '"&sInvType&"', '"&sSalType&"', "&_
										 "'"&sAgTy&"', "&iSeries&", "&iInvSeriesCode&", '0') "

								Response.Write "<br><br>aaa=" & sQuery
								Con.Execute sQuery
						
								IF CStr(sItmType) = "0" Then
									sItemValue = "0"
								End IF
						
								IF CStr(sInvType) = "0" Then
									sInvValue = "0"
								End IF
						
								IF CStr(sSalType) = "0" Then
									sSalValue = "0"
								End IF
							'	sQuery = "INSERT INTO Sal_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
							'			 "SaleType, AgentCode, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
							'			 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"', '"&sInvValue&"',  "&_
							'			 "'"&sSalValue&"', "&arrTemp(iCtr)&", "&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode  &") "
										 
								sQuery = "INSERT INTO Sal_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
										 "SaleType, AgentCode, SeriesNo, SeriesCode, NoSeriesStatus) "&_
										 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"', '"&sInvValue&"',  "&_
										 "'"&sSalValue&"', "&arrTemp(iCtr)&", "&iSeries&", "&iInvSeriesCode&", '0') "
								 
								Response.Write "<br><br>" & sQuery
								Con.Execute sQuery
							Next
							''Response.Write "End of Num Series Entry "&"<br><br>"
						Else 'Agent Check 
							''Response.Write "Outside Agent Insert "
							iInvSeriesCode=GenSeriesCode(iUnitNo,"3","3",iSeries,iSeriesType,"",sActName,iLength)
							sQuery = "Select isNull(Max(NoSeriesTransactionNo),0) + 1 From Sal_M_Noseries "
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
						'	sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
						'			 "NumberFor, NoSeriesStatus,SeriesNo,SeriesCode,ClassificationCode) "&_
						'			 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', '0',"&iSeries&","& iInvSeriesCode &","& sClassCode &") "
									 
							sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
									 "NumberFor, NoSeriesStatus,SeriesNo,SeriesCode) "&_
									 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', '0',"&iSeries&","& iInvSeriesCode &") "
							Response.Write sQuery &"<br><br><br>"
							Con.Execute sQuery
							
							if UBound(sArrCatCode)=UBound(sArrClassCode) then
    						    For iCnt = 0 to UBound(sArrClassCode)
						            sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,ClassCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrClassCode(iCnt)&","& sArrCatCode(iCnt) &") "
						            Response.Write "<p>"& sQuery
						            con.execute sQuery
						        Next
					        elseif UBound(sArrCatCode)>0 then
					            For iCnt = 0 to UBound(sArrCatCode)
						            sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrCatCode(iCnt)&") "
						            Response.Write "<p>"& sQuery
						            con.execute sQuery
						        Next
					        end if
						
							'IF the Agent Type is Yes/0 Then the Details SeriesCode and SeriesNo has to be Null
						'	sQuery = "INSERT INTO Sal_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
						'			 "SaleType, CommissionAgent, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
						'			 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', '"&sInvType&"', '"&sSalType&"', "&_
						'			 "'"&sAgTy&"', "&iSeries&", "&iInvSeriesCode&", '0'," & sClassCode &") "
									 
							sQuery = "INSERT INTO Sal_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
									 "SaleType, CommissionAgent, SeriesNo, SeriesCode, NoSeriesStatus) "&_
									 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', '"&sInvType&"', '"&sSalType&"', "&_
									 "'"&sAgTy&"', "&iSeries&", "&iInvSeriesCode&", '0') "
								 
							Response.Write "<br><br>" & sQuery
							Con.Execute sQuery
						
							IF CStr(sItmType) = "0" Then
								sItemValue = "0"
							End IF
						
							IF CStr(sInvType) = "0" Then
								sInvValue = "0"
							End IF
						
							IF CStr(sSalType) = "0" Then
								sSalValue = "0"
							End IF
						
						'	sQuery = "INSERT INTO Sal_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
						'			 "SaleType, AgentCode, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
						'			 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"', '"&sInvValue&"',  "&_
						'			 "'"&sSalValue&"', Null,"&iSeries&","& iInvSeriesCode &", '0',"& sClassCode &") "
									 
							sQuery = "INSERT INTO Sal_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
									 "SaleType, AgentCode, SeriesNo, SeriesCode, NoSeriesStatus) "&_
									 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"', '"&sInvValue&"',  "&_
									 "'"&sSalValue&"', Null,"&iSeries&","& iInvSeriesCode &", '0') "
								 
							Response.Write sQuery &"<br><br><br>"
							Con.Execute sQuery
						End IF 'Agent Check Ends
						'Response.Write "Outside Agent Insert "
					End IF
'****************************** End of Insertion Other Than Quotation ***********************************************
				Else 'End of Insertion of Other Activity than Quotation
				
				'Response.Write "Inside Quotataion Entry "
					
					'IF the Selected Number Series is For Quotation then the Series Has to be Created to all defined units.
						sQuery = "Select OUDefinitionID from DCS_OrganizationUnitDefinitions Where Len(OUDefinitionID) > 4 "
						
						Response.Write sQuery & "<br>"
						With objrs
							.CursorLocation = 3
							.CursorType = 3
							.ActiveConnection = Con
							.Source = sQuery
							.Open
						End With
						Set objrs.ActiveConnection = Nothing
						Do While Not objrs.EOF 
							iUnitNo = Trim(objrs(0))
							
							Response.Write sNumFor &" " & sItmType &" " & sSalType &" " & sInvType & sAgTy &"<br><br><br>"
							
							IF CStr(sNumFor) = "B" and CStr(sItmType) = "0" and CStr(sSalType) = "0" and CStr(sInvType) = "0" and CStr(sAgTy) = "1" Then
							
								Response.Write "Inissssssssssssss "
								iInvSeriesCode=GenSeriesCode(iUnitNo,"3","3",iSeries,iSeriesType,"",sActName,iLength)
								sQuery = "Select isNull(Max(NoSeriesTransactionNo),0) + 1 From Sal_M_Noseries "
								With objrs1
									.CursorLocation = 3
									.CursorType = 3
									.Source = sQuery
									.ActiveConnection = Con
									.Open
								End With
								Set objrs1.ActiveConnection = Nothing
								If not objrs1.EOF Then
									iTransNo = objrs1(0)
								End IF
								objrs1.Close
								
							'	sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
							'			 "NumberFor, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
							'			 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', "&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode &") "
										 
								sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
										 "NumberFor, SeriesNo, SeriesCode, NoSeriesStatus) "&_
										 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', "&iSeries&", "&iInvSeriesCode&", '0') "
										 
								Response.Write sQuery &"<Br><Br><Br>"
								 
								Con.Execute sQuery
								
								
								if UBound(sArrCatCode)=UBound(sArrClassCode) then
    						        For iCnt = 0 to UBound(sArrClassCode)
						                sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,ClassCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrClassCode(iCnt)&","& sArrCatCode(iCnt) &") "
						                Response.Write "<p>"& sQuery
						                con.execute sQuery
						            Next
					            elseif UBound(sArrCatCode)>0 then
					                For iCnt = 0 to UBound(sArrCatCode)
						                sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrCatCode(iCnt)&") "
						                Response.Write "<p>"& sQuery
						                con.execute sQuery
						            Next
					            end if
							
							Elseif CStr(sItmType) = "1" or CStr(sSalType) = "1" or CStr(sInvType) = "1" Then
						
								IF Cstr(sAgTy) = "0" Then ' IS Agent is Sel as Yes 
									arrTemp = Split(sAgentcode,":")
									For iCtr = 0 To UBound(arrTemp)
									
										iInvSeriesCode=GenSeriesCode(iUnitNo,"3","3",iSeries,iSeriesType,arrTemp(iCtr),sActName,iLength)
										
										sQuery = "Select isNull(Max(NoSeriesTransactionNo),0) + 1 From Sal_M_Noseries "
										With objrs1
											.CursorLocation = 3
											.CursorType = 3
											.Source = sQuery
											.ActiveConnection = Con
											.Open
										End With
										Set objrs1.ActiveConnection = Nothing
						
										If not objrs.EOF Then
											iTransNo = objrs1(0)
											
										End IF
										objrs1.Close
										iEntryNo = 1
								
									'	sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
									'			 "NumberFor, NoSeriesStatus,SeriesNo,SeriesCode,ClassificationCode) "&_
									'			 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', '0',"&iSeries&","&iInvSeriesCode&","& sClassCode &") "
										
										sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
												 "NumberFor, NoSeriesStatus,SeriesNo,SeriesCode) "&_
												 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', '0',"&iSeries&","&iInvSeriesCode&") "		 
										Response.Write sQuery &"<Br><Br><Br>"		 
										
										Con.Execute sQuery
		
										if UBound(sArrCatCode)=UBound(sArrClassCode) then
    						                For iCnt = 0 to UBound(sArrClassCode)
						                        sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,ClassCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrClassCode(iCnt)&","& sArrCatCode(iCnt) &") "
						                        Response.Write "<p>"& sQuery
						                        con.execute sQuery
						                    Next
					                    elseif UBound(sArrCatCode)>0 then
					                        For iCnt = 0 to UBound(sArrCatCode)
						                        sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrCatCode(iCnt)&") "
						                        Response.Write "<p>"& sQuery
						                        con.execute sQuery
						                    Next
					                    end if

										
									'	sQuery = "INSERT INTO Sal_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
									'			 "SaleType, CommissionAgent, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
									'			 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', '"&sInvType&"', '"&sSalType&"', "&_
									'			 "'"&sAgTy&"', "& iSeries &", "& iInvSeriesCode&", '0',"& sClassCode &") "
												 
										 sQuery = "INSERT INTO Sal_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
										 "SaleType, CommissionAgent, SeriesNo, SeriesCode, NoSeriesStatus) "&_
										 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', '"&sInvType&"', '"&sSalType&"', "&_
										 "'"&sAgTy&"', "& iSeries &", "& iInvSeriesCode&", '0') "
										Response.Write "<br><br>" & sQuery 
										Con.Execute sQuery
						
										IF CStr(sItmType) = "0" Then
											sItemValue = "0"
										End IF
						
										IF CStr(sInvType) = "0" Then
											sInvValue = "0"
										End IF
						
										IF CStr(sSalType) = "0" Then
											sSalValue = "0"
										End IF
										
									'	sQuery = "INSERT INTO Sal_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
									'			 "SaleType, AgentCode, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
									'			 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"', '"&sInvValue&"',  "&_
									'			 "'"&sSalValue&"', "&sAgentcode&", "&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode &") "
												 
												 sQuery = "INSERT INTO Sal_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
												 "SaleType, AgentCode, SeriesNo, SeriesCode, NoSeriesStatus) "&_
												 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"', '"&sInvValue&"',  "&_
												 "'"&sSalValue&"', "&sAgentcode&", "&iSeries&", "&iInvSeriesCode&", '0') "
										
										Response.Write sQuery &"<Br><Br><Br>"
										Con.Execute sQuery
									Next 'Agent Loop
								Else 'Agent Loop and IF Agent is Sel as NO
									iInvSeriesCode=GenSeriesCode(iUnitNo,"3","3",iSeries,iSeriesType,"",sActName,iLength)
									
									sQuery = "Select isNull(Max(NoSeriesTransactionNo),0) + 1 From Sal_M_Noseries "
									With objrs1
										.CursorLocation = 3
										.CursorType = 3
										.Source = sQuery
										.ActiveConnection = Con
										.Open
									End With
									Set objrs1.ActiveConnection = Nothing
						
									If not objrs.EOF Then
										iTransNo = objrs1(0)
									End IF
									objrs1.Close
									iEntryNo = 1
								
								'	sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
								'			 "NumberFor, NoSeriesStatus,SeriesNo,SeriesCode,ClassificationCode) "&_
								'			 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', '0',"&iSeries&","&iInvSeriesCode&","& sClassCode &") "
											 
											 sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
											 "NumberFor, NoSeriesStatus,SeriesNo,SeriesCode) "&_
											 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', '0',"&iSeries&","&iInvSeriesCode&") "
											 Response.Write "<p>"& sQuery 
									Con.Execute sQuery
									
									if UBound(sArrCatCode)=UBound(sArrClassCode) then
    						            For iCnt = 0 to UBound(sArrClassCode)
						                    sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,ClassCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrClassCode(iCnt)&","& sArrCatCode(iCnt) &") "
						                    Response.Write "<p>"& sQuery
						                    con.execute sQuery
						                Next
					                elseif UBound(sArrCatCode)>0 then
					                    For iCnt = 0 to UBound(sArrCatCode)
						                    sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrCatCode(iCnt)&") "
						                    Response.Write "<p>"& sQuery
						                    con.execute sQuery
						                Next
					                end if
									
								'	sQuery = "INSERT INTO Sal_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
								'			 "SaleType, CommissionAgent, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
								'			 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', '"&sInvType&"', '"&sSalType&"', "&_
								'			 "'"&sAgTy&"', "&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode &") "
								
											 sQuery = "INSERT INTO Sal_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
											 "SaleType, CommissionAgent, SeriesNo, SeriesCode, NoSeriesStatus) "&_
											 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', '"&sInvType&"', '"&sSalType&"', "&_
											 "'"&sAgTy&"', "&iSeries&", "&iInvSeriesCode&", '0') "
									Response.Write "<br><br>" & sQuery		 
									Con.Execute sQuery
						
									IF CStr(sItmType) = "0" Then
										sItemValue = "0"
									End IF
						
									IF CStr(sInvType) = "0" Then
										sInvValue = "0"
									End IF
						
									IF CStr(sSalType) = "0" Then
										sSalValue = "0"
									End IF
									
								'	sQuery = "INSERT INTO Sal_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
								'			 "SaleType, AgentCode, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
								'			 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"', '"&sInvValue&"',  "&_
								'			 "'"&sSalValue&"', Null, "& iSeries &", "& iInvSeriesCode &", '0',"& sClassCode &") "
											 
											 
									sQuery = "INSERT INTO Sal_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
											 "SaleType, AgentCode, SeriesNo, SeriesCode, NoSeriesStatus) "&_
											 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"', '"&sInvValue&"',  "&_
											 "'"&sSalValue&"', Null, "& iSeries &", "& iInvSeriesCode &", '0') "
									Response.Write sQuery
									Con.Execute sQuery
								End IF 'Agent Check Ends 
							End IF 'End For Inserion Checking 
							objrs.MoveNext
						Loop 'Unit Loop
						objrs.Close
					End IF 'End For Activity Check
				End IF 'End of Insert/Update Check 
'********************************* New Insertion is Over **************************************************
'To Check the Consered entry is for Amendment "Y" For Amendment 
		ElseIF CStr(TempNode.Item(iCount).Attributes.getNamedItem("EditCheck").value) = "Y" Then 
			'Response.Write "Inside Amendment " &"<br>"
			iTransNo = TempNode.Item(iCount).Attributes.getNamedItem("TransNo").Value
			iSeries = TempNode.Item(iCount).Attributes.getNamedItem("SeriesNo").Value
			iInvSeriesCode = TempNode.Item(iCount).Attributes.getNamedItem("SeriesCode").Value
			
		'Helps to Update the Sratrt No Prefix and Suffix 
			
			
			sQuery = "Select ActivityType From Sal_M_Noseries Where NoSeriesTransactionNo = "&iTransNo&" "
			
			
			objrs.Open sQuery,Con
			IF not objrs.EOF Then
				iActivity = objrs(0)
			End IF
			objrs.Close
			
'************************************ Deletion of Old Values From the Table ************************************
			sQuery = "Delete From Sal_M_NoSeriesAddDet Where NoSeriesTransactionNo = "&iTransNo&" "
			
			Response.Write "<p>"& sQuery
			con.execute sQuery
			sQuery = "Delete From Sal_M_NoSeriesDetails Where NoSeriesTransactionNo = "&iTransNo&" "
			Response.Write "<p>"& sQuery
			con.execute sQuery
			sQuery = "Delete From Sal_M_Noseries Where NoSeriesTransactionNo = "&iTransNo&" "
			Response.Write "<p>"& sQuery
			con.execute sQuery
			sQuery = "Delete from Sal_M_NoSeriesClass where SeriesNo = "& iSeries &" and SeriesCode ="& iInvSeriesCode
			Response.Write "<p>"& sQuery
			con.execute sQuery
'************************************ Deletion of Old Values From the Table Ends *******************************
			IF CStr(Left(TempNode.Item(iCount).Attributes.getNamedItem("NumFor").Value,1)) = "B" Then
				sNumFor = "B"
			ElseIF CStr(Left(TempNode.Item(iCount).Attributes.getNamedItem("NumFor").Value,1)) = "E" Then
				sNumFor = "E"
			Else
				sNumFor = "D"
			End IF
			
			IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("ItemTy").Value) = "Specific" Then
				sItmType = "1"
			Else
				sItmType = "0"
			End IF
			
			IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("InvTy").Value) = "Specific" Then
				sInvType = "1"
			Else
				sInvType = "0"
			End IF
			
			IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("SaleTy").Value) = "Specific" Then
				sSalType = "1"
			Else
				sSalType = "0"
			End IF
			
			IF CStr(TempNode.Item(iCount).Attributes.getNamedItem("AgentTy").Value) = "Yes" Then
				sAgTy = "0"
			Else
				sAgTy = "1"
			End IF
			
			sItemValue = TempNode.Item(iCount).Attributes.getNamedItem("ItemValue").Value
			sInvValue = TempNode.Item(iCount).Attributes.getNamedItem("InvValue").Value
			sSalValue = TempNode.Item(iCount).Attributes.getNamedItem("SaleValue").Value
			sAgentcode = TempNode.Item(iCount).Attributes.getNamedItem("AgentCode").Value
			iEntryNo = 1
			
			UpdateNoSerValue iSeries,iInvSeriesCode,sTotEntNo,sAgTy,sAgentcode,TempNode.Item(iCount)
			
'************************** Amendment Insertion Starts Here **************************************************
			IF CStr(iActivity) = "QUT" Then
				iUnitNo = "010101"
			End IF
				IF CStr(sNumFor) = "B" and CStr(sItmType) = "0" and CStr(sSalType) = "0" and CStr(sInvType) = "0" and CStr(sAgTy) = "1" Then
					'Response.Write "Inside All "
				'	sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
				'			 "NumberFor, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
				'			 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', "&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode &") "
							 
					sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
							 "NumberFor, SeriesNo, SeriesCode, NoSeriesStatus) "&_
							 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', "&iSeries&", "&iInvSeriesCode&", '0') "
									 
					Response.Write sQuery &"<br><br>"
					Con.Execute sQuery
					
					if UBound(sArrCatCode)=UBound(sArrClassCode) then
    						For iCnt = 0 to UBound(sArrClassCode)
						        sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,ClassCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrClassCode(iCnt)&","& sArrCatCode(iCnt) &") "
						        Response.Write "<p>"& sQuery
						        con.execute sQuery
						    Next
					    elseif UBound(sArrCatCode)>0 then
					        For iCnt = 0 to UBound(sArrCatCode)
						        sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrCatCode(iCnt)&") "
						        Response.Write "<p>"& sQuery
						        con.execute sQuery
						    Next
					    end if
					
				Elseif CStr(sItmType) = "1" or CStr(sSalType) = "1" or CStr(sInvType) = "1" or CStr(sItmType) = "0" or CStr(sSalType) = "0" or CStr(sInvType) = "0" Then
					'Response.Write "Inside Other " &"<br>"
					
					IF CStr(sAgTy) = "0" Then
						'Response.Write "Inside Agent Insert "&"<br><br>"
						arrTemp = Split(sAgentcode,":")
						For iCtr = 0 To UBound(arrTemp)
						'	sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
						'			 "NumberFor, NoSeriesStatus,SeriesNo,SeriesCode,ClassificationCode) "&_
						'			 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', '0',"& iSeries &","&iInvSeriesCode&","& sClassCode &") "
									 
							sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
									 "NumberFor, NoSeriesStatus,SeriesNo,SeriesCode) "&_
									 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', '0',"& iSeries &","&iInvSeriesCode&") "
													 
							Response.Write sQuery &"<br><br>"
							Con.Execute sQuery
							
							if UBound(sArrCatCode)=UBound(sArrClassCode) then
    						    For iCnt = 0 to UBound(sArrClassCode)
						            sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,ClassCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrClassCode(iCnt)&","& sArrCatCode(iCnt) &") "
						            Response.Write "<p>"& sQuery
						            con.execute sQuery
						        Next
					        elseif UBound(sArrCatCode)>0 then
					            For iCnt = 0 to UBound(sArrCatCode)
						            sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrCatCode(iCnt)&") "
						            Response.Write "<p>"& sQuery
						            con.execute sQuery
						        Next
					        end if
							
						'	sQuery = "INSERT INTO Sal_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
						'			 "SaleType, CommissionAgent, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
						'			 "VALUES ("&iTransNo&", 1, '"&sItmType&"', '"&sInvType&"', '"&sSalType&"', "&_
						'			 "'"&sAgTy&"', "& iSeries &", "&iInvSeriesCode&", '0',"& sClassCode &") "
									 
							sQuery = "INSERT INTO Sal_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
									 "SaleType, CommissionAgent, SeriesNo, SeriesCode, NoSeriesStatus) "&_
									 "VALUES ("&iTransNo&", 1, '"&sItmType&"', '"&sInvType&"', '"&sSalType&"', "&_
									 "'"&sAgTy&"', "& iSeries &", "&iInvSeriesCode&", '0') "
											 
							Response.Write "<br><br>" & sQuery
							Con.Execute sQuery
							
							
								
								
							IF CStr(sItmType) = "0" Then
								sItemValue = "0"
							End IF
									
							IF CStr(sInvType) = "0" Then
								sInvValue = "0"
							End IF
									
							IF CStr(sSalType) = "0" Then
								sSalValue = "0"
							End IF
								
						'	sQuery = "INSERT INTO Sal_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
						'			 "SaleType, AgentCode, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
						'			 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"', '"&sInvValue&"',  "&_
						'			 "'"&sSalValue&"', "&arrTemp(iCtr)&", "&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode &") "
									 
							sQuery = "INSERT INTO Sal_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
									 "SaleType, AgentCode, SeriesNo, SeriesCode, NoSeriesStatus) "&_
									 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"', '"&sInvValue&"',  "&_
									 "'"&sSalValue&"', "&arrTemp(iCtr)&", "&iSeries&", "&iInvSeriesCode&", '0') "
											 
							Response.Write sQuery &"<br><br>"
							Con.Execute sQuery
							
							'sQuery = "Select isNull(Max(NoSeriesTransactionNo),0) + 1 From Sal_M_Noseries "
							'With objrs
							'	.CursorLocation = 3
							'	.CursorType = 3
							'	.Source = sQuery
							'	.ActiveConnection = Con
							'	.Open
							'End With
							'Set objrs.ActiveConnection = Nothing
							'If not objrs.EOF Then
							'	iTransNo = objrs(0)
							'End IF
							'objrs.Close
							
						Next
						'Response.Write "End of Num Series Entry "&"<br><br>"
					Else 'Agent Check 
						'Response.Write "Outside Agent Insert " &"<br><br>"
					'	sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
					'			 "NumberFor, NoSeriesStatus,SeriesNo,SeriesCode,ClassificationCode) "&_
					'			 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', '0',"& iSeries &","&iInvSeriesCode&","&  sClassCode &") "
			
	
						sQuery = "INSERT INTO Sal_M_Noseries (NoSeriesTransactionNo, OrganisationCode, ActivityType, "&_
								 "NumberFor, NoSeriesStatus,SeriesNo,SeriesCode) "&_
								 "VALUES ("&iTransNo&", '"&iUnitNo&"', '"&iActivity&"', '"&sNumFor&"', '0',"& iSeries &","&iInvSeriesCode&") "

						Response.Write sQuery &"<br><br>"
						Con.Execute sQuery
						
													
						if UBound(sArrCatCode)=UBound(sArrClassCode) then
    						For iCnt = 0 to UBound(sArrClassCode)
						        sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,ClassCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrClassCode(iCnt)&","& sArrCatCode(iCnt) &") "
						        Response.Write "<p>"& sQuery
						        con.execute sQuery
						    Next
					    elseif UBound(sArrCatCode)>0 then
					        For iCnt = 0 to UBound(sArrCatCode)
						        sQuery = "Insert into Sal_M_NoSeriesClass (SeriesNo,SeriesCode,CatCode) Values("& iSeries &","& iInvSeriesCode &","&sArrCatCode(iCnt)&") "
						        Response.Write "<p>"& sQuery
						        con.execute sQuery
						    Next
					    end if

					
						'IF the Agent Type is Yes/0 Then the Details SeriesCode and SeriesNo has to be Null
					'	sQuery = "INSERT INTO Sal_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
					'			 "SaleType, CommissionAgent, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
					'			 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', '"&sInvType&"', '"&sSalType&"', "&_
					'			 "'"&sAgTy&"', "&iSeries&", "&iInvSeriesCode&", '0',"& sClassCode  &") "
								 
						sQuery = "INSERT INTO Sal_M_NoSeriesDetails (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
								 "SaleType, CommissionAgent, SeriesNo, SeriesCode, NoSeriesStatus) "&_
								 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItmType&"', '"&sInvType&"', '"&sSalType&"', "&_
								 "'"&sAgTy&"', "&iSeries&", "&iInvSeriesCode&", '0') "
								 
						Response.Write "<br><br>" & sQuery		 
						Con.Execute sQuery
								
						IF CStr(sItmType) = "0" Then
							sItemValue = "0"
						End IF
								
						IF CStr(sInvType) = "0" Then
							sInvValue = "0"
						End IF
								
						IF CStr(sSalType) = "0" Then
							sSalValue = "0"
						End IF
								
					'	sQuery = "INSERT INTO Sal_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
					'			 "SaleType, AgentCode, SeriesNo, SeriesCode, NoSeriesStatus,ClassificationCode) "&_
					'			 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"', '"&sInvValue&"',  "&_
					'			 "'"&sSalValue&"', Null, "& iSeries &", "& iInvSeriesCode &", '0',"& sClassCode &") "
								 
								 sQuery = "INSERT INTO Sal_M_NoSeriesAddDet (NoSeriesTransactionNo, EntryNo, ItemType, InvoiceType, "&_
								 "SaleType, AgentCode, SeriesNo, SeriesCode, NoSeriesStatus) "&_
								 "VALUES ("&iTransNo&", "&iEntryNo&", '"&sItemValue&"', '"&sInvValue&"',  "&_
								 "'"&sSalValue&"', Null, "& iSeries &", "& iInvSeriesCode &", '0') "
										 
						Response.Write sQuery &"<br><br><br>"
						Con.Execute sQuery
					End IF 'Agent Check Ends
				End IF 'Item Series Check 
			'End IF 'Activity Check 
'********************************************* Amendment Insertion Ends **************************************************
		End IF 'New Entry Check 
	Next 'Tempnode Loop
	objfs.DeleteFile(server.MapPath("./Temp/master/NoSeries_SA_"&Session.SessionID&".xml"))
End IF 'End of Objfs

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

con.close
set con = nothing	
Response.Redirect "SalesNoSeriesEntry.asp"
%>      


<%
	Function UpdateNoSerValue(iSerNo,iSerCode,iLoop,sAgTy,iAgCode,sRoot)
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
