<%
'InventoryAccountingUpdate.asp
%>
<%
Function InvAccountUpdate(sReceiptFor,iRcptNo,sOrgID,iAccountedBy)

Dim sSrcType,sQuery
Dim rsTemp1,dcrs,adoCmd
Dim ndStock,ndInvItem,ndStorage,ndLot
Dim ObjFS,oDOMInventory
Dim iMillSerNo,iPackNumber,iCounter
Dim iInvRecNo,iLotNumber,iEntryNo
Dim iRcptAs,sPartyCode,iClass,iInvItemCode
Set ObjFS = Server.CreateObject("Scripting.FileSystemObject")
Set oDOMInventory = Server.CreateObject("Microsoft.XMLDOM")

Set rsTemp1 = Server.CreateObject("ADODB.Recordset")
Set dcrs = Server.CreateObject("ADODB.Recordset")
	if ObjFS.FileExists(Server.MapPath("../temp/transaction/InventoryAcc_"&Session.SessionID&".xml")) then
		oDOMInventory.load(Server.MapPath("../temp/transaction/InventoryAcc_"&Session.SessionID&".xml"))
					if sReceiptFor = "09" or sReceiptFor = "08" then
						sSrcType = "RT"
					' Sub Contracting
					elseif sReceiptFor = "04" then
						sSrcType = "RC"
					' Job Work
					elseif sReceiptFor = "03" then
						sSrcType = "RW"
					' Purchase Return - Replacement
					elseif sReceiptFor = "02" then
						sSrcType = "RM"
					' Sales Return
					elseif sReceiptFor = "07" then
						sSrcType = "RN"
						with dcrs
							.CursorLocation = 3
							.CursorType = 3
							.Source = "SELECT ACTIONNO,SALESRETURNNO FROM SAL_T_SALESRETURNACTION WHERE ACTIONNO IN (4,5,6,7,8,9,10) AND SALESRETURNNO = (SELECT SALESRETURNNO FROM SAL_T_SALESRETURNREFDET WHERE RECEIPTNUMBER = " & iRcptNo & ")"
							.ActiveConnection = con
							.Open
						end with
						set dcrs.ActiveConnection = nothing
						if not dcrs.EOF then
							' Sales Return - To be replacement
							if trim(dcrs(0)) = "4" or trim(dcrs(0)) = "5" then
								sSrcType = "RL"
							' Sales Return - To be reworked,resent
							elseif trim(dcrs(0)) = "6" or trim(dcrs(0)) = "7" then
								sSrcType = "RK"
							' Sales Return - To be resent
							elseif trim(dcrs(0)) = "8" or trim(dcrs(0)) = "9" then
								sSrcType = "RE"
							' Sales Return
							elseif trim(dcrs(0)) = "10" then
								sSrcType = "RN"
								
									' Update the Action completed for this Sales Return
									sSql = "UPDATE SAL_T_SALESRETURNACTION SET ACTIONCOMPLETED = '1' " &_
										"WHERE SALESRETURNNO = " & trim(dcrs(1)) & " AND ACTIONNO = " & trim(dcrs(0)) & ""
									'Response.Write "<p>"& sSql & vbCrLf & vbCrLf
								con.Execute sSql
							end if
						end if
						dcrs.close

					else
						sSrcType = "R"
					end if
	
					set ndStock = oDOMInventory.documentElement
						ndStock.setAttribute "SRCTYPE",sSrcType 
						ndStock.setAttribute "TRANSACTIONTYPE",sSrcType 
					if ndStock.hasChildNodes() then
						for each ndInvItem in ndStock.childNodes 
							if trim(ndInvItem.nodeName)="ITEM" then
							    iEntryNo = ndInvItem.getAttribute("ITEMENTRYNO")
							    iInvItemCode = ndInvItem.getAttribute("ITEM")
							    iClass = ndInvItem.getAttribute("CLASS")
							    
							    
							    sSql = "UPDATE RCV_T_ACTUALRCPTITEMDET SET ITEMSTATUS = 'Y'" &_
										" WHERE ITEMCODE = " & iInvItemCode & " AND " &_
										"CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND " &_
										"RECEIPTNUMBER = " & iRcptNo & ""
										'response.write sSql
									'objTxt.Write sSql & vbCrLf & vbCrLf
									con.Execute sSql
			
								for each ndStorage in ndInvItem.childNodes
									if trim(ndStorage.nodeName)="STORAGE" then
										for each ndLot in ndStorage.childNodes
											if trim(ndLot.nodeName)="LOT" then
												iMillSerNo = ndLot.getAttribute("SERIALNO")
												
												sQuery = "Select isNull(MillPackingNumber,''),isNull(MillLotNo,'') from RCV_T_ActualRcptLotSerial where "&_
														 "MillSerialNo = "&  iMillSerNo &" and ReceiptNumber = "& iRcptNo  &" and EntryNO ="&iEntryNo
												'Response.Write "<p>"&sQuery 
												rsTemp.Open sQuery,con 
												if not rsTemp.EOF then
													iPackNumber = trim(rsTemp(0))
													iLotNumber = Trim(rsTemp(1))
												end if
												if trim(iPackNumber)="" or IsNull(iPackNumber) then	 iPackNumber = "NULL"
												if Trim(iLotNumber)="" or IsNull(iLotNumber) then iLotNumber = "N/A"
													ndLot.setAttribute "PACKINGNUMBER",iPackNumber 
													ndLot.setAttribute "LOT",iLotNumber
												rsTemp.Close 
											end if 'if trim(ndStorage.nodeName)="LOT" then
										next
									end if 'if trim(ndStorage.nodeName)="STORAGE" then
								next
							end if ' if trim(ndItem.nodeName)="ITEM" then
						next
					end if 'if ndStock.hasChildNodes() then

			oDOMInventory.save(Server.MapPath("../temp/transaction/InventoryAcc_"&Session.SessionID&".xml"))

			with dcrs 
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MAX(INVENTORYRECEIPTNO)+1,1) FROM INV_T_LOCATIONLOT"
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing
			if not dcrs.EOF then
				iInvRecNo = dcrs(0)
			end if
			dcrs.Close
			
			'Response.Write "<p>InvRecNo = "& iInvRecNo
					
			Set adoCmd = Server.CreateObject("ADODB.Command")
			Set adoCmd.ActiveConnection = con

			adoCmd.CommandText = "StockUpdation"
			adoCmd.CommandType = 4
			adoCmd.Parameters.Append adoCmd.CreateParameter("@XMLDoc",201,1,len(oDOMInventory.xml),oDOMInventory.xml)
			adoCmd.Execute()

			set adoCmd = nothing
					
			if trim(sSrcType)="RL" then
				sSql = "UPDATE SAL_T_SALESRETURNREFDET SET INVENTORYRECEIPTNO = " & iInvRecNo & " " &_
				" WHERE RECEIPTNUMBER = " & iRcptNo & " AND UNIT = " & Pack(sOrgID) & ""
				'Response.Write "<p>"& ssql
				con.Execute sSql
			end if
					
			sSql = "UPDATE RCV_T_ACTUALRECEIPTHEADER SET ACCOUNTEDBY = " & iAccountedBy & ", ACCOUNTEDON = CONVERT(DATETIME," & Pack(FormatDate(date)) & ",103),INVENTORYRECNO = " & iInvRecNo & ",STATUS='Accounted'" &_
			"WHERE RECEIPTNUMBER = " & iRcptNo & ""
			'Response.Write "<p>"& ssql
			con.execute sSql
			
			sQuery = "Select (Select PartyCode from RCV_T_GateReceiptHeader where GRNNumber =H.GRNNumber) as PartyCode from RCV_T_ActualReceiptHeader H where ReceiptNumber = " & iRcptNo 
			'Response.Write "<p>"& sQuery 
			rsTemp1.Open sQuery,con
			if not rsTemp1.EOF then
			    'iRcptAs = rsTemp1(0)
			    'sPartyCode = rsTemp1(1)
			    sPartyCode = rsTemp1(0)
			end if
			rsTemp1.Close 
			
			'sQuery = "Update INV_T_ItemLedger set PartyCode="& sPartyCode &",ReceiptType="& iRcptAs &" where TransactionNo = "& iInvRecNo &" and TransactionType = 'R'"
			sQuery = "Update INV_T_ItemLedger set PartyCode="& sPartyCode &" where TransactionNo = "& iInvRecNo &" and TransactionType = 'R'"
			'Response.Write "<p>"& sQuery
			con.execute sQuery 
			
			
	end if ' if ObjFS.FileExists(Server.MapPath("../temp/transaction/InventoryAcc.xml")) then

End Function
%>