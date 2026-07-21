<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	MRApprovalInsert.asp
	'Module Name				:	Inventory (Transaction)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	August 11, 2005
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
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/mrsStatus.asp"-->
<%
	dim newxml
	dim dcrs,dcrs1,dcrs2,sSql,objfs,RootNode,ItemNode,HeaderNode,ScheduleNode,ScheduleDetNode
	dim iItmCode,sOrgCode,iClassCode,dMRDate,sMRType,sUsage,sRemarks,iApprover,sItemRemarks
	dim sTemp,sLoc,sBin,sMonYr,iQty,iDefinedBy,sUoM,sSchOn,iSchQty,iSchDetCtr,iSchEntNo
	dim arrFin,sFinFrom,sFinTo,arrTemp,sTempMonYr,sSchType,sSchValue,iSchNo
	dim objFSO,objTxt,iCtr,iSchCtr,iMRSNo,sExp,iMRSAmdNo,sAction,sTempAction
	dim WCNode,MCNode,sMCCode,iMCQty,iWCounter,iMCounter,sWCCode,PCNode, iPCCounter, sPono, sPOQty
	dim sLotCardNo,sMachineNo,sCC,sAddDetail,iEntNo,sAttributeList,sMRAction
	Dim sIssToCode,sIssToType,sIssToSubCode,sIssueTypeCode
	set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	set objTxt = objFSO.CreateTextFile(server.MapPath("../temp/transaction/MRApproval.txt"))

	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

	' Create our DOM Document Objects
	Set newxml = Server.CreateObject("Microsoft.XMLDOM")

	iDefinedBy = getUserid
	iApprover = iDefinedBy
	sMRAction =  Request.QueryString("Action")
	
	newxml.async = false
	newxml.load(Request)

	newxml.Save server.MapPath("../temp/transaction/MRApproval.xml")
	'Response.Write newxml.xml
	'Response.End

	Set RootNode = newxml.documentElement

	sExp ="//HEADER"
	Set HeaderNode = RootNode.selectSingleNode(sExp)
	sOrgCode = trim(HeaderNode.Attributes.getNamedItem("FORUNIT").Value)
	dMRDate = trim(HeaderNode.Attributes.getNamedItem("CREATEDON").Value)
	sMRType = trim(HeaderNode.Attributes.getNamedItem("TYPE").Value)
	sUsage = trim(HeaderNode.Attributes.getNamedItem("USAGE").Value)
	sRemarks = trim(HeaderNode.Attributes.getNamedItem("REMARKS").Value)
	iMRSNo = trim(HeaderNode.Attributes.getNamedItem("MRNO").Value)
	sLotCardNo = trim(HeaderNode.Attributes.getNamedItem("LOTCARDNO").Value)
	sMachineNo = trim(HeaderNode.Attributes.getNamedItem("MACHINENO").Value)
	sCC = trim(HeaderNode.Attributes.getNamedItem("COSTCENTER").Value)
	sAddDetail	= trim(HeaderNode.Attributes.getNamedItem("REFTYPE").Value)
	sIssToType 	= trim(HeaderNode.Attributes.getNamedItem("ISSTOTYPE").Value)
	sIssToCode 	= trim(HeaderNode.Attributes.getNamedItem("ISSTOCODE").Value)
	sIssToSubCode 	= trim(HeaderNode.Attributes.getNamedItem("ISSTOSUBCODE").Value)
	sIssueTypeCode	= trim(HeaderNode.Attributes.getNamedItem("ISSUETYPECODE").Value)
	
	if sRemarks = "" or IsNull(sRemarks) then
		sRemarks = "NULL"
	else
		sRemarks = Pack(sRemarks)
	end if
	if sLotCardNo = "" or IsNull(sLotCardNo) then
		sLotCardNo = "NULL"
	else
		sLotCardNo = Pack(sLotCardNo)
	end if

	if sMachineNo = "" or IsNull(sMachineNo) then
		sMachineNo = "NULL"
	else
		sMachineNo = Pack(sMachineNo)
	end if

	if sCC = "select" then sCC = "NULL"
	
	if sAddDetail = "" then sAddDetail = "W" 
	
	iCtr = 0
	
	if trim(sIssToType)="" or IsNull(sIssToType) then sIssToType = "NULL" 
	if trim(sIssToType)<>"NULL" then sIssToType= pack(sIssToType)
	
	if trim(sIssToCode)="" or IsNull(sIssToCode) then sIssToCode = "NULL" 
	if trim(sIssToCode)<>"NULL" then sIssToCode= pack(sIssToCode)
	
	if trim(sIssToSubCode)="" or IsNull(sIssToSubCode) then sIssToSubCode = "NULL" 
	if trim(sIssToSubCode)<>"NULL" then sIssToSubCode= pack(sIssToSubCode)
	
	con.beginTrans

	with dcrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(MAX(MRSAMENDNUMBER)+1,1) FROM INV_A_MRSHEADER"
		.ActiveConnection = con
		.Open
	end with
	set dcrs1.ActiveConnection = nothing

	if not dcrs1.EOF then
		iMRSAmdNo = dcrs1(0)
		sSql = "INSERT INTO INV_A_MRSHEADER EXECUTE ('SELECT " & iMRSAmdNo & " AS MRSAMENDNUMBER," &_
			"CONVERT(DATETIME,GETDATE(),103) AS MRSAMENDDATE,REMARKS AS AMENDREMARKS," &_
			"" & iApprover & " AS AMENDENDBY,CONVERT(DATETIME,GETDATE(),103) AS AMENDENDON," &_
			"''AP'' AS AMENDORAPPROVE,* FROM INV_T_MRSHEADER WHERE MRSNUMBER = " & iMRSNo & "') "
		'Response.Write sSql & vbCrLf & vbCrLf
    	con.Execute sSql

		sSql = "INSERT INTO INV_A_MRSITEMDETAILS EXECUTE ('SELECT " & iMRSAmdNo & "," &_
			"* FROM INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNo & " ') "
		 Response.Write sSql & vbCrLf & vbCrLf
		con.Execute sSql

		sSql = "INSERT INTO INV_A_MRSITEMSCHEDULES EXECUTE ('SELECT " & iMRSAmdNo & "," &_
			"* FROM INV_T_MRSITEMSCHEDULES WHERE MRSNUMBER = " & iMRSNo & "') "
		'Response.Write sSql & vbCrLf & vbCrLf
		'con.Execute sSql

		sSql = "INSERT INTO INV_A_MRSITEMSPECS EXECUTE ('SELECT " & iMRSAmdNo & "," &_
			"* FROM INV_T_MRSITEMSPECS WHERE MRSNUMBER = " & iMRSNo & "') "
		'Response.Write sSql & vbCrLf & vbCrLf
		'con.Execute sSql

		sSql = "INSERT INTO INV_A_MRSSTOCKTRANSFER EXECUTE ('SELECT " & iMRSAmdNo & "," &_
			"* FROM INV_T_MRSSTOCKTRANSFER WHERE MRSNUMBER = " & iMRSNo & "') "
		'Response.Write sSql & vbCrLf & vbCrLf
		'con.Execute sSql
'Response.Write "  XML = " &RootNode.xml

		sSql = "UPDATE INV_T_MRSHEADER SET APPROVEDBY = " & iApprover & ", APPROVEDON = CONVERT(DATETIME,'"& date() &"',103),LOTCARDNO = " & sLotCardNo & ",MACHINENO = " & sMachineNo & ",COSTCENTERHEAD = " & sCC & ",IssToType="& sIssToType &",IssToCode="& sIssToCode &",IssToSubCode="& sIssToSubCode &",IssueTypeCode="& pack(sIssueTypeCode) &" WHERE MRSNUMBER = " & iMRSNo & ""
		Response.write sSql& vbcrlf
		con.Execute sSql
		sExp ="//ITEMDETAILS"
		Set ItemNode = RootNode.Selectnodes(sExp)
		if ItemNode.Length > 0 then

			For iCtr = 0 to ItemNode.Length - 1
				iEntNo = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ENTRYNO").Value)
				iItmCode = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ITEMCODE").Value)
				iClassCode = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("CLASSCODE").Value)
		
					sSql = "DELETE INV_T_MRSADDITIONALDETAILS WHERE MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ITEMCODE = " & iItmCode & ""
					'Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql
								
					sSql = "DELETE INV_T_MRSITEMSPECS WHERE MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ITEMCODE = " & iItmCode & ""
					con.Execute sSql

					sSql = "DELETE INV_T_MRSITEMSCHEDULES WHERE MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ITEMCODE = " & iItmCode & ""
					con.Execute sSql

					sSql = "DELETE INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & iMRSNo & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ITEMCODE = " & iItmCode & ""
					'Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql
			Next
		End IF
		sExp ="//ITEMDETAILS"
		Set ItemNode = RootNode.Selectnodes(sExp)
		if ItemNode.Length > 0 then

			For iCtr = 0 to ItemNode.Length - 1
				iEntNo = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ENTRYNO").Value)
				iItmCode = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ITEMCODE").Value)
				iClassCode = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("CLASSCODE").Value)
				sUoM = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("UOM").Value)
				iQty = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("QTY").Value)
				sSchType = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("REQUIREDBY").Value)
				sSchValue = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("REQUIREDVALUE").Value)
				sAction = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("DISPLAYED").Value)
				sAttributeList = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ATTRIBUTELIST").Value)
				sItemRemarks = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("REMARKS").Value)
				'Response.Write "<BR><P>iEntNo="&iEntNo &"<BR><BR>"
				sTempAction = sTempAction & sAction

				if sItemRemarks = "" then
					sItemRemarks = "NULL"
				else
					sItemRemarks = Pack(sItemRemarks)
				end if

				if sSchType = "S" then
					sSchValue = "NULL"
				elseif sSchType = "I" then
					sSchValue = Pack(dMRDate)
				else
					sExp ="//ITEMDETAILS [@ITMCODE = " & iItmCode & " and @ENTRYNO = " & iEntNo & "]/Schedule [@ITEMCODE = " & iItmCode & " and @SCHENTRYNO = " & iEntNo & " ]"
					Set ScheduleNode = RootNode.Selectnodes(sExp)
					if ScheduleNode.Length > 0 then
						sSchValue = Pack(trim(ScheduleNode.Item(0).Attributes.getNamedItem("SVALUE").Value))
					end if
				end if
				
				iWCounter = 0
				iPCCounter = 0
				iMCounter = 0
				
				IF sSchType<>"S" THEN ' sSchType start
					
					sSql = "INSERT INTO INV_T_MRSITEMDETAILS (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
							"ITEMCODE,REQUIREDBY,REQUIREDVALUE,QUANTITYREQUESTED,QUANTITYAPPROVED,ITEMREMARKS,ICOUNTER,ITEMATTRIBUTES) VALUES " &_
							"(" & iMRSNo & "," & Pack(sOrgCode) & "," & iClassCode & "," & iItmCode & "," &_
							"" & Pack(sSchType) & "," & sSchValue & "," & iQty&"," & iQty & "," & sItemRemarks & "," & iEntNo & ",'" & sAttributeList & "')"
						Response.Write sSql & vbCrLf & vbCrLf
						con.Execute sSql
						
                    if Trim(sMRAction)="Approve" then
					    if sAction = "N" then
						    ' Function Call to Update the Line Status of an MR for Inventory Application
						    MRLineStatusUpdate "Requisition","Approved",iMRSNo,iItmCode,iClassCode,iEntNo,sOrgCode,"4","","",""
					    elseif sAction = "Y" then
						    ' Function Call to Update the Line Status of an MR for Inventory Application
						    MRLineStatusUpdate "Requisition","Rejected",iMRSNo,iItmCode,iClassCode,iEntNo,sOrgCode,"4","","",""
					    end if
					end if 'if triM(sMRAction)="Approve" then
				ELSE
				
					sExp ="//ITEMDETAILS [@ITEMCODE = " & iItmCode & " and @ENTRYNO = " & iEntNo & "]/Schedule [@ITEMCODE = " & iItmCode & "  and @SCHENTRYNO = " & iEntNo & "]/ScheduleDetails"
					Set ScheduleDetNode = RootNode.Selectnodes(sExp)
					if ScheduleDetNode.Length > 0 then

						sSql = "INSERT INTO INV_T_MRSITEMDETAILS (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
							"ITEMCODE,REQUIREDBY,REQUIREDVALUE,QUANTITYREQUESTED,QUANTITYAPPROVED,ITEMREMARKS,ICOUNTER,ITEMATTRIBUTES) VALUES " &_
							"(" & iMRSNo & "," & Pack(sOrgCode) & "," & iClassCode & "," & iItmCode & "," &_
							"" & Pack(sSchType) & "," & sSchValue & "," & iQty&"," & iQty & "," & sItemRemarks & "," & iEntNo & ",'" & sAttributeList & "')"
						Response.Write sSql & vbCrLf & vbCrLf
						con.Execute sSql
						
						if triM(sMRAction)="Approve" then
							if sAction = "N" then
								' Function Call to Update the Line Status of an MR for Inventory Application
								MRLineStatusUpdate "Requisition","Approved",iMRSNo,iItmCode,iClassCode,iEntNo,sOrgCode,"4","","",""
							elseif sAction = "Y" then
								' Function Call to Update the Line Status of an MR for Inventory Application
								MRLineStatusUpdate "Requisition","Rejected",iMRSNo,iItmCode,iClassCode,iEntNo,sOrgCode,"4","","",""
							end if
					    end if 'if triM(sMRAction)="Approve" then
					    
						iSchDetCtr = 0
						for iSchDetCtr = 0 to ScheduleDetNode.Length - 1
							sSchOn = trim(ScheduleDetNode.Item(iSchDetCtr).Attributes.getNamedItem("NEED").Value)
							iSchQty = trim(ScheduleDetNode.Item(iSchDetCtr).Attributes.getNamedItem("QTY").Value)
							sSchType = trim(ScheduleDetNode.Item(iSchDetCtr).Attributes.getNamedItem("TYPE").Value)
							iSchEntNo = trim(ScheduleDetNode.Item(iSchDetCtr).Attributes.getNamedItem("SNO").Value)	
							if not sSchOn = "" then
								with dcrs
									.CursorLocation = 3
									.CursorType = 3
									.Source = "SELECT ISNULL(MAX(SCHEDULENO)+1,1) FROM INV_T_MRSITEMSCHEDULES WHERE MRSNUMBER = " & iMRSNo & ""
									.ActiveConnection = con
									.Open
								end with
								set dcrs.ActiveConnection = nothing
								iSchNo = dcrs(0)
								dcrs.Close

								sSql = "INSERT INTO INV_T_MRSITEMSCHEDULES (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
									"ITEMCODE,SCHEDULENO,SCHEDULETYPE,SCHEDULEDON,SCHEDULEDQTY,ITEMENTRYNO) VALUES " &_
									"(" & iMRSNo & "," & Pack(sOrgCode) & "," & iClassCode & "," & iItmCode & "," &_
									"" & iSchEntNo & "," & Pack(sSchType) & "," & Pack(sSchOn) & "," & iSchQty & "," & iEntNo & ")"
								Response.Write sSql & vbCrLf & vbCrLf
								con.Execute sSql
							end if
						next 'for iSchDetCtr = 0 to ScheduleDetNode.Length - 1
						
						MRLineStatusUpdate "Requisition","Create",iMRSNo,iItmCode,iClassCode,iEntNo,sOrgCode,"4","","",""
						
					end if 'if ScheduleDetNode.Length > 0 then
				END IF 'sSchType End
''''''''''''''''''''''''''''''''''''

				' Work Center Based / Maintenance based
				
				if sUsage = "WIP" or sUsage = "MAT" then
					sExp ="//ITEMDETAILS [ @ITEMCODE = "&iItmCode&" and @ENTRYNO = " & iEntNo & "]/AddDet/WorkCenter"
					Set WCNode = RootNode.Selectnodes(sExp)
					For iWCounter = 0 to WCNode.Length - 1
						sWCCode = trim(WCNode.Item(iWCounter).Attributes.getNamedItem("WCODE").Value)

						sExp ="//ITEMDETAILS [ @ITEMCODE = "&iItmCode&"]/AddDet/WorkCenter [ @WCODE = '"&sWCCode&"']/MachineCenter"
						Set MCNode = RootNode.Selectnodes(sExp)
						if MCNode.length > 0 then
							For iMCounter = 0 to MCNode.Length - 1
								sMCCode = trim(MCNode.Item(iMCounter).Attributes.getNamedItem("MCODE").Value)
								iMCQty = trim(MCNode.Item(iMCounter).Attributes.getNamedItem("QTY").Value)

								if sMCCode = "select" then
									sMCCode = "NULL"
								else
									sMCCode = Pack(sMCCode)
								end if

								sSql = "INSERT INTO INV_T_MRSADDITIONALDETAILS (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
									"WORKCENTERCODE,MACHINECENTERCODE,QUANTITYISSUED,REFTYPE) VALUES " &_
									"(" & iMRSNo & "," & Pack(sOrgCode) & "," & iClassCode & "," & iItmCode & "," &_
									"" & Pack(sWCCode) & "," & sMCCode & "," & iMCQty & "," & Pack(sAddDetail) & ")"
								'Response.Write sSql & vbCrLf & vbCrLf
								con.Execute sSql
							next
						else
							sSql = "INSERT INTO INV_T_MRSADDITIONALDETAILS (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
								"WORKCENTERCODE,QUANTITYISSUED,REFTYPE) VALUES " &_
								"(" & iMRSNo & "," & Pack(sOrgCode) & "," & iClassCode & "," & iItmCode & "," &_
								"" & Pack(sWCCode) & "," & iQty & ",'W')"
							'Response.Write sSql & vbCrLf & vbCrLf
							con.Execute sSql
						end if
					next
				' Packing Based	
				elseif sUsage = "PAC" then
					sExp ="//ITEMDETAILS [ @ITEMCODE = "&iItmCode&" and @ENTRYNO = " & iEntNo & "]/PackingDet/PCode"
					Set PCNode = RootNode.Selectnodes(sExp)
					For iPCCounter = 0 to PCNode.Length - 1
						sPono = trim(PCNode.Item(iPCCounter).Attributes.getNamedItem("PONO").Value)
						sPOQty = trim(PCNode.Item(iPCCounter).Attributes.getNamedItem("QTY").Value)

						sSql = "INSERT INTO INV_T_MRSADDITIONALDETAILS (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
							"PRODUCTIONORDERNO,QUANTITYISSUED,REFTYPE) VALUES " &_
							"(" & iMRSNo & "," & Pack(sOrgCode) & "," & iClassCode & "," & iItmCode & "," &_
							"" & Pack(sPono) & "," & sPOQty & ",'P')"
						'Response.Write sSql & vbCrLf & vbCrLf
						con.Execute sSql
					next
				' Mixing Based
				elseif sUsage = "PRD" then
					sExp ="//ITEMDETAILS [ @ITEMCODE = "&iItmCode&" and @ENTRYNO = " & iEntNo & "]/MixDet/MCode"
					Set MCNode = RootNode.Selectnodes(sExp)
					For iMCounter = 0 to MCNode.Length - 1
						sMCCode = trim(MCNode.Item(iMCounter).Attributes.getNamedItem("MIXCODE").Value)
						iMCQty = trim(MCNode.Item(iMCounter).Attributes.getNamedItem("QTY").Value)

						sSql = "INSERT INTO INV_T_MRSADDITIONALDETAILS (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
							"MIXCODE,QUANTITYISSUED,REFTYPE) VALUES " &_
							"(" & iMRSNo & "," & Pack(sOrgCode) & "," & iClassCode & "," & iItmCode & "," &_
							"" & Pack(sMCCode) & "," & iMCQty & ",'M')"
						'Response.Write sSql & vbCrLf & vbCrLf
						con.Execute sSql
					next
				end if

			next
			
		'	Response.Write " sTempAction = " &sTempAction 
		'	Response.Write " iMRSNo = "& iMRSNo 
		
		    if triM(sMRAction)="Approve" then
			    if InStr(1,sTempAction,"N") > 0 and InStr(1,sTempAction,"Y") = 0 then
				    ' Function Call to Update the Header Status of an MR
				    MRStatusUpdate "Requisition","Approved",iMRSNo,iItmCode,iClassCode,iEntNo,sOrgCode
			    elseif InStr(1,sTempAction,"Y") > 0 and InStr(1,sTempAction,"N") = 0 then
				    ' Function Call to Update the Header Status of an MR
				    MRStatusUpdate "Requisition","Rejected",iMRSNo,iItmCode,iClassCode,iEntNo,sOrgCode
			    elseif InStr(1,sTempAction,"N") > 0 and InStr(1,sTempAction,"Y") > 0 then
				    ' Function Call to Update the Header Status of an MR
				    MRStatusUpdate "Requisition","Partial Approved",iMRSNo,iItmCode,iClassCode,iEntNo,sOrgCode
			    end if
			end if 'if triM(sMRAction)="Approve" then
		end if
	end if
	dcrs1.Close
	'RootNode.appendChild Root
	'newxml.Save server.MapPath("../xmldata/transaction/MRAPPROVAL"&iMRSNo&".xml")
Response.Clear 
'Response.End 
	if con.Errors.count <> 0 then
		dim iCounter
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) & vbCrLf
		next
		'Redirect to Error Handling System
	else
	'	con.RollbackTrans
	'	Response.End
		
		Response.clear
		con.CommitTrans
	end if

	con.close
	set con = nothing
	

%>
