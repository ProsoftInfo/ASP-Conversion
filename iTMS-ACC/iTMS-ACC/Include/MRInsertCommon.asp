<%
	'Program Name				:	MRInsertCommon.asp
	'Module Name				:	Include
	
%>
<!--#include file="mrsStatus.asp"-->
<!--#include file="NoSeries.asp"-->
<!--#include file="NoSeriesCommonFunctions.asp"-->
<%
Function MRInsert()
	dim newxml
	dim dcrs,dcrs1,dcrs2,sSql,objfs,RootNode,ItemNode,HeaderNode,ScheduleNode,ScheduleDetNode
	dim iItmCode,sOrgCode,iClassCode,dMRDate,sUsage,sRemarks,iApprover,iEntNo
	dim sTemp,sLoc,sBin,sMonYr,iQty,iDefinedBy,sUoM,sSchOn,iSchQty,sCC,sAttributeList
	dim arrFin,sFinFrom,sFinTo,arrTemp,sTempMonYr,sSchType,sSchValue,iSchNo,iSchEntNo,iSchDetCtr
	dim objFSO,objTxt,iCtr,iSchCtr,iMRSNo,sExp,sSrcRefNo,iRcptNo,iAccHead
	dim WCNode,MCNode,sMCCode,iMCQty,iWCounter,iMCounter,sWCCode,PCNode, iPCCounter, sPono, sPOQty
	dim iSeriesNo,iSeriesCode,sIssueCode,sItemRemarks,sRef,sIssueFor,sAddDetail
	Dim iParCode,sCallFrom,sRedirectTo,sItemRefNo
	Dim sAppRefType,sAppRefNo,sAppRefDate,sImmediateApprover
	Dim sAttID,sAttList,sArrList,sRequestedByUnit,iNumIssueClassCode,sTempSeries,sArrSeries
	Dim sIssToType,sIssToCode,sIssToSubCode,sIssueTypeCode,sNumClassName
	set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	

	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
	set objfs = Server.CreateObject("Scripting.FileSystemObject")
	response.write "<font color=red>"

	' Create our DOM Document Objects
	Set newxml = Server.CreateObject("Microsoft.XMLDOM")

	iDefinedBy = getUserid

	newxml.async = false
	
	newxml.load server.MapPath("../temp/transaction/MRS"& Session.SessionID &".xml")
	Set RootNode = newxml.documentElement
	
	sExp = "//ITEMDETAILS"
	set ItemNode = RootNode.selectSingleNode(sExp)
	iNumIssueClassCode = trim(ItemNode.Attributes.getNamedItem("CLASSCODE").value)

	sExp ="//HEADER"
	Set HeaderNode = RootNode.selectSingleNode(sExp)
	sOrgCode = trim(HeaderNode.Attributes.getNamedItem("FORUNIT").Value)
	dMRDate = trim(HeaderNode.Attributes.getNamedItem("CREATEDON").Value)
	
	'sUsage = trim(HeaderNode.Attributes.getNamedItem("USAGE").Value)
	sRemarks = trim(HeaderNode.Attributes.getNamedItem("REMARKS").Value)
	iApprover = trim(HeaderNode.Attributes.getNamedItem("APPROVER").Value)
	iRcptNo = trim(HeaderNode.Attributes.getNamedItem("RECEIPTNO").Value)
	iAccHead = trim(HeaderNode.Attributes.getNamedItem("ACCHEAD").Value)
	
	sCC = trim(HeaderNode.Attributes.getNamedItem("COSTCENTER").Value)
'	sIssueFor = trim(HeaderNode.Attributes.getNamedItem("ISSUEFOR").Value)
	sAddDetail =  trim(HeaderNode.Attributes.getNamedItem("REFTYPE").Value)
	'iParCode = trim(HeaderNode.Attributes.getNamedItem("PARTYCODE").Value)
	sAppRefType = trim(HeaderNode.getAttribute("AppRefType"))
	sAppRefNo = trim(HeaderNode.getAttribute("AppRefNo"))
	sAppRefDate = trim(HeaderNode.getAttribute("AppRefDate"))
	sCallFrom = trim(HeaderNode.getAttribute("CallFrom"))
	sRedirectTo = trim(HeaderNode.getAttribute("RedirectTo"))
	sImmediateApprover = trim(HeaderNode.getAttribute("ImmediateApprover"))
	iMRSNo = trim(HeaderNode.getAttribute("MRNo"))
	
	sRequestedByUnit = Trim(HeaderNode.getAttribute("RequestedByUnit"))
	
	sIssToType = trim(HeaderNode.getAttribute("ISSTOTYPE"))
	sIssToCode = trim(HeaderNode.getAttribute("ISSTOCODE"))
	sIssToSubCode = trim(HeaderNode.getAttribute("ISSTOSUBCODE"))
	sIssueTypeCode =  trim(HeaderNode.getAttribute("ISSUETYPECODE"))
	
	if trim(iApprover)="IM" then
	    iApprover = iDefinedBy
	end if
	
	if trim(sOrgCode)="" or IsNull(sOrgCode) then sOrgCode = session("organizationcode")
	
	if trim(sAppRefType)="" or IsNull(sAppRefType) then sAppRefType = "NULL"
	
	if trim(sAppRefNo)="" or IsNull(sAppRefNo) then sAppRefNo ="NULL"
	if sAppRefNo<>"NULL" then sAppRefNo = pack(sAppRefNo)
	
	if trim(sAppRefDate)="" or IsNull(sAppRefDate) then sAppRefDate ="NULL"
	if sAppRefDate<>"NULL" then sAppRefDate = pack(sAppRefDate)
	
	if Trim(dMRDate)<>"" then dMRDate = FormatDate(dMRDate)
	
	if trim(iParCode) ="" or IsNull(iParCode) then iParCode = "NULL"
	'Response.Write "sAddDetail="&dMRDate& vbCrLf 
	
	'Response.Write DateDiff("d",formatdate(FormatDate(date())),formatdate(dMRDate))

	'if CheckFinYear(dMRDate) <> "0" then
	'	Response.Write "N"
	'	Response.End
	'end if

	if sRemarks = "" or IsNull(sRemarks) then
		sRemarks = "NULL"
	else
		sRemarks = Pack(sRemarks)
	end if

	if iAccHead = "" or IsNull(iAccHead) or iAccHead = "select" then
		iAccHead = "NULL"
	end if
	
	if iApprover = "0" then iApprover = "NULL"

	if sCC = "select" then sCC = "NULL"
	
	if sIssueFor = "select" then
		sIssueFor = "NULL"
	else
		sIssueFor = Pack(sIssueFor)
	end if
	
'	if sUsage = "WIT" then
'		sRef = "1"
'	else
'		sRef = "0"
'	end if
	
	if sAddDetail = "Select" then 
		sAddDetail = "W" 
	end if 
	
	iCtr = 0
	
	''added by ragav on Dec 28,2011 for Inter Unit Trasfer demo purpose
'	if ucase(Trim(sUsage))="IUT" then
'	    sRequestedByUnit = pack(sRequestedByUnit)
'	else
'	    sRequestedByUnit = "NULL"
'	end if
'	''end 

if trim(sRequestedByUnit) ="" or IsNull(sRequestedByUnit) then sRequestedByUnit="NULL"
if trim(sRequestedByUnit)<>"NULL" then sRequestedByUnit=pack(sRequestedByUnit)

	sExp ="//ITEMDETAILS"
	Set ItemNode = RootNode.Selectnodes(sExp)
	if ItemNode.Length > 0 then
	
	
	    if trim(iMRSNo)="" then
    	    with dcrs
			    .CursorLocation = 3
			    .CursorType = 3
			    .Source = "SELECT ISNULL(MAX(MRSNUMBER)+1,1) FROM INV_T_MRSHEADER"
			    response.write "<p>"& dcrs.source
			    .ActiveConnection = con
			    .Open
		    end with
		    set dcrs.ActiveConnection = nothing
		    if not dcrs.EOF then
			    iMRSNo = dcrs(0)
		    end if
		    dcrs.Close
		else
		
		    sSql = "Delete from Inv_T_MRSAdditionalDetails where MRSNumber = "& iMRSNo
		    response.write "<p>"& ssql
		    con.Execute sSql
		    
            sSql = " Delete from Inv_T_MRSHeader where MRSNumber = "& iMRSNo
            response.write "<p>"& ssql
            con.Execute sSql
            
            sSql = " Delete from Inv_T_MRSIssuePick where MRSNumber = "& iMRSNo
            response.write "<p>"& ssql
            con.Execute sSql
            
            'sSql = " Delete from Inv_T_MRSIssuePickDetails where MRSNumber = "& iMRSNo
            'con.Execute sSql
            
            sSql = " Delete from Inv_T_MRSItemDetails where MRSNumber = "& iMRSNo
            response.write "<p>"& ssql
            con.Execute sSql
            
            sSql = " Delete from Inv_T_MRSItemPRSTSchedules where MRSNumber = "& iMRSNo
            response.write "<p>"& ssql
            con.Execute sSql
            
            sSql = " Delete from Inv_T_MRSItemSchedules where MRSNumber = "& iMRSNo
            response.write "<p>"& ssql
            con.Execute sSql
            
            sSql = " Delete from Inv_T_MRSItemSpecs where MRSNumber = "& iMRSNo
            response.write "<p>"& ssql
            con.Execute sSql
		end if

	'	if sUsage = "JWK" then
	'		sSrcRefNo = iRcptNo
	'	else
	'		sSrcRefNo = iMRSNo
	'	end if

		sIssueCode = "NULL"

		' Issue Number - Number Series
		' OrganisationCode,SeriesNumber,SeriesCode,Date
	'	with dcrs
	'		.CursorLocation = 3
	'		.CursorType = 3
	'		.Source = "SELECT SERIESNO,SERIESCODE FROM INV_M_NUMBERSERIES WHERE ACTIVITYTYPE = 'MR' AND ORGANISATIONCODE = " & Pack(sOrgCode) & ""
	'		.ActiveConnection = con
	'		.Open
	'	end with
	'	set dcrs.ActiveConnection = nothing
	'	if not dcrs.EOF then
	'		iSeriesNo = trim(dcrs(0))
	'		iSeriesCode = trim(dcrs(1))
	'		sIssueCode = GenSeriesNumber(sOrgCode,iSeriesNo,iSeriesCode,FormatDate(date()))
	'		sIssueCode = Pack(sIssueCode)
	'	end if
	'	dcrs.close
	
	
	    sTempSeries = GetInvNumberSeriesCodes("MR",sOrgCode,iNumIssueClassCode)
                sArrSeries = Split(sTempSeries,":")
                iSeriesNo = sArrSeries(0)
                iSeriesCode = sArrSeries(1)
                Response.write "<p>sOrgCode="& sOrgCode
                Response.write "<p>iSeriesNo = "&iSeriesNo
                Response.write "<p>iSeriesCode = "&iSeriesCode
                
                if Trim(iSeriesCode)="0" and Trim(iSeriesNo)="0" then
            	
	                sSql = "Select GroupName from INV_M_Classification where GroupCode = "& iNumIssueClassCode
	                dcrs.Open sSql,con
	                if not dcrs.EOF then
	                    sNumClassName = Trim(dcrs(0))
	                end if
	                dcrs.Close 
            	
	                Response.Clear 
	                Response.Write "<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p>&nbsp;<p><H2>Number Series is not defined for Material Requisition - "& sNumClassName &"  Classification</H2></p>"
	                Response.End 
	            else
		            sIssueCode=GenSeriesNumber(sOrgCode,iSeriesNo,iSeriesCode,FormatDate(date()))				   
	            end if
	
	
	if trim(sIssToType)="" or IsNull(sIssToType) then sIssToType = "NULL"
	if trim(sIssToType)<>"NULL" then sIssToTYpe = Pack(sIssToType)
	
	if trim(sIssToCode)="" or IsNull(sIssToCode) then sIssToCode = "NULL"
	if trim(sIssToCode)<>"NULL" then sIssToCode = Pack(sIssToCode)
	
	if trim(sIssToSubCode)="" or IsNull(sIssToSubCode) then sIssToSubCode = "NULL"
	if trim(sIssToSubCode)<>"NULL" then sIssToSubCode = Pack(sIssToSubCode)
	
	
		
		sSql = "INSERT INTO INV_T_MRSHEADER (MRSFORUNIT,MRSNUMBER,MRSDATE," &_
			"MRSTYPE,ISSTOTYPE,ISSTOCODE,ISSTOSUBCODE,REMARKS,CREATEDBY,CREATEDON,SOURCEREFNO,GENERATEDFROM,APPROVEDBY,MRSCODE,ACCOUNTHEAD,COSTCENTERHEAD,REFERENCE,AppRefType,AppRefNo,AppRefDate,RequestedByUnit,IssueTypeCode) VALUES " &_
			"(" & Pack(sOrgCode) & "," & iMRSNo & ",CONVERT(DATETIME," & Pack(dMRDate) & ",103),NULL," &_
			"" & sIssToType & ","& sIssToCode &","& sIssToSubCode &"," & sRemarks & "," & iDefinedBy & "," &_
			"CONVERT(DATETIME,'"& date() &"',103)," & Pack(sSrcRefNo) & ",4," & iApprover & "," & PACK(sIssueCode) & "," & iAccHead & "," & sCC & ",NULL,"& sAppRefType  &","& sAppRefNo &","& sAppRefDate &","& sRequestedByUnit &","& pack(sIssueTypeCode) &")"
		response.write "<p>"& ssql
		con.Execute sSql

		For iCtr = 0 to ItemNode.Length - 1
			iEntNo = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ENTRYNO").Value)
			iItmCode = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ITEMCODE").Value)
			iClassCode = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("CLASSCODE").Value)
			sUoM = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("UOM").Value)
			iQty = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("QTY").Value)
			sItemRemarks = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("REMARKS").Value)
			sAttributeList = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ATTRIBUTELIST").Value)
			sItemRefNo = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("RefNo").Value)
			if sItemRemarks = "" then
				sItemRemarks = "NULL"
			else
				sItemRemarks = Pack(sItemRemarks)
			end if

			iWCounter = 0
			iPCCounter = 0
			iMCounter = 0
			if trim(sAttributeList)<>"" then
			    sArrList = split(sAttributeList,":")
			    sAttList = split(sArrList(0),"#")
			    sAttID = sAttList(1)
			end if
			
			
			if trim(sAttID)="" or IsNull(sAttID) then  sAttID = "NULL"
			if trim(sAttID)<>"NULL"  then sAttID = Pack(sAttID)
            'sExp ="//ITEMDETAILS [@ITEMCODE = " & iItmCode & " and @ENTRYNO = " & iEntNo & "]/Schedule [@ITEMCODE = " & iItmCode & "  and @SCHENTRYNO = " & iEntNo & "]"
			sExp ="//ITEMDETAILS [@ITEMCODE = " & iItmCode & " and @ENTRYNO = " & iEntNo & "]/Schedule [@ITEMCODE = " & iItmCode & "]"
			Set ScheduleNode = RootNode.Selectnodes(sExp)
			if ScheduleNode.Length > 0 then
			iSchCtr = 0
				for iSchCtr = 0 to ScheduleNode.Length - 1
					sSchType = trim(ScheduleNode.Item(iSchCtr).Attributes.getNamedItem("STYPE").Value)
					sSchValue = trim(ScheduleNode.Item(iSchCtr).Attributes.getNamedItem("SVALUE").Value)
					iSchEntNo = trim(ScheduleNode.Item(iSchCtr).Attributes.getNamedItem("SCHENTRYNO").Value)
					if sSchType = "S" then
						sSchValue = "NULL"
					elseif sSchType = "I" then
						sSchValue = Pack(dMRDate)
					else
						sSchValue = Pack(sSchValue)
					end if
					
					sSql = "INSERT INTO INV_T_MRSITEMDETAILS (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE," &_
						"ITEMCODE,REQUIREDBY,REQUIREDVALUE,QUANTITYREQUESTED,ITEMREMARKS,ICOUNTER,ITEMATTRIBUTES) VALUES " &_
						"(" & iMRSNo & "," & Pack(sOrgCode) & "," & iClassCode & "," & iItmCode & "," &_
						"" & Pack(sSchType) & "," & sSchValue & "," & iQty & "," & sItemRemarks & "," & iEntNo & "," & sAttID & ")"
					response.write "<p>"& ssql
					con.Execute sSql
					
					
					if trim(sAppRefType)="14" then ' Mix Code
    					sSql = "Insert into Inv_T_MRSAdditionalDetails (MRSNumber,OrganisationCode,ClassificationCode,ItemCode,MixCode,QuantityIssued)"&_
    					       " values ("& iMRSNo &",'"& sOrgCode &"',"& iClassCode &","& iItmCode &","& sItemRefNo &","& iQty &")"
    					response.write "<p>"& ssql
    					con.Execute sSql
    			    end if
					
					

					' Function Call to Update the Line Status of an MR for Inventory Application
					MRLineStatusUpdate "Requisition","Create",iMRSNo,iItmCode,iClassCode,iEntNo,sOrgCode,"4","","",""
					
					if trim(sImmediateApprover)="Y" then
					    MRLineStatusUpdate "Requisition","Approved",iMRSNo,iItmCode,iClassCode,iEntNo,sOrgCode,"4","","",""
					    
					    sSql = " UPDATE INV_T_MRSITEMDETAILS set QuantityApproved = QUANTITYREQUESTED where MRSNUMBER = "& iMRSNo &" and ITEMCODE = "& iItmCode &" and ClassificationCode = "& iClassCode
					    Response.Write sSql & vbCrLf & vbCrLf 
					    con.execute sSql
					    
					end if

					sExp ="//ITEMDETAILS [@ITEMCODE = " & iItmCode & " and @ENTRYNO = " & iEntNo & "]/Schedule [@ITEMCODE = " & iItmCode & "  and @SCHENTRYNO = " & iEntNo & "]/ScheduleDetails"
					Set ScheduleDetNode = RootNode.Selectnodes(sExp)

					if ScheduleDetNode.Length > 0 then
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
									response.write "<p>"& dcrs.source
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
								response.write "<p>"& ssql
								con.Execute sSql
							end if
						next 'for iSchDetCtr = 0 to ScheduleDetNode.Length - 1
					end if 'if ScheduleDetNode.Length > 0 then
				next 'for iSchCtr = 0 to ScheduleNode.Length - 1
			end if  'if ScheduleNode.Length > 0 then

			' Work Center Based / Maintenance based
			if sUsage = "WIP" or sUsage = "MAT" then
				sExp ="//ITEMDETAILS [ @ITEMCODE = "&iItmCode&"  and @ENTRYNO = " & iEntNo & "]/AddDet/WorkCenter"
				Set WCNode = RootNode.Selectnodes(sExp)
				For iWCounter = 0 to WCNode.Length - 1
					sWCCode = trim(WCNode.Item(iWCounter).Attributes.getNamedItem("WCODE").Value)

					sExp ="//ITEMDETAILS [ @ITEMCODE = "&iItmCode&"  and @ENTRYNO = " & iEntNo & "]/AddDet/WorkCenter [ @WCODE = '"&sWCCode&"']/MachineCenter"
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
							response.write "<p>"& ssql
							con.Execute sSql
						next
					else
						sSql = "INSERT INTO INV_T_MRSADDITIONALDETAILS (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
							"WORKCENTERCODE,QUANTITYISSUED,REFTYPE) VALUES " &_
							"(" & iMRSNo & "," & Pack(sOrgCode) & "," & iClassCode & "," & iItmCode & "," &_
							"" & Pack(sWCCode) & "," & iQty & ",'W')"
						response.write "<p>"& ssql
						con.Execute sSql
					end if
				next
			' Packing Based
			elseif sUsage = "PAC" then
				sExp ="//ITEMDETAILS [ @ITEMCODE = "&iItmCode&"  and @ENTRYNO = " & iEntNo & "]/PackingDet/PCode"
				Set PCNode = RootNode.Selectnodes(sExp)
				For iPCCounter = 0 to PCNode.Length - 1
					sPono = trim(PCNode.Item(iPCCounter).Attributes.getNamedItem("PONO").Value)
					sPOQty = trim(PCNode.Item(iPCCounter).Attributes.getNamedItem("QTY").Value)

					sSql = "INSERT INTO INV_T_MRSADDITIONALDETAILS (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
						"PRODUCTIONORDERNO,QUANTITYISSUED,REFTYPE) VALUES " &_
						"(" & iMRSNo & "," & Pack(sOrgCode) & "," & iClassCode & "," & iItmCode & "," &_
						"" & Pack(sPono) & "," & sPOQty & ",'P')"
					response.write "<p>"& ssql
					con.Execute sSql
				next
			' Mixing Based
			elseif sUsage = "PRD" then
				sExp ="//ITEMDETAILS [ @ITEMCODE = "&iItmCode&"  and @ENTRYNO = " & iEntNo & "]/MixDet/MCode"
				Set MCNode = RootNode.Selectnodes(sExp)
				For iMCounter = 0 to MCNode.Length - 1
					sMCCode = trim(MCNode.Item(iMCounter).Attributes.getNamedItem("MIXCODE").Value)
					iMCQty = trim(MCNode.Item(iMCounter).Attributes.getNamedItem("QTY").Value)

					sSql = "INSERT INTO INV_T_MRSADDITIONALDETAILS (MRSNUMBER,ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE," &_
						"MIXCODE,QUANTITYISSUED,REFTYPE) VALUES " &_
						"(" & iMRSNo & "," & Pack(sOrgCode) & "," & iClassCode & "," & iItmCode & "," &_
						"" & Pack(sMCCode) & "," & iMCQty & ",'M')"
					response.write "<p>"& ssql
					con.Execute sSql
				next
			end if

		next

		' Function Call to Update the Header Status of an MR
		MRStatusUpdate "Requisition","Create",iMRSNo,iItmCode,iClassCode,iEntNo,sOrgCode
		
		if trim(sImmediateApprover)="Y" then
		    MRStatusUpdate "Requisition","Approved",iMRSNo,iItmCode,iClassCode,iEntNo,sOrgCode
		end if

	end if
	
'	if objfs.FileExists(server.MapPath("../temp/transaction/MRS"&Session.SessionID&".xml")) then
'	    objfs.DeleteFile server.MapPath("../temp/transaction/MRS"&Session.SessionID&".xml")
'   end if
	
	MRInsert = iMRSNo 
End Function 'Function MRInsert()
%>
