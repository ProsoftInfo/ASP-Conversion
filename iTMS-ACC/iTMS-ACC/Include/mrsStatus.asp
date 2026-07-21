<%
	'Program Name				:	mrsStatus.asp
	'Module Name				:	Inventory (MR Status)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	September 08, 2003
	'Modified On				:
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
	'Counters					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>

<%	' Function for Updating the Line Status of an MR
	'1.	Activity (Requisition or Issue or Pick)
	'2.	Action (Create, Approve, Reject, Amend, Hold, Short Close, Cancel)
	'3.	Internally generated MR number
	'4.	Item Code
	'5.	Classification Code
	'6.	Organization Code
	'7.	Application code (Sales, Inventory, etc) from which application the MR has been created
	'8.	ISS Type – FIRM / MARKED
	'9.	SCH type – Schedule Type (Single or Multiple)

	Function MRLineStatusUpdate(sActivity,sAction,iMRNumber,iItemCode,iClassCode,iEntNo,sUnitCode,iApplication,sIssueType,sScheduleType,iItemCtr)
		dim dcrs,sSql
		dim iQuantity,iSTQuantity,iPRQuantity,iISSQuantity,iPurdQuantity,iPurdQuantityRej

		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		if iItemCtr = "" then iItemCtr = "NULL"
		
		if sActivity = "Requisition" and sAction = "Create" then
			' Status Code - 040101 (Created)
			sSql = "UPDATE INV_T_MRSITEMDETAILS SET MRSITEMSTATUS = '040101' WHERE " &_
				"MRSNUMBER = " & iMRNumber & " AND ORGANISATIONCODE = " & Pack(sUnitCode) & " AND " &_
				"CLASSIFICATIONCODE = " & iClassCode & " AND ITEMCODE = " & iItemCode & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
			Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if

		if sActivity = "Requisition" and sAction = "Amend" then
			' Status Code - 040107 (Changed)
			sSql = "UPDATE INV_T_MRSITEMDETAILS SET MRSITEMSTATUS = '040107' WHERE " &_
				"MRSNUMBER = " & iMRNumber & " AND ORGANISATIONCODE = " & Pack(sUnitCode) & " AND " &_
				"CLASSIFICATIONCODE = " & iClassCode & " AND ITEMCODE = " & iItemCode & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
			Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if

		if sActivity = "Requisition" and sAction = "Approved" then
			' Status Code - 040102 (Approved)
			sSql = "UPDATE INV_T_MRSITEMDETAILS SET MRSITEMSTATUS = '040102' WHERE " &_
				"MRSNUMBER = " & iMRNumber & " AND ORGANISATIONCODE = " & Pack(sUnitCode) & " AND " &_
				"CLASSIFICATIONCODE = " & iClassCode & " AND ITEMCODE = " & iItemCode & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
			Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if

		if sActivity = "Requisition" and sAction = "Rejected" then
			' Status Code - 040103 (Rejected)
			sSql = "UPDATE INV_T_MRSITEMDETAILS SET MRSITEMSTATUS = '040103' WHERE " &_
				"MRSNUMBER = " & iMRNumber & " AND ORGANISATIONCODE = " & Pack(sUnitCode) & " AND " &_
				"CLASSIFICATIONCODE = " & iClassCode & " AND ITEMCODE = " & iItemCode & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
			Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if

		if sActivity = "Issue" and sAction = "Create" and sIssueType = "F" then
			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				If iMRNumber <> "" then 
					.Source = "SELECT QUANTITYAPPROVED,ISNULL(QUANTITYISSUED,0),ISNULL(QUANTITYFORTRANSFER,0),ISNULL(QUANTITYTOPURCHASE,0),ISNULL(QUANTITYPURCHASED,0) FROM INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & iMRNumber & " AND ORGANISATIONCODE = " & Pack(sUnitCode) & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ITEMCODE = " & iItemCode & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
				Else
					.Source = "SELECT QUANTITYAPPROVED,ISNULL(QUANTITYISSUED,0),ISNULL(QUANTITYFORTRANSFER,0),ISNULL(QUANTITYTOPURCHASE,0),ISNULL(QUANTITYPURCHASED,0) FROM INV_T_MRSITEMDETAILS WHERE  ORGANISATIONCODE = " & Pack(sUnitCode) & " AND CLASSIFICATIONCODE = " & iClassCode & " AND ITEMCODE = " & iItemCode & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
				End If
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				iQuantity = cdbl(dcrs(0))
				iISSQuantity = cdbl(dcrs(1))
				iSTQuantity = cdbl(dcrs(2))
				iPRQuantity = cdbl(dcrs(3))
				iPurdQuantity = cdbl(dcrs(4))

				Response.Write iISSQuantity
				'if (iISSQuantity + iSTQuantity + iPRQuantity) = iQuantity then
				if (iISSQuantity) = iQuantity then
					'Status Code - 040111 (Issued)
					if iMRNumber <> "" then 
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040111' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040111' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if
				elseif (iISSQuantity > 0 and iPRQuantity = 0 and iSTQuantity = 0) then
					if iMRNumber <> "" then 
						'Partially Issued
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040112' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040112' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & "  AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if
				elseif (iISSQuantity > 0 and iPRQuantity > 0 and iSTQuantity = 0 and iPurdQuantity = 0) then
					if iMRNumber <> "" then 
						'Partially Issued / PR
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040113' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040113' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if			
				elseif (iISSQuantity > 0 and iPRQuantity > 0 and iSTQuantity = 0 and iPurdQuantity > 0 and iPRQuantity = iPurdQuantity) then
					if iMRNumber <> "" then 
					'Partially Issued / Purchased
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040122' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040122' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if
				elseif (iISSQuantity > 0 and iPRQuantity > 0 and iSTQuantity = 0 and iPurdQuantity > 0 and iPRQuantity > iPurdQuantity) then
					if iMRNumber <> "" then 
						'Partially Issued / Partially Purchased
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040123' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040123' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if
				elseif (iISSQuantity > 0 and iPRQuantity = 0 and iSTQuantity > 0) then
					if iMRNumber <> "" then 
						'Partially Issued / ST
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040114' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
							"MRSITEMSTATUS = '040114' WHERE " &_
							"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
							"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if
				elseif (iISSQuantity = 0 and iPRQuantity > 0 and iSTQuantity > 0 and iPurdQuantity = 0) then
					if iMRNumber <> "" then 
						'Partially PR / ST
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040115' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040115' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & "  AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if
				elseif (iISSQuantity = 0 and iPRQuantity > 0 and iSTQuantity > 0 and iPurdQuantity > 0 and iPRQuantity > iPurdQuantity) then
					if iMRNumber <> "" then 
						'Partially Purchased / ST
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040126' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else	
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
							"MRSITEMSTATUS = '040126' WHERE " &_
							"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
							"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if
				elseif (iISSQuantity = 0 and iPRQuantity > 0 and iSTQuantity > 0 and iPurdQuantity > 0 and iPRQuantity = iPurdQuantity) then
					if iMRNumber <> "" then 
						'Partially ST / Purchased
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040127' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040127' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & "   AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if
				elseif (iISSQuantity > 0 and iSTQuantity > 0 and iPRQuantity > 0 and iPurdQuantity = 0) then
					if iMRNumber <> "" then 
					'Partially Issued / ST / PR
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040116' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040116' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & "   AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if		
				elseif (iISSQuantity > 0 and iSTQuantity > 0 and iPRQuantity > 0 and iPurdQuantity > 0 and iPRQuantity > iPurdQuantity) then
					if iMRNumber <> "" then 
						'Partially Issued / ST / Partially Purchased
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040128' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else			
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040128' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & "   AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if			
				elseif (iISSQuantity > 0 and iSTQuantity > 0 and iPRQuantity > 0 and iPurdQuantity > 0 and iPRQuantity = iPurdQuantity) then
					if iMRNumber <> "" then 
						'Partially Issued / ST / Purchased
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040129' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else			
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040129' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & "  AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if
				elseif (iISSQuantity = 0 and iSTQuantity = 0 and iPRQuantity > 0 and iPurdQuantity = 0) then
					if iMRNumber <> "" then 
						'Purchase Request
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040117' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else	
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040117' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & "   AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if
				elseif (iISSQuantity = 0 and iSTQuantity = 0 and iPRQuantity > 0 and iPurdQuantity > 0 and iPRQuantity = iPurdQuantity) then
					if iMRNumber <> "" then 
						'Purchased
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040124' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040124' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"			
					end if
				elseif (iISSQuantity = 0 and iSTQuantity = 0 and iPRQuantity > 0 and iPurdQuantity > 0 and iPRQuantity > iPurdQuantity) then
					if iMRNumber <> "" then 
						'Partially Purchased
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040125' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else	
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040125' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & "  AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if
				elseif (iISSQuantity = 0 and iSTQuantity > 0 and iPRQuantity = 0) then
					if iMRNumber <> "" then 
						'Request For Transfer
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040118' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040118' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if							
				else
					if iMRNumber <> "" then 
						'Partially Issued
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040112' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & " AND MRSNUMBER = " & iMRNumber & " AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					else
						sSql = "UPDATE INV_T_MRSITEMDETAILS SET " &_
								"MRSITEMSTATUS = '040112' WHERE " &_
								"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & " AND " &_
								"ORGANISATIONCODE = " & Pack(sUnitCode) & "  AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL)"
					end if
				end if
				Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql
			end if
			dcrs.close
		end if

	End Function
%>

<%	' Function for Updating the Header Status of an MR
	'1.	Activity (Requisition or Issue or Pick)
	'2.	Action (Create, Approve, Reject, Amend, Hold, Short Close, Cancel)
	'3.	Internally generated MR number
 
	Function MRStatusUpdate(sActivity,sAction,iMRNumber,sItemCode,sClass,iEntNo,sOrgID)

		dim dcrs,sSql,iItemCode,iClassCode,sUnitCode,iCtr
		dim iQuantity,iSTQuantity,iPRQuantity,iISSQuantity,iPurdQuantity

		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		
'		with dcrs
'			.CursorLocation = 3
'			.CursorType = 3
'			.Source = "SELECT MRSNUMBER FROM INV_T_MRSITEMDETAILS WHERE ITEMCODE = " & sItemCode & " AND CLASSIFICATIONCODE = " & sClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & "  AND (ICOUNTER = " & iEntNo & " OR ICOUNTER IS NULL) and MRSNUMBER =  "& iMRNumber
'			.ActiveConnection = con
'			.Open
'		end with
'		'Response.Write "CHECK ="& dcrs.Source
'		If not dcrs.eof then 
'			iMRNumber = dcrs(0)
'		End If 
'		dcrs.close
	'Response.Write "<p>MRSNUMBER = "& iMRNumber &"<BR><BR>"
		if sActivity = "Requisition" and sAction = "Create" then
			' Status Code - 040101 (Created)
			sSql = "UPDATE INV_T_MRSHEADER SET MRSHEADERSTATUS = '040101' WHERE " &_
				"MRSNUMBER = " & iMRNumber & ""
			Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if

		if sActivity = "Requisition" and sAction = "Amend" then
			' Status Code - 040107 (Changed)
			sSql = "UPDATE INV_T_MRSHEADER SET MRSHEADERSTATUS = '040107' WHERE " &_
				"MRSNUMBER = " & iMRNumber & ""
			Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if

		if sActivity = "Requisition" and sAction = "Approved" then
			' Status Code - 040102 (Approved)
			sSql = "UPDATE INV_T_MRSHEADER SET MRSHEADERSTATUS = '040102' WHERE " &_
				"MRSNUMBER = " & iMRNumber & ""
			Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MRSISSUESTATUS,'-') FROM INV_T_MRSHEADER WHERE MRSNUMBER = " & iMRNumber & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				if trim(dcrs(0)) = "-" or trim(dcrs(0)) = "040105" then
					' Status Code - 040102 (Approved)
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040102' WHERE " &_
						"MRSNUMBER = " & iMRNumber & ""
					Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql
				end if
			end if
			dcrs.close

		end if

		if sActivity = "Requisition" and sAction = "Rejected" then
			' Status Code - 040103 (Rejected)
			sSql = "UPDATE INV_T_MRSHEADER SET MRSHEADERSTATUS = '040103' WHERE " &_
				"MRSNUMBER = " & iMRNumber & ""
			Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if

		if sActivity = "Requisition" and sAction = "Partial Approved" then
			' Status Code - 040105 (Partially Approved)
			sSql = "UPDATE INV_T_MRSHEADER SET MRSHEADERSTATUS = '040105' WHERE " &_
				"MRSNUMBER = " & iMRNumber & ""
			Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MRSISSUESTATUS,'-') FROM INV_T_MRSHEADER WHERE MRSNUMBER = " & iMRNumber & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				if trim(dcrs(0)) = "-" then
					' Status Code - 040105 (Partially Approved)
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040105' WHERE " &_
						"MRSNUMBER = " & iMRNumber & ""
					Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql
				end if
			end if
			dcrs.close

		end if

		if sActivity = "Requisition" and sAction = "Partial Rejected" then
			' Status Code - 040106 (Partially Rejected)
			sSql = "UPDATE INV_T_MRSHEADER SET MRSHEADERSTATUS = '040106' WHERE " &_
				"MRSNUMBER = " & iMRNumber & ""
			'Response.Write sSql & vbCrLf & vbCrLf
			con.Execute sSql
		end if

		if sActivity = "Issue" and sAction = "Mark" then
			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(MRSPICKSTATUS,'-') FROM INV_T_MRSHEADER WHERE MRSNUMBER = " & iMRNumber & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				if trim(dcrs(0)) = "-" then
					' Status Code - 040119 (Mark for Issue)
					sSql = "UPDATE INV_T_MRSHEADER SET MRSPICKSTATUS = '040119' WHERE " &_
						"MRSNUMBER = " & iMRNumber & ""
					Response.Write sSql & vbCrLf & vbCrLf
					con.Execute sSql
				end if
			end if
			dcrs.close
		end if
	IF iMRNumber  <> "" then 
		if sActivity = "Issue" and sAction = "Create" then
			
			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT SUM(QUANTITYAPPROVED),ISNULL(SUM(QUANTITYISSUED),0),ISNULL(SUM(QUANTITYFORTRANSFER),0),ISNULL(SUM(QUANTITYTOPURCHASE),0),ISNULL(SUM(QUANTITYPURCHASED),0) FROM INV_T_MRSITEMDETAILS WHERE MRSNUMBER = " & iMRNumber & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if not dcrs.EOF then
				iQuantity = cdbl(dcrs(0))
				iISSQuantity = cdbl(dcrs(1))
				iSTQuantity = cdbl(dcrs(2))
				iPRQuantity = cdbl(dcrs(3))
				iPurdQuantity = cdbl(dcrs(4))

				'if (iISSQuantity + iSTQuantity + iPRQuantity) = iQuantity then
				if (iISSQuantity) = iQuantity then
					'Status Code - 040111 (Issued)
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040111' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity > 0 and iPRQuantity = 0 and iSTQuantity = 0) then
					'Partially Issued
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040112' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity > 0 and iPRQuantity > 0 and iSTQuantity = 0 and iPurdQuantity = 0) then
					'Partially Issued / PR
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040113' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity > 0 and iPRQuantity > 0 and iSTQuantity = 0 and iPurdQuantity > 0 and iPRQuantity > iPurdQuantity) then
					'Partially Issued / Partially Purchased
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040123' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity > 0 and iPRQuantity > 0 and iSTQuantity = 0 and iPurdQuantity > 0 and iPRQuantity = iPurdQuantity) then
					'Partially Issued / Purchased
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040122' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity > 0 and iPRQuantity = 0 and iSTQuantity > 0) then
					'Partially Issued / ST
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040114' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity = 0 and iPRQuantity > 0 and iSTQuantity > 0 and iPurdQuantity = 0) then
					'Partially PR / ST
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040115' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity = 0 and iPRQuantity > 0 and iSTQuantity > 0 and iPurdQuantity > 0 and iPRQuantity > iPurdQuantity) then
					'Partially Purchased / ST
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040126' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity = 0 and iPRQuantity > 0 and iSTQuantity > 0 and iPurdQuantity > 0 and iPRQuantity = iPurdQuantity) then
					'Partially ST / Purchased
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040127' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity > 0 and iSTQuantity > 0 and iPRQuantity > 0 and iPurdQuantity = 0) then
					'Partially Issued / ST / PR
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040116' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity > 0 and iSTQuantity > 0 and iPRQuantity > 0 and iPurdQuantity > 0 and iPRQuantity > iPurdQuantity) then
					'Partially Issued / ST / Partially Purchased
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040128' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity > 0 and iSTQuantity > 0 and iPRQuantity > 0 and iPurdQuantity > 0 and iPRQuantity = iPurdQuantity) then
					'Partially Issued / ST / Purchased
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040129' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity = 0 and iSTQuantity = 0 and iPRQuantity > 0 and iPurdQuantity = 0) then
					'Purchase Request
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040117' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity = 0 and iSTQuantity = 0 and iPRQuantity > 0 and iPurdQuantity > 0 and iPRQuantity > iPurdQuantity) then
					'Partially Purchased
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040125' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity = 0 and iSTQuantity = 0 and iPRQuantity > 0 and iPurdQuantity > 0 and iPRQuantity = iPurdQuantity) then
					'Purchased
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040124' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity = 0 and iSTQuantity > 0 and iPRQuantity = 0) then
					'Request For Transfer
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040118' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				elseif (iISSQuantity = 0 and iSTQuantity = 0 and iPRQuantity = 0) then
					'Partially Issued
					sSql = "UPDATE INV_T_MRSHEADER SET MRSISSUESTATUS = '040112' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				end if
				Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql
			end if
			dcrs.close
		end if
	end if 'IF iMRNumber  <> "" then 
		if sActivity = "Pick" and sAction = "Create" then
			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ISNULL(SUM(QUANTITYFORISSUE),0),ISNULL(SUM(QUANTITYISSUED),0) FROM INV_T_MRSISSUEPICK WHERE MRSNUMBER = " & iMRNumber & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing

			if not dcrs.EOF and cdbl(dcrs(0)) > 0 then
				iQuantity = cdbl(dcrs(0))
				iISSQuantity = cdbl(dcrs(1))

				if (iISSQuantity) = iQuantity then
					'Status Code - 040121 (Pick Complete)
					sSql = "UPDATE INV_T_MRSHEADER SET MRSPICKSTATUS = '040121' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				else
					'Pick Partial
					sSql = "UPDATE INV_T_MRSHEADER SET MRSPICKSTATUS = '040120' WHERE " &_
							"MRSNUMBER = " & iMRNumber & ""
				end if
				Response.Write sSql & vbCrLf & vbCrLf
				con.Execute sSql
			end if
			dcrs.close

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT ORGANISATIONCODE,CLASSIFICATIONCODE,ITEMCODE  FROM INV_T_MRSISSUEPICK WHERE MRSNUMBER = " & iMRNumber & ""
				.ActiveConnection = con
				.Open
			end with
			set dcrs.ActiveConnection = nothing
			iCtr = 0 
			do while not dcrs.EOF
				iCtr = iCtr + 1
				sUnitCode = trim(dcrs(0))
				iClassCode = trim(dcrs(1))
				iItemCode = trim(dcrs(2))
				MRLineStatusUpdate "Issue","Create",iMRNumber,iItemCode,iClassCode,iEntNo,sUnitCode,"4","F","",iCtr

			dcrs.movenext
			loop
			dcrs.close

			MRStatusUpdate "Issue","Create",iMRNumber,iItemCode,iClassCode,iEntNo,sUnitCode

		end if

	End Function
%>

<%
'New Function to Update INV_T_LOCATIONLOT added by Ragavendran on April 01,2010
Function UpdateLocLot(iEntNo,iItemCode,iClass,sOrgID,sLoc,sBin,iIssQty,iCreatedBy,dCreatedOn,iSerialNo,iLotNo)
	dim dcrs,dcrs1,sSql,iInvRecNo,iQtyIss,iLotNetQty,iCtr,iTempQty,iChkIssQty,sAccountType
	
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	sSql = "Select IsNull(AccountingType,'W')  from Inv_M_ItemMaster where ItemCode in ("& iItemCode &")"
	dcrs.open sSql,con
	if not dcrs.eof then
	    sAccountType = dcrs(0)
	end if
	dcrs.close
	
	if Trim(sAccountType)="L" then
	        
	        if trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
	        
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
				        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID) &" and SerialNumber = " & iSerialNo &" and LotNumber = "& Pack(iLotNo) &""&_
				        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO desc"
	        elseif trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and ( trim(iLotNo)="0" or trim(iLotNo)="") then
        	
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
				        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID) &" and SerialNumber = " & iSerialNo  &""&_
				        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO  desc"
	        elseif (trim(iSerialNo) = "0" or trim(iSerialNo)="") and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
        	
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
				        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID)  &" and LotNumber = "& Pack(iLotNo)  &""&_
				        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO desc"
        				
	        else
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
					        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID) &" " &_
					        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO desc"
	        end if
	elseif Trim(sAccountType)="F" or Trim(sAccountType)="W" then
	
	        if trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
	        
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
				        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID) &" and SerialNumber = " & iSerialNo &" and LotNumber = "& Pack(iLotNo) &""&_
				        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO"
	        elseif trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and ( trim(iLotNo)="0" or trim(iLotNo)="") then
        	
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
				        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID) &" and SerialNumber = " & iSerialNo  &""&_
				        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO"
	        elseif (trim(iSerialNo) = "0" or trim(iSerialNo)="") and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
        	
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
				        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID)  &" and LotNumber = "& Pack(iLotNo)  &""&_
				        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO"
        				
	        else
			        sSql = " Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,INVENTORYRECEIPTNO from INV_T_LOCATIONLOT "&_
					        " Where itemcode = "& iItemCode &" and ClassificationCode = "& iClass & " and OrganisationCode = " & pack(sOrgID) &" " &_
					        " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by INVENTORYRECEIPTNO"
	        end if
	end if ' if Trim(sAccountType)="L" then
	
	
	Response.Write "<p> First = "&   sSql
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	end with
	'Response.Write "INV_T_LOCATIONLOT="& sSql & vbcrlf & vbcrlf 
	set dcrs.ActiveConnection = nothing
	iCtr = 1
	iChkIssQty = iIssQty 
	IF NOT dcrs.EOF THEN
		while not dcrs.EOF
		    
	'IF not  dcrs.EOF then
				iQtyIss = dcrs(0)
				Response.Write "iQtyIss = "& iQtyIss
				'iInvRecNo = dcrs(1)
				IF cdbl(iQtyIss) > 0 then	
				    if Trim(sAccountType)="L" then
				        if trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode &" and SerialNumber =  "& iSerialNo &" and LotNumber = "& Pack(iLotNo)  &""&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo desc"
							
						elseif trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and (trim(iLotNo)="0" or trim(iLotNo)="") then
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode &" and SerialNumber =  "& iSerialNo &" "&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo desc"
							
						elseif (trim(iSerialNo) = "0" or trim(iSerialNo)="") and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode  &" and LotNumber = "& Pack(iLotNo)  &""&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo desc"
							
						else
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode &" "&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo desc"
						end if
				    elseif Trim(sAccountType)="F" or Trim(sAccountType)="W" then	
						if trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode &" and SerialNumber =  "& iSerialNo &" and LotNumber = "& Pack(iLotNo)  &""&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo"
							
						elseif trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and (trim(iLotNo)="0" or trim(iLotNo)="") then
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode &" and SerialNumber =  "& iSerialNo &" "&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo"
							
						elseif (trim(iSerialNo) = "0" or trim(iSerialNo)="") and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode  &" and LotNumber = "& Pack(iLotNo)  &""&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo"
							
						else
							sSql = "Select (isnull(LOTQUANTITYNETT,0) - isnull(QUANTITYISSUED,0)) as QUANTITYISSUED,InventoryReceiptNo from INV_T_LOCATIONLOT where itemcode  = "& iItemCode &" "&_
							" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) Order by InventoryReceiptNo"
						end if
					end if
						Response.Write "<p> ->SeconD = "&  sSql
								with dcrs1
									.CursorLocation = 3
									.CursorType = 3
									.Source = sSql
									.ActiveConnection = con
									.Open
								end with		
								if not dcrs1.EOF then
									while not dcrs1.EOF 
										iLotNetQty = dcrs1(0)
										iInvRecNo  = dcrs1(1)
										iCtr = iCtr + 1
									'	 Response.Write "<BR><BR>iChkIssQty="&iChkIssQty&"<BR><BR>"
										IF cdbl(iChkIssQty) <> 0 then 
									'	Response.Write "<P>"& iLotNetQty &"   <   "& iChkIssQty &"<BR><BR>"
											
											If cdbl(iLotNetQty) < cdbl(iChkIssQty) then 	
												iTempQty = iIssQty
												
												'iIssQty = iLotNetQty
												if trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and trim(iLotNo)<>"0" and trim(iLotNo)<>"" then
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED +  "& iLotNetQty &"),RESERVED = (ISNULL(RESERVED,0) + " & iLotNetQty & "),"&_
															"CREATEDBY = "&iCreatedBy &",CREATEDON =  convert(datetime,'"&dCreatedOn&"',103) WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and SerialNumber = "& iSerialNo & " and LotNumber ="& Pack(iLotNo) &" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null)"
												elseif trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and (trim(iLotNo)="0" or trim(iLotNo)="" ) then
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED +  "& iLotNetQty &"),RESERVED = (ISNULL(RESERVED,0) + " & iLotNetQty & "),"&_
															"CREATEDBY = "&iCreatedBy &",CREATEDON =  convert(datetime,'"&dCreatedOn&"',103) WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and SerialNumber = "& iSerialNo &" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null)"
												elseif (trim(iSerialNo) = "0" or trim(iSerialNo)="") and trim(iLotNo)<>"" and trim(iLotNo)<>"0" then
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED +  "& iLotNetQty &"),RESERVED = (ISNULL(RESERVED,0) + " & iLotNetQty & "),"&_
															"CREATEDBY = "&iCreatedBy &",CREATEDON =  convert(datetime,'"&dCreatedOn&"',103) WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and LotNumber ="& Pack(iLotNo) &" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null)"
												else
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED +  "& iLotNetQty &"),RESERVED = (ISNULL(RESERVED,0) + " & iLotNetQty & "),"&_
															"CREATEDBY = "&iCreatedBy &",CREATEDON =  convert(datetime,'"&dCreatedOn&"',103) WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) "
												end if

												Response.Write "<p> lot= "& sSql & vbcrlf & vbcrlf 
												Con.Execute sSql			
												iChkIssQty = cdbl(iChkIssQty) - cdbl(iLotNetQty)
											Else
												
												if trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and trim(iLotNo)<>"" and trim(iLotNo)<>"0" then
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED +  "& iChkIssQty &"),RESERVED = (ISNULL(RESERVED,0) + " & iChkIssQty & "),"&_
															"CREATEDBY = "&iCreatedBy &",CREATEDON =  convert(datetime,'"&dCreatedOn&"',103) WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and SerialNumber ="& iSerialNo & " and LotNumber = "& Pack(iLotNo) &" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null)"
															
												elseif trim(iSerialNo) <> "0" and trim(iSerialNo)<>"" and (trim(iLotNo)="" or trim(iLotNo)="0") then
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED +  "& iChkIssQty &"),RESERVED = (ISNULL(RESERVED,0) + " & iChkIssQty & "),"&_
															"CREATEDBY = "&iCreatedBy &",CREATEDON =  convert(datetime,'"&dCreatedOn&"',103) WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and SerialNumber ="& iSerialNo &" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null)"
															
												elseif (trim(iSerialNo) = "0" or trim(iSerialNo)="") and trim(iLotNo)<>"" and trim(iLotNo)<>"0" then
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED +  "& iChkIssQty &"),RESERVED = (ISNULL(RESERVED,0) + " & iChkIssQty & "),"&_
															"CREATEDBY = "&iCreatedBy &",CREATEDON =  convert(datetime,'"&dCreatedOn&"',103) WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and LotNumber ="& Pack(iLotNo)  &" and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null)"
												else
													sSql =  "UPDATE INV_T_LOCATIONLOT SET QUANTITYISSUED = (QUANTITYISSUED +  "& iChkIssQty &"),RESERVED = (ISNULL(RESERVED,0) + " & iChkIssQty & "),"&_
															"CREATEDBY = "&iCreatedBy &",CREATEDON =  convert(datetime,'"&dCreatedOn&"',103) WHERE  "&_
															"ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClass & " AND  ORGANISATIONCODE =" & Pack(sOrgID) & " AND "&_
															" INVENTORYRECEIPTNO = "& iInvRecNo & " and StorageLocationNo = "& sLoc &" and (StorageBinNUmber = "& sBin &" or StorageBinNumber is Null) "
												end if

												Response.Write "<p> lot1= "& sSql & vbcrlf & vbcrlf 
												Con.Execute sSql
												iChkIssQty = 0 'cint(iChkIssQty) - cint(iLotNetQty)
												
											End IF 'If cint(iLotNetQty) < cint(iIssQty) then 	
															
											iTempQty = cdbl(iTempQty) - cdbl(iIssQty)
										 End If
											'Response.Write "<BR><p>"&iChkIssQty&"   =   " &iTempQty&"<BR><BR>"
										dcrs1.MoveNext 				
									wend
								end if
								dcrs1.Close
					End IF
	
		
			dcrs.movenext
		wend
	End IF ' IF NOT DCRS.EOF THEN
	dcrs.Close
	
End Function
%>