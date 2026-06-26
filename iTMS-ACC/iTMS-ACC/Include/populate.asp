<%
	Dim iSAApplicationPop,iSAProcessPop,iSAActivityPop,iEmpNoPopulate
	iSAApplicationPop = Session("iApplication")
	iSAProcessPop = Session("iProcess")
	iSAActivityPop = Session("iActivity")
	iEmpNoPopulate = Session("employeenumber")
%>

<%
	' Function to populate the Organization list
	Function populateOrganization()
		' Declaration of variables
		Dim dcrs,sOrgID,sOrgName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		With dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ORGANIZATIONID,ORGANIZATIONNAME FROM DCS_ORGANIZATION ORDER BY ORGANIZATIONID"
			.ActiveConnection = con
			.Open
		End With
		Set dcrs.ActiveConnection = Nothing
		Set sOrgID = dcrs(0)
		Set sOrgName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(sOrgID)&""">"&trim(sOrgName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		End If
		dcrs.Close
	End Function
%>
<%
	' Function to populate the Employee list with Value
	Function popEmployee()
		' Declaration of variables
		Dim dcrs,sempNumber,sempName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT EMPLOYEENUMBER,isNull(USERNAME,''),isNull(MIDDLENAME,''),isNull(LASTNAME,'') FROM VW_OrgEmployee_List ORDER BY USERNAME"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		Do While Not dcrs.EOF
			sempNumber = dcrs(0)
			if trim(dcrs(2)) <> "" then
				sempName = dcrs(1)&" "&dcrs(2)&" "&dcrs(3)
			else
				sempName = dcrs(1)&" "&dcrs(3)
			end if

			Response.Write("<OPTION VALUE="""&trim(sempNumber)&""" >"&trim(sempName)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close
	End Function
%>


<%
	' Function to populate the Employee list with Value
	Function populateEmployeeWithVal(sVal)
		' Declaration of variables
		Dim dcrs,sempNumber,sempName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT EMPLOYEENUMBER,EMPLOYEENAME FROM MS_EMPLOYEEMASTER ORDER BY EMPLOYEENUMBER"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sempNumber = dcrs(0)
		set sempName = dcrs(1)
		Do While Not dcrs.EOF
			IF CStr(sVal) = CStr(dcrs(0)) Then
				Response.Write("<OPTION VALUE="""&trim(sempNumber)&""" Selected>"&trim(sempName)&"</OPTION>" &vbcrlf)
			Else
				Response.Write("<OPTION VALUE="""&trim(sempNumber)&""">"&trim(sempName)&"</OPTION>" &vbcrlf)
			End IF
			dcrs.MoveNext
		Loop
		dcrs.Close
	End Function
%>

<%
	' Function to populate the Organization Units list
	Function populateOrganizationUnit()
		' Declaration of variables
		Dim dcrs,sUnitID,sUnitName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		With dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ORGANIZATIONUNITID,ORGANIZATIONUNITNAME FROM DCS_ORGANIZATIONUNITS ORDER BY ORGANIZATIONUNITID"
			.ActiveConnection = con
			.Open
		End With

		Set dcrs.ActiveConnection = Nothing

		Set sUnitID = dcrs(0)
		Set sUnitName = dcrs(1)

		If Not dcrs.EOF Then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(sUnitID)&""">"&trim(sUnitName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		End If
		dcrs.Close
	End Function
%>
<%
	' Function to populate the Organization Units Definition list
	Function populateOrganizationUnitDefinition()
		' Declaration of variables
		Dim dcrs,sUnitID,sUnitName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION,ORGANIZATIONUNITID FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE ORGANIZATIONUNITID = (SELECT MIN(ORGANIZATIONUNITID) FROM DCS_ORGANIZATIONUNITS) ORDER BY ORGANIZATIONUNITID"
			.ActiveConnection = con
			.Open
		end with

		set dcrs.ActiveConnection = nothing

		set sUnitID = dcrs(0)
		set sUnitName = dcrs(1)

		If not dcrs.EOF then

			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(sUnitID)&""">"&trim(sUnitName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop

		end if
				dcrs.Close
	End Function
%>

<%
	' Function to populate the Units list
	Function populateUnit()
	' Declaration of variables
	Dim dcrs,sUnitID,sUnitName,sQuery
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

'	If iSAApplicationPop <> "" then
'		sQuery = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID IN (SELECT DISTINCT ORGANISATIONCODE FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & iEmpNoPopulate & " AND APPLICATIONCODE = " & iSAApplicationPop & " AND PROCESSCODE = " & iSAProcessPop & " AND ACTIVITYCODE = " & iSAActivityPop & ") ORDER BY OUDEFINITIONID"
'	Else
		sQuery = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID"
'	End If

	With dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	End With
	Set dcrs.ActiveConnection = Nothing
	Set sUnitID = dcrs(0)
	Set sUnitName = dcrs(1)

	If Not dcrs.EOF Then
		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUnitID)&""">"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
	End If
	dcrs.Close
End Function
%>

<%
	' Function to populate the Units list with default selected
	Function populateUnitSelected(sOrg)
	' Declaration of variables
	Dim dcrs,sUnitID,sUnitName
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")

	With dcrs
		.CursorLocation = 3
		.CursorType = 3
		If iSAApplicationPop <> "" Then
			.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE OUDEFINITIONID IN (SELECT DISTINCT ORGANISATIONCODE FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & iEmpNoPopulate & " AND APPLICATIONCODE = " & iSAApplicationPop & " AND PROCESSCODE = " & iSAProcessPop & " AND ACTIVITYCODE = " & iSAActivityPop & ") ORDER BY OUDEFINITIONID"
		Else
			.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID"
		End If
			.ActiveConnection = con
			.Open
	End With

	Set dcrs.ActiveConnection = Nothing

	Set sUnitID = dcrs(0)
	Set sUnitName = dcrs(1)

	If Not dcrs.EOF Then
		Do While Not dcrs.EOF

			If sUnitID = trim(sOrg) Then
				Response.Write("<OPTION VALUE="""&trim(sUnitID)&""" SELECTED>"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			Else
				Response.Write("<OPTION VALUE="""&trim(sUnitID)&""">"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			End If
				dcrs.MoveNext
		Loop
	End If
	dcrs.Close
End Function
%>
<%
	' Function to populate the salary Group
	Function popSalaryGroup(SalType)
		' Declaration of variables
		Dim dcrs,sSalType,sSalGrpName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		With dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT InvoiceType,InvoiceTypeName FROM PAY_M_PayRollGroup where useable=1 ORDER BY InvoiceTypeName"
			.ActiveConnection = con
			.Open
		End With
		Set dcrs.ActiveConnection = Nothing
		Set sSalType = dcrs(0)
		Set sSalGrpName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				if trim(SalType) = trim(sSalType) then
					Response.Write("<OPTION VALUE="""&trim(sSalType)&""" selected>"&trim(sSalGrpName)&"</OPTION>" &vbcrlf)
				else
					Response.Write("<OPTION VALUE="""&trim(sSalType)&""">"&trim(sSalGrpName)&"</OPTION>" &vbcrlf)
				end if
				dcrs.MoveNext
			Loop
		End If
		dcrs.Close
	End Function
%>

<%
	' Function to populate the Organization list which has been defined
	Function populateOrgList()
		' Declaration of variables
		Dim dcrs,sUnitLID,sUnitLName,sUnitSName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		With dcrs
			.CursorLocation = 3
			.CursorType = 3
			If iSAApplicationPop <> "" Then
			.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION,ORGANIZATIONUNITID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS D1 WHERE ORGANIZATIONUNITID = (SELECT MAX(ORGANIZATIONUNITID) FROM DCS_ORGANIZATIONUNITS) AND OUDEFINITIONID = (SELECT DISTINCT OUDEFINITIONID FROM VWORGANIZATIONDEFINED WHERE OUDEFINITIONID = D1.OUDEFINITIONID) AND OUDEFINITIONID IN (SELECT DISTINCT ORGANISATIONCODE FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & iEmpNoPopulate & " AND APPLICATIONCODE = " & iSAApplicationPop & " AND PROCESSCODE = " & iSAProcessPop & " AND ACTIVITYCODE = " & iSAActivityPop & ") ORDER BY ORGANIZATIONUNITID"
			Else
			.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION,ORGANIZATIONUNITID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS D1 WHERE ORGANIZATIONUNITID = (SELECT MAX(ORGANIZATIONUNITID) FROM DCS_ORGANIZATIONUNITS) AND OUDEFINITIONID = (SELECT DISTINCT OUDEFINITIONID FROM VWORGANIZATIONDEFINED WHERE OUDEFINITIONID = D1.OUDEFINITIONID) ORDER BY ORGANIZATIONUNITID"
			End If
			.ActiveConnection = con
			.Open
		End With

		Set dcrs.ActiveConnection = Nothing

		Set sUnitLID = dcrs(0)
		Set sUnitLName = dcrs(1)
		Set sUnitSName = dcrs(3)

		If Not dcrs.EOF Then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(sUnitLID)&""">"&trim(sUnitSName)&"</OPTION>")
				dcrs.MoveNext
			Loop
		End If
		dcrs.Close
	End Function
%>
<%
	' Function to populate the Item Type list
	Function populateItemType()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
		'	if iSAApplicationPop <> "" then
		'	.Source = "SELECT DISTINCT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE WHERE ITEMTYPEID IN (SELECT DISTINCT ITEMTYPEID FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & iEmpNoPopulate & " AND APPLICATIONCODE = " & iSAApplicationPop & " AND PROCESSCODE = " & iSAProcessPop & " AND ACTIVITYCODE = " & iSAActivityPop & ") ORDER BY ITEMTYPENO"
		'	else
			.Source = "SELECT DISTINCT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE ORDER BY ITEMTYPENO"
		'	end if
			.ActiveConnection = con
			.Open
		end with

		set dcrs.ActiveConnection = nothing
		set stypID = dcrs(0)
		set stypName = dcrs(1)

		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		End If
		dcrs.Close
	End Function
%>
<%
	' Function to populate the Item Type list with default selected
	Function populateItemTypeSelected(sItemType)
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
		'	If iSAApplicationPop <> "" Then
		'		.Source = "SELECT DISTINCT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE WHERE ITEMTYPEID IN (SELECT DISTINCT ITEMTYPEID FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & iEmpNoPopulate & " AND APPLICATIONCODE = " & iSAApplicationPop & " AND PROCESSCODE = " & iSAProcessPop & " AND ACTIVITYCODE = " & iSAActivityPop & ") ORDER BY ITEMTYPENO"
		'	Else
				.Source = "SELECT DISTINCT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE ORDER BY ITEMTYPENO"
		'	End If
				.ActiveConnection = con
				.Open
		End With
		Set dcrs.ActiveConnection = Nothing
		Set stypID = dcrs(0)
		Set stypName = dcrs(1)
'Response.Write "<option value=>"&dcrs.source &"</option>"
		If Not dcrs.EOF Then
		    
			Do While Not dcrs.EOF
				if sItemType = trim(stypID) then
					Response.Write("<OPTION VALUE="""&trim(stypID)&""" SELECTED>"&trim(stypName)&"</OPTION>" &vbcrlf)
				else
					Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				end if
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
	End Function
%>

<%
	' Function to populate the Item Type list without Plant and Capitial
	Function populateItemTypeWOCap()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3

			if iSAApplicationPop <> "" then
				.Source = "SELECT DISTINCT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE WHERE ITEMTYPEID <> 'PLA' AND ITEMTYPEID IN (SELECT DISTINCT ITEMTYPEID FROM MS_USERACTIVITY WHERE INTERNALUSERID = " & iEmpNoPopulate & " AND APPLICATIONCODE = " & iSAApplicationPop & " AND PROCESSCODE = " & iSAProcessPop & " AND ACTIVITYCODE = " & iSAActivityPop & ") ORDER BY ITEMTYPENO"
			else
				.Source = "SELECT DISTINCT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE WHERE ITEMTYPEID <> 'PLA' ORDER BY ITEMTYPENO"
			end if
				.ActiveConnection = con
				.Open
		end with
		set dcrs.ActiveConnection = nothing
		set stypID = dcrs(0)
		set stypName = dcrs(1)

		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
	End Function
%>

<%
	' Function to populate the Category list with default selected
	Function populateCategorySelected(sCat)
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT CATEGORYCODE,CATEGORYNAME FROM INV_M_CLASSIFICATIONCATEGORY ORDER BY 1"
			.ActiveConnection = con
			.Open
		End With
		Set dcrs.ActiveConnection = Nothing
		Set stypID = dcrs(0)
		Set stypName = dcrs(1)

		If Not dcrs.EOF Then
			Do While Not dcrs.EOF
				if sCat = trim(stypID) then
					Response.Write("<OPTION VALUE="""&trim(stypID)&""" SELECTED>"&trim(stypName)&"</OPTION>" &vbcrlf)
				else
					Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				end if
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
	End Function
%>

<%
	' Function to populate the Storage Location list
	Function populateStorageLocation()
	' Declaration of variables
	Dim dcrs,sUnitID,sUnitName
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID"
		.ActiveConnection = con
		.Open
	end with

	set dcrs.ActiveConnection = nothing
	set sUnitID = dcrs(0)
	set sUnitName = dcrs(1)

	If not dcrs.EOF then
		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUnitID)&""">"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
	end if
	dcrs.Close
End Function
%>

<%
	' Function to populate the Organization list
	Function populateOrganizationListDB()
		' Declaration of variables

		Dim dcrs,sUnitLID,sUnitLName,sUnitSName,sQy
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		'SELECT OUDEFINITIONID,ORGUNITDESCRIPTION,ORGANIZATIONUNITID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE ORGANIZATIONUNITID = (SELECT MAX(ORGANIZATIONUNITID) FROM DCS_ORGANIZATIONUNITS) and OUDEFINITIONID IN (Select Distinct OrganisationCode From Ms_UserActivity  Where InternalUserId = "&getUserID&") ORDER BY ORGANIZATIONUNITID"

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT Distinct M.OrganisationCode,D.ORGUNITDESCRIPTION,D.ORGANIZATIONUNITID,D.ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS D, Ms_UserActivity M Where D.OUDEFINITIONID = M.OrganisationCode and M.InternalUserId = "&getUserID&" ORDER BY M.OrganisationCode "
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUnitLID = dcrs(0)
		set sUnitLName = dcrs(1)
		set sUnitSName = dcrs(3)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				'Response.Write("<OPTION VALUE="""&trim(sUnitLID)&""">"&trim(sUnitLName)&"</OPTION>")
				Response.Write("<OPTION VALUE="""&trim(sUnitLID)&""">"&trim(sUnitSName)&"</OPTION>")
				dcrs.MoveNext
			Loop
		end if

		dcrs.Close
	End Function
%>

<%
	' Function to populate the Organization list
	Function populateOrganizationListWithFName()
		' Declaration of variables
		Dim dcrs,sUnitLID,sUnitLName,sUnitSName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION,ORGANIZATIONUNITID,OrgUnitDescription FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE ORGANIZATIONUNITID = (SELECT MAX(ORGANIZATIONUNITID) FROM DCS_ORGANIZATIONUNITS) ORDER BY ORGANIZATIONUNITID"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUnitLID = dcrs(0)
		set sUnitLName = dcrs(1)
		set sUnitSName = dcrs(3)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(sUnitLID)&""">"&trim(sUnitSName)&"</OPTION>")
				dcrs.MoveNext
			Loop
		end if

		dcrs.Close
	End Function
%>

<%
	' Function to populate the Organization list
	Function populateOrganizationList()
		' Declaration of variables
		Dim oDom,fs,Root,PGNode
		dim sUnitID,sUnitName,sUnitShName

		Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
		Set fs = CreateObject("Scripting.FileSystemObject")
		if fs.FileExists(Server.MapPath("../../admin/xmldata/Unit.xml")) then
			oDOM.Load server.MapPath("../../admin/xmldata/Unit.xml")
			Set Root = oDOM.documentElement
			if Root.HaschildNodes() then
				For Each PGNode In Root.childNodes
					sUnitID = trim(PGNode.Attributes.Item(0).nodeValue)
					sUnitName = trim(PGNode.Attributes.Item(2).nodeValue)
					sUnitShName = trim(PGNode.Attributes.Item(3).nodeValue)
					Response.Write("<OPTION VALUE="""&trim(sUnitID)&""">"&trim(sUnitName)&"</OPTION>" &vbcrlf)
				next
			end if
		else
			populateOrganizationListDB
		end if
	End Function
%>

<%
	' Function to populate the UoM list
	Function populateUoM()
		' Declaration of variables
		Dim oDom,fs,Root,PGNode,rsObj
		dim sUoMID,sUoMName,sUoMShName,sQuery

		Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
		Set fs = CreateObject("Scripting.FileSystemObject")
		set rsobj = Server.CreateObject("ADODB.Recordset")
		
		    sQuery = "Select UoMCode,UoMDescription,UoMShortDescription from Ms_UnitOfMeasurement"
		    rsObj.Open sQuery,con
		    if not rsObj.EOF then
		        do while not rsObj.EOF 
			    	    sUoMID = trim(rsObj(0))
					    sUoMName =trim(rsObj(1))
					    sUoMShName = trim(rsObj(2))
				        Response.Write("<OPTION VALUE="""&trim(sUoMID)&""">"&trim(sUoMShName)&"</OPTION>" &vbcrlf)
				    rsObj.MoveNext 
				loop
		    end if
		    rsObj.Close 
	End Function
%>
<%	' Function to populate the UoM Selected list

	Function populateUoMSelected(sUomCode)
		' Declaration of variables
		Dim oDom,fs,Root,PGNode
		dim sUoMID,sUoMName,sUoMShName

		Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
		Set fs = CreateObject("Scripting.FileSystemObject")
		if fs.FileExists(Server.MapPath("../xmldata/UoM.xml")) then
			oDOM.Load server.MapPath("../xmldata/UoM.xml")
			Set Root = oDOM.documentElement
			if Root.HaschildNodes() then
				For Each PGNode In Root.childNodes
					sUoMID = trim(PGNode.Attributes.Item(0).nodeValue)
					sUoMName = trim(PGNode.Attributes.Item(1).nodeValue)
					sUoMShName = trim(PGNode.Attributes.Item(2).nodeValue)
					IF trim(sUoMID) = trim(sUomCode)  then
						Response.Write("<OPTION VALUE="""&trim(sUoMID)&""" Selected>"&trim(sUoMShName)&"</OPTION>" &vbcrlf)
					Else
						Response.Write("<OPTION VALUE="""&trim(sUoMID)&""">"&trim(sUoMShName)&"</OPTION>" &vbcrlf)
					End IF
				next
			end if
		end if
	End Function
%>
<%
	' Function to populate the Employee list
	Function populateEmployee()
		' Declaration of variables
		Dim dcrs,sempNumber,sempName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT EMPLOYEENUMBER,EMPLOYEENAME FROM MS_EMPLOYEEMASTER ORDER BY EMPLOYEENUMBER"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sempNumber = dcrs(0)
		set sempName = dcrs(1)
		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sempNumber)&""">"&trim(sempName)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close
	End Function
%>

<%
	' Function to populate the Employee list with default selected
	Function populateEmployeeSelected(iUser)
		' Declaration of variables
		Dim dcrs,sempNumber,sempName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT EMPLOYEENUMBER,EMPLOYEENAME FROM MS_EMPLOYEEMASTER WHERE EMPLOYEENUMBER = " & iUser & " ORDER BY EMPLOYEENUMBER"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sempNumber = dcrs(0)
		set sempName = dcrs(1)
		Do While Not dcrs.EOF
			if cint(sempNumber) = cint(iUser) then
				Response.Write("<OPTION VALUE="""&trim(sempNumber)&""" SELECTED>"&trim(sempName)&"</OPTION>" &vbcrlf)
			else
				Response.Write("<OPTION VALUE="""&trim(sempNumber)&""">"&trim(sempName)&"</OPTION>" &vbcrlf)
			end if
			dcrs.MoveNext
		Loop
		dcrs.Close
	End Function
%>


<%
	' Function to populate the Country list
	Function populateCountry()
		' Declaration of variables
		Dim dcrs1,sconCode,sconName
		'Declaration of Objects
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT COUNTRYCODE,COUNTRYNAME FROM MS_COUNTRY ORDER BY COUNTRYCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing
		set sconCode = dcrs1(0)
		set sconName = dcrs1(1)

		Do While Not dcrs1.EOF
			Response.Write("<OPTION VALUE="""&trim(sconCode)&""">"&trim(sconName)&"</OPTION>" &vbcrlf)
			dcrs1.MoveNext
		Loop
		dcrs1.Close
	End Function
%>

<%
	' Function to populate the Currency list
	Function populateCurrency()
		' Declaration of variables
		Dim dcrs2,scurCode,scurName
		'Declaration of Objects
		Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
		with dcrs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CURRENCYCODE,CURRENCYNAME FROM MS_CURRENCYMASTER ORDER BY CURRENCYCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs2.ActiveConnection = nothing

		set scurCode = dcrs2(0)
		set scurName = dcrs2(1)

		Do While Not dcrs2.EOF
			Response.Write("<OPTION VALUE="""&trim(scurCode)&""">"&trim(scurName)&"</OPTION>" &vbcrlf)
			dcrs2.MoveNext
		Loop
		dcrs2.Close
	End Function
%>

<%
	' Function to format the date in dd/mm/yyyy
	Function FormatDate(Date1)
		dim dDate2,sDa,sMo,sYe,sDastr
		dDate2 = Date1
		sDa = Day(dDate2)
		sMo = Month(dDate2)
		sYe = Year(dDate2)

		If sDa < 10 then
			sDa = "0"&sDa
		End if

		If sMo < 10 then
			sMo = "0"&sMo
		End if

		sDastr = sDa&"/"&sMo&"/"&sYe
		FormatDate = sDastr
	End Function
%>

<%
	' Function to get the Current Financial Year
	Function GetFinancialYear(sYrMon)
		dim sTempFrom,sTempTo

		if cint(mid(sYrMon,1,2)) < 4 and cint(mid(sYrMon,3)) = cint(Year(date)) then
			sTempFrom = "01/04/"&mid(sYrMon,3)-1
			sTempTo = "31/03/"&mid(sYrMon,3)
			GetFinancialYear = sTempFrom&":"&sTempTo
		else
			sTempFrom = "01/04/"&mid(sYrMon,3)
			sTempTo = "31/03/"&mid(sYrMon,3)+1
			GetFinancialYear = sTempFrom&":"&sTempTo
		end if

	end Function
%>
<%
	' Function to replace single Quote
	Private Function Pack(Value)
	  dim sValue
	  sValue = Value
	  Const SQ = "'" ' single quote
	  Const AM = "&" ' Ambersand
	  IF SQ <> "" then  sValue = trim(SQ & Replace(sValue, SQ, SQ & SQ) & SQ)
	  Pack = ucase(sValue)
	  'Pack = sValue
	End Function
%>
<%
	' Function to populate the Country list
	Function populateShift()
		' Declaration of variables
		Dim dcrs1,sShiftCode,sShiftName,sComName
		'Declaration of Objects
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT SHIFTCODE,SHIFTNAME FROM PRD_M_WORKSHIFT ORDER BY SHIFTCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing

		set sShiftCode = dcrs1(0)
		set sShiftName = dcrs1(1)

		Do While Not dcrs1.EOF
			Response.Write("<OPTION VALUE="""&trim(sShiftCode)&":"&trim(sShiftName)&""">"&trim(sShiftName)&"</OPTION>" &vbcrlf)
			dcrs1.MoveNext
		Loop
		dcrs1.Close
	End Function
%>

<%
Function getUserid()
	getUserid=session("userid")
End Function
%>

<%
 '	Function LastDayOfMonth(DateIn)
 '   	Dim TempDate
 '  	TempDate = Year(dateIn) & "-" & Month(DateIn) & "-"
 ' 	    if IsDate(TempDate & "28") Then LastDayOfMonth = 28
 '   	if IsDate(TempDate & "29") Then LastDayOfMonth = 29
 '   	if IsDate(TempDate & "30") Then LastDayOfMonth = 30
 '   	if IsDate(TempDate & "31") Then LastDayOfMonth = 31
 '  End function
    
    Function LastDayOfMonth(sMonthYear)
    	Dim TempDate,sMonth,sYear
    	sMonth = Left(sMonthYear,2)
    	sYear = Right(sMonthYear,4)
    	
    	if (sMonth = "01" or sMonth = "03" or sMonth = "05" or sMonth = "07" or sMonth = "08" or sMonth = "10" or sMonth = "12") then
    	    LastDayOfMonth = "31"
    	elseif (sMonth = "04" or sMonth = "06" or sMonth = "09" or sMonth = "11") then
    	    LastDayOfMonth = "30"
    	elseif sMonth = "02" then
    	    LastDayOfMonth = "28"
    	    if ((cdbl(sYear)/4)-(CInt(sYear)/4)=0) then
    	       LastDayOfMonth = "29"
    	    end if 
    	end if
    End function
    
%>

<%
	' Function to Get Status Name
	Function GetStatus(sCode)
		' Declaration of variables
		Dim dcrs2,sStatusName
		'Declaration of Objects
		Set dcrs2 = Server.CreateObject("ADODB.RecordSet")
		with dcrs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT STATUSNAME FROM MS_STATUS WHERE STATUSCODE = '"&sCode&"'"
			.ActiveConnection = con
			.Open
		end with
		set dcrs2.ActiveConnection = nothing

		set sStatusName = dcrs2(0)
		if not dcrs2.EOF then
			GetStatus = trim(sStatusName)
		end if
		dcrs2.Close
	End Function
%>

<%
	Function PopWrkGroupCode()
		dim wrkgrs

		set wrkgrs=server.CreateObject("Adodb.Recordset")

		With wrkgrs
			.CursorLocation =3
			.CursorType=3
			.Source ="SELECT WORKGROUPCODE,WORKGROUPNAME FROM PRD_M_WORKGROUP"
			.ActiveConnection =con
			.Open
		End With

		'To Add Work Group Name to the Select Control.
		while not wrkgrs.EOF
			Response.Write "<option Value=""" & trim(wrkgrs(0)) & """>" & wrkgrs(1) & "</option>"&vbcrlf
			wrkgrs.MoveNext
		wend
	End Function
%>

<%
	' Function to populate Store
	Function populateManuStore(sOrgID)
		' Declaration of variables
		Dim dcrs,dcrs1,sLoc,sBin,sBinName,sLocName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT DISTINCT IM.LOCATIONNUMBER,ISNULL(BINNUMBER,0),LOCATIONNAME,LOCATIONCODE FROM INV_M_ORGSTORAGE IC,INV_M_ITEMORGSTORAGE IM WHERE IM.LOCATIONNUMBER = IC.LOCATIONNUMBER AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & " AND IM.APPLICABLEFOR = 'IN' AND IC.APPLICABLEFOR = 'IN' ORDER BY 1,2"
			.Source = "SELECT DISTINCT LOCATIONNUMBER,LOCATIONNAME,LOCATIONCODE FROM INV_M_ORGSTORAGE WHERE OUDEFINITIONID = " & Pack(sOrgID) & " AND APPLICABLEFOR = 'MA' ORDER BY 1,2"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sLoc     = dcrs(0)
		set sLocName = dcrs(1)

		Do While Not dcrs.EOF
			with dcrs1
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT BINNUMBER,BINNAME,BINCODE FROM INV_M_ORGSLBINDETAILS WHERE LOCATIONNUMBER = " & sLoc & ""
				.ActiveConnection = con
				.Open
			end with

			set dcrs1.ActiveConnection = Nothing

			if not dcrs1.EOF then
				Response.Write("<OPTION VALUE="""&trim(sLoc)&":"&trim(dcrs1(0))&""">"&trim(sLocName)&" -- "&trim(dcrs1(2))&"</OPTION>" &vbcrlf)
			else
				Response.Write("<OPTION VALUE="""&trim(sLoc)&""">"&trim(sLocName)&"</OPTION>" &vbcrlf)
			end if
			dcrs1.Close
		dcrs.MoveNext
		Loop
		dcrs.Close
	End Function
%>

<%
	' Function to populate the Organization list With Selected Value
	Function populateOrganizationListWithVal(sTemp)
		' Declaration of variables
		Dim oDom,fs,Root,PGNode
		dim sUnitID,sUnitName,sUnitShName

		Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
		Set fs = CreateObject("Scripting.FileSystemObject")
		if fs.FileExists(Server.MapPath("../../admin/xmldata/Unit.xml")) then
			oDOM.Load server.MapPath("../../admin/xmldata/Unit.xml")
			Set Root = oDOM.documentElement
			if Root.HaschildNodes() then
				For Each PGNode In Root.childNodes
					sUnitID = trim(PGNode.Attributes.Item(0).nodeValue)
					sUnitName = trim(PGNode.Attributes.Item(2).nodeValue)
					sUnitShName = trim(PGNode.Attributes.Item(3).nodeValue)
					If CStr(sTemp) = CStr(sUnitID) Then
						Response.Write("<OPTION VALUE="""&trim(sUnitID)&""" Selected>"&trim(sUnitShName)&"</OPTION>" &vbcrlf)
					Else
						Response.Write("<OPTION VALUE="""&trim(sUnitID)&""">"&trim(sUnitShName)&"</OPTION>" &vbcrlf)
					End IF
				next
			end if
		else
			populateOrganizationListDB
		end if
	End Function
%>

<%
	' Function to populate the Organization list
		Function populateOrganizationListDBWithVal(sVal)
			' Declaration of variables
			Dim dcrs,sUnitLID,sUnitLName,sUnitSName
			'Declaration of Objects
			Set dcrs = Server.CreateObject("ADODB.RecordSet")

			'Old Qry = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION,ORGANIZATIONUNITID,ORGUNITSHORTDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE ORGANIZATIONUNITID = (SELECT MAX(ORGANIZATIONUNITID) FROM DCS_ORGANIZATIONUNITS) ORDER BY ORGANIZATIONUNITID"

			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "Select Distinct OUDEFINITIONID,ORGUNITDESCRIPTION,ORGANIZATIONUNITID,ORGUNITSHORTDESCRIPTION From VwUserUnitList Where ApplicationCode = " & iSAApplicationPop & " and InternalUserID = "&getUserID()&" Order By OUDEFINITIONID "
				.ActiveConnection = con
				.Open
			end with
			Response.write dcrs.Source
			set dcrs.ActiveConnection = nothing
			set sUnitLID = dcrs(0)
			set sUnitLName = dcrs(1)
			set sUnitSName = dcrs(3)
			If not dcrs.EOF then
				Do While Not dcrs.EOF
					IF CStr(sUnitLID) = CStr(sVal) Then
						Response.Write("<OPTION VALUE="""&trim(sUnitLID)&""" Selected>"&trim(sUnitSName)&"</OPTION>")
					Else
						Response.Write("<OPTION VALUE="""&trim(sUnitLID)&""">"&trim(sUnitSName)&"</OPTION>")
					End IF
					dcrs.MoveNext
				Loop
			end if

			dcrs.Close
	End Function
%>
<%
	Function PopulateRecSel()
		Response.Write("<Option Value=""0"" Selected>Select</Option>")
		Response.Write("<Option Value=""K"">Commission Agent - Known Customers</Option>")
		Response.Write("<Option Value=""U"">Commission Agent</Option>")
		Response.Write("<Option Value=""D"">Direct Customers</Option>")
		Response.Write("<Option Value=""P"">Depot Agents</Option>")
	End Function
%>

<%
	'Added by Tajudeen
	' Function to Display the Organization Name
	Function DisplayOrganization()
		' Declaration of variables
		Dim dcrs
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ORGANIZATIONNAME, CITY FROM DCS_ORGANIZATION ORDER BY ORGANIZATIONID"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		If not dcrs.EOF then
			DisplayOrganization = trim(dcrs(0)) & ". " & trim(dcrs(1))
		End If
		dcrs.Close
	End Function
%>

<%
	'Added by Tajudeen
	' Function to Display the UOM
	Function DisplayUOM(iItemCode,iClassCode,sOrgID)
		' Declaration of variables
		Dim dcrs
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT STORESUOM FROM VWITEM WHERE ITEMCODE = " & iItemCode & " AND CLASSIFICATIONCODE = " & iClassCode & "  AND ORGANISATIONCODE = " & Pack(sOrgID)
			.ActiveConnection = con
			.Open
		end with

		set dcrs.ActiveConnection = nothing

		If not dcrs.EOF then
			DisplayUOM = trim(dcrs(0))
		End If
		dcrs.Close
	End Function
%>
<%
	Function GetEmployeeName(sUserID)
		Dim sQuery,dcrs,sEmpName
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		sQuery = "SELECT EmployeeName FROM Ms_EmployeeMaster WHERE EmployeeNumber = "&sUserID&" "

		With dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		End With

		set dcrs.ActiveConnection = nothing

		If Not dcrs.Eof Then
			sEmpName = dcrs(0)
		End If
		dcrs.close
		GetEmployeeName = sEmpName
	End Function
%>

<%
Function PopulateWorkGroup(sWorkGroup)
	Dim Matrs
	Set Matrs=server.CreateObject("Adodb.Recordset")

	With Matrs
		.CursorLocation = 3
		.CursorType =3
		.Source = "SELECT WORKGROUPCODE,WORKGROUPNAME FROM PRD_M_WORKGROUP"
		.ActiveConnection =con
		.Open
	End With
	Set Matrs.Activeconnection = nothing
	'To Add Work Group Name to the Select Control.
	While Not Matrs.EOF
		If Matrs(0)= sWorkGroup Then
			Response.Write "<option selected Value=""" & trim(Matrs(0)) & """>" & Matrs(1) & "</option>"&vbcrlf
		Else
			Response.Write "<option  Value=""" & trim(Matrs(0)) & """>" & Matrs(1) & "</option>"&vbcrlf
		End If
		Matrs.MoveNext
	Wend
		Matrs.close
End function
%>

<%
	' Function to populate the Classification Category
	Function populateCategory()
		Dim dcrs

		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT CATEGORYCODE, CATEGORYNAME FROM INV_M_CLASSIFICATIONCATEGORY ORDER BY CATEGORYNAME"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(dcrs(0))&""">"&trim(dcrs(1))&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close
	End Function
%>

<%
	' Function to Check for Fin. Year
	Function CheckFinYear(dDate)
		' Declaration of variables
		Dim dcrs
		dim sTempMonYr,sMonYr,sFinFrom,sFinTo,arrFin
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		if len(Month(date())) = 1 then
			sTempMonYr = "0"&Month(date())
		else
			sTempMonYr = Month(date())
		end if
		sMonYr = sTempMonYr&Year(date())

		arrFin = split(GetFinancialYear(sMonYr),":")
		sFinFrom = arrFin(0)
		sFinTo = arrFin(1)
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT ORGANISATIONCODE FROM INV_T_ITEMYEARLYSTOCK WHERE CONVERT(DATETIME," & Pack(sFinFrom) & ",103) = CONVERT(DATETIME,FINANCIALYEARFROM,103) AND CONVERT(DATETIME," & Pack(sFinTo) & ",103) = CONVERT(DATETIME,FINANCIALYEARTO,103)"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if Not dcrs.EOF then
			if DateDiff("d",formatdate(sFinFrom),formatdate(dDate)) < 0 then
				CheckFinYear = "1"
			else
				CheckFinYear = "0"
			end if
		else
			if DateDiff("d",formatdate(dDate),formatdate(sFinFrom)) > 0 then
				CheckFinYear = "0"
			else
				CheckFinYear = "2"
			end if
		end if
		dcrs.Close

	End Function
%>

<%'Function to check Log in Finanial Year Added by Maheswari on 6th March 2008

Function CheckLoginFinYear(dDate)
	Dim sFinPeriod,sMaxDate,sMinDate,sFinTemp
	sFinPeriod = Session("FinPeriod")
	IF CStr(sFinPeriod) <> "" Then
		sFinTemp = Split(sFinPeriod,":")
		sMaxDate = "31/03/"&sFinTemp(1)
		sMinDate = "01/04/"&sFinTemp(0)
	End IF
	'Response.Write sMaxDate &"==="& sMinDate
		if DateDiff("d",formatdate(sMinDate),formatdate(dDate)) < 0 then
			CheckLoginFinYear = "0"
		else
			CheckLoginFinYear = "1"
		end if
	'		Response.Write DateDiff("d",formatdate(dDate),formatdate(sMaxDate))
		if DateDiff("d",formatdate(dDate),formatdate(sMaxDate)) > 0 then
			CheckLoginFinYear = "0"
		else
			CheckLoginFinYear = "2"
		end if

End Function


%>
<%'Function to get Attributes Name --  Newly added By Maheswari on March 19th 2008
Function FunAttribName(sAttribList)
Dim sTemp,i,sOptName,iOptVal,dcrs1,sTempAttribList
Set dcrs1= Server.CreateObject("ADODB.RecordSet")
sTempAttribList= split(sAttribList,"#")
If UBound(sTempAttribList)>0 then
	If trim(sTempAttribList(1)) <> "0" and trim(sTempAttribList(1))<>"" then
		sOptName = ""
		iOptVal = ""

		sTemp = split(sTempAttribList(1),",")
		For i = 0 to UBOUND(sTemp)
			iOptVal = sTemp(i)

			if iOptVal <> "" then
				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = "Select OptionName from Inv_M_ItemTypeOptions where OptionValue = "&iOptVal&" "
					.ActiveConnection = con
					.Open
				end with
				If not dcrs1.EOF then
					sOptName = sOptName &","& dcrs1(0)
				End If
				dcrs1.Close
			else
				sOptName = ""
			end if
		Next
	else
	    sOptName = ""
	End If
End if'If UBound(sTempAttribList)>0 then
	IF sOptName <> "" then
		sOptName = " [" & mid(sOptName,2) &"] "
	End IF
	FunAttribName = sOptName
	
End Function
%>
<%
Function popItemTypesNew()
Dim rsTemp,ssql
set rsTemp = Server.CreateObject("ADODB.Recordset")
ssql = "Select ItemTypeID,ItemTypeDescription from INV_M_ItemTypes"
rsTemp.Open ssql,con
if not rsTemp.EOF then
    do while not rsTemp.EOF
        Response.Write "<option value="&Trim(rsTemp(0))&">"&Trim(rsTemp(1))&"</option>"

        rsTemp.MoveNext
    loop
end if
rsTemp.Close
End Function
%>

<%
Function popItemTypes(iTypeID)
Dim rsTemp,ssql
set rsTemp = Server.CreateObject("ADODB.Recordset")
ssql = "Select ItemTypeID,ItemTypeDescription from INV_M_ItemTypes where ItemTypeID ="& iTypeID
rsTemp.Open ssql,con
if not rsTemp.EOF then
    do while not rsTemp.EOF
        Response.Write "<option value="&Trim(rsTemp(0))&" selected>"&Trim(rsTemp(1))&"</option>"

        rsTemp.MoveNext
    loop
end if
rsTemp.Close
End Function
%>