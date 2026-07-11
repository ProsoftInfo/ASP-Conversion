<%
	' Function to populate the UoM list
	Function populateUoM()
		' Declaration of variables
		Dim oDom,fs,Root,PGNode
		dim sUoMID,sUoMName,sUoMShName
			
		Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
		Set fs = CreateObject("Scripting.FileSystemObject")
		if fs.FileExists(Server.MapPath("../../Inventory/xmldata/UoM.xml")) then
			
			oDOM.Load server.MapPath("../../Inventory/xmldata/UoM.xml")
			Set Root = oDOM.documentElement
			if Root.HaschildNodes() then
				For Each PGNode In Root.childNodes
					
					sUoMID = trim(PGNode.Attributes.Item(0).nodeValue)
					sUoMName = trim(PGNode.Attributes.Item(1).nodeValue)
					sUoMShName = trim(PGNode.Attributes.Item(2).nodeValue)
					Response.Write("<OPTION VALUE="&trim(sUoMID)&">"&trim(sUoMShName)&"</OPTION>" &vbcrlf)
				next
			end if
		end if
	End Function 
%>

<%
''''''''Function to populate activities'''''
Function popActivity
Dim rsTemp, sSQL, iActivityNo, sActivityName

Set rsTemp = Server.CreateObject("ADODB.RecordSet")
	sSql = "SELECT ActivityNumber,ActivityName FROM MTN_M_ActivitiesForNoSeries order by ActivityNumber"
	with rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	end with
	set rsTemp.ActiveConnection = nothing

	set iActivityNo	= rsTemp(0)
	set sActivityName = rsTemp(1)

	Do While Not rsTemp.EOF
		Response.Write("<OPTION VALUE="""&iActivityNo&""">"&sActivityName&"</OPTION>" &vbcrlf)
		rsTemp.MoveNext
	Loop
	rsTemp.Close
End Function
%>

<%
Function CheckNoSeriesExistMaintenance(sOrgID,iActivityNo,sPassWorkGroupType)

	Dim sRetVal
	Dim iSeriesNo,iSeriesCode
	Dim rsNumber

	set rsNumber = Server.CreateObject("ADODB.RecordSet")	
	sRetVal = true
	
	'' To Generate Code (Using Number Series)'---------------------------------------------
	With rsNumber
		.CursorLocation = 3
		.CursorType = 3
		.Source =  "SELECT MainSeriesNo,MainSeriesCode FROM VwMtnNoSeriesSel where " &_
			" ActivityType=" & iActivityNo & " AND OrganisationCode='" & trim(sOrgID) & "' "&_
			" and WorkGroupValue in ('0','"&sPassWorkGroupType&"') "			
		.ActiveConnection = con
		.Open
	End With
	'Response.Write "<p> " & rsNumber.Source 
	Set rsNumber.ActiveConnection = nothing

	If rsNumber.EOF then
		sRetVal = false
	End If
	rsNumber.Close
	
	CheckNoSeriesExistMaintenance = sRetVal
End Function 
%>
