<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	NoSeriesDetailsEntry.asp
	'Module Name				:	Number Series Transfer
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 19, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	NoSeriesDetailsInsert.asp
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
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%
'XML DOM Variables
dim YrNoSeriesXML,RootNode,EntryNode,sQuery,sResult,sSql,objFSO

dim dcrs,sUnit,sUnitName,dPreFinStartDate,dPreFinEndDate,dCurFinStartDate,dCurFinEndDate
dim iPrePeriodFrom,iPrePeriodTo,iPrePeriod,iCurPeriod,iApplication
dim iSeriesNo,iSeriesCode,iEntryNo,sPreFix,sSuffix,sPeriod,sCurFinPeriod,sPrevFinPeriod
dim Rt,oDOM,sTemp,Arr,node
set objFSO = Server.CreateObject("Scripting.FileSystemObject")

sTemp = Request("Temp")
IF trim(sTemp) <> "" then
	Arr = Split(sTemp,"||")
	sUnit = trim(Arr(0))
	sUnitName = trim(Arr(1))
	dCurFinStartDate = trim(Arr(2))
	dCurFinEndDate = trim(Arr(3))
	dPreFinStartDate = trim(Arr(4))
	dPreFinEndDate = trim(Arr(5))
	iApplication = trim(Arr(6))


End IF
'Response.Write sTemp
sPrevFinPeriod = right(dPreFinStartDate,4) &":" & right(dPreFinEndDate,4)
sCurFinPeriod = right(dCurFinStartDate,4) & ":" & right(dCurFinEndDate,4)
'sTemp = Request("NoSer")
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
'oDOM.async = false
'oDOM.load (Request)
if objFSO.FileExists(server.MapPath("../temp/NoSeries.xml")) then
    oDOM.load server.MapPath("../temp/NoSeries.xml")
    Set Rt = oDOM.DocumentElement
else
    Set Rt = oDOM.createElement("Root")
    oDOM.appendChild Rt
end if




iPrePeriodFrom = right(dPreFinStartDate,4)&mid(dPreFinStartDate,4,2)
iPrePeriodTo = right(dPreFinEndDate,4)&mid(dPreFinEndDate,4,2)

' Create our DOM Document Objects
Set YrNoSeriesXML = Server.CreateObject("Microsoft.XMLDOM")

Set dcrs = Server.CreateObject("ADODB.RecordSet")

YrNoSeriesXML.async=false

con.beginTrans

sQuery = "SELECT Distinct OUDEFINITIONID,SERIESNO,SERIESCODE,ENTRYNO,ISNULL(PREFIX,'-') AS PREFIX,ISNULL(SUFFIX,'-') AS SUFFIX,PERIOD FROM APP_R_NOSERIESMODULEENTRY WHERE (STR(OUDEFINITIONID)+STR(SERIESNO)+STR(SERIESCODE)) IN (SELECT (STR(OUDEFINITIONID)+STR(SERIESNO)+STR(SERIESCODE)) FROM APP_R_NOSERIESMODULES WHERE OUDEFINITIONID = " & Pack(sUnit) & " AND APPLICATIONCODE = " & iApplication & " AND (PERIOD >= " & iPrePeriodFrom & " AND PERIOD <= " & iPrePeriodTo & ")) ORDER BY 1,2,3 FOR XML AUTO"
'Response.Write sQuery
'Response.End
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing
if not dcrs.EOF then
	do while not dcrs.EOF
		sResult = sResult & dcrs(0)
	dcrs.MoveNext
	loop
else
	MsgBox ("Number Series already been carry forwarded or created for the Current Financial Year")
	Response.Redirect("CloseEntry.asp?Frm=NS")
	Response.End
end if
dcrs.Close
sResult = "<ROOT>" & sResult & "</ROOT>"

YrNoSeriesXML.loadXML sResult
YrNoSeriesXML.save server.MapPath("../Temp/"&iApplication&"-NumberSeriesDataBefore.xml")


set RootNode = YrNoSeriesXML.documentElement

if RootNode.HaschildNodes() then

	for each EntryNode in RootNode.ChildNodes
		iSeriesNo = trim(EntryNode.attributes.getNamedItem("SERIESNO").value)
		iSeriesCode = trim(EntryNode.attributes.getNamedItem("SERIESCODE").value)
		iEntryNo = trim(EntryNode.attributes.getNamedItem("ENTRYNO").value)
		sPreFix = trim(EntryNode.attributes.getNamedItem("PREFIX").value)
		sSuffix = trim(EntryNode.attributes.getNamedItem("SUFFIX").value)
		iPrePeriod = trim(EntryNode.attributes.getNamedItem("PERIOD").value)

		if right(iPrePeriod,2) = "01" or right(iPrePeriod,2) = "02" or right(iPrePeriod,2) = "03" then
			sPeriod = right(dCurFinEndDate,4)&right(iPrePeriod,2)
		else
			sPeriod = right(dCurFinStartDate,4)&right(iPrePeriod,2)
		end if

		if sPreFix = "-" then
			sPreFix = "NULL"
		else
			sPreFix = Pack(sPreFix)
		end if
		if Rt.HaschildNodes() then
				for each node in Rt.childnodes
					if trim(node.NodeName) = "NoSeries" then
						if strComp(trim(node.getAttribute("ExistingSuffix")),sSuffix) = 0 then
							sSuffix = trim(node.getAttribute("ChangeSuffix"))
							exit for
						else
							'exit for
						end if
					end if
				next
			End IF

		if sSuffix = "-" then
			sSuffix = "NULL"
		else
			sSuffix = pack(sSuffix)
		end if

		sQuery = "SELECT OUDEFINITIONID FROM APP_R_NOSERIESMODULEENTRY WHERE OUDEFINITIONID = " & Pack(sUnit) & " AND SERIESNO = " & iSeriesNo & " AND SERIESCODE = " & iSeriesCode & " AND ENTRYNO = " & iEntryNo & " AND PERIOD = " & Pack(dCurFinEndDate) & ""
		'Response.Write sQuery
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = sQuery
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if dcrs.EOF then
			sSql = "INSERT INTO APP_R_NOSERIESMODULEENTRY (OUDEFINITIONID,SERIESNO,SERIESCODE," &_
				"ENTRYNO,PERIOD,NUMBER,PREFIX,SUFFIX) VALUES " &_
				"(" & Pack(sUnit) & "," & iSeriesNo & "," & iSeriesCode & "," &_
				"" & iEntryNo & "," & Pack(sPeriod) & ",1," & sPreFix & "," & sSuffix & ")"
			 Response.Write sSql & "<BR>"
			con.Execute sSql
			'Exit For
		end if
		dcrs.Close
		sSuffix =""
	next
	YrNoSeriesXML.save server.MapPath("../Temp/"&iApplication&"-NumberSeriesDataAfter.xml")
	If iApplication = 1 Then
		sSql = "Insert Into Acc_M_BookNumberSeries  Select OUDefinitionID,BookCode,BookNumber,'"&sCurFinPeriod&"',DrSeriesNo,DrSeriesCode,"
		sSql = sSql&"CrSeriesNo,CrSeriesCode,CreatedDrSeriesNo,CreatedDrSeriesCode,CreatedCrSeriesNo,"
		sSql = sSql&"CreatedCrSeriesCode,LastChequeNo,ReceiptNo from Acc_M_BookNumberSeries where FinPeriod = '"&sPrevFinPeriod&"'"
		sSql = sSql&" and OUDefinitionID = '"&sUnit&"'"
		Response.Write sSql
		con.Execute sSql
	End If
	if iApplication = 2 then
		'Under Purchase Module reset the LAst Number column to 0 in these two tables

		'tpotrlastnocont
		'tpotrpolastcont

		sSql = "UPDATE TPOTRLASTNOCONT SET LASTNUMBER = 0"
		'Response.Write sSql & "<BR>"
		'con.Execute sSql

		sSql = "UPDATE TPOTRPOLASTCONT SET LASTNUMBER = '000'"
		'Response.Write sSql & "<BR>"
		'con.Execute sSql

	end if
end if

if con.Errors.count <> 0 then
	dim iErrCounter
	con.RollbackTrans
	for iErrCounter=0 to con.Errors.count
		Response.Write con.Errors(iErrCounter) & "<BR>"
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
MsgBox("Number Series has been carry forwarded to the Current Financial Year")
Response.Redirect ("CloseEntry.asp?Frm=NS")
%>

