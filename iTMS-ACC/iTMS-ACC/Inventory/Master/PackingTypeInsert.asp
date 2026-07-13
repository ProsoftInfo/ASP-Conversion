<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	PackingTypeInsert.asp
	'Module Name				:	INVENTORY (Master)
	'Author Name				:	UmaMaheswari S
	'Created On					:	May 31, 2011
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
<!--#include file="../../include/Accpopulate.asp"-->
<!--#include file="../../include/populate.asp"-->
<%
Dim dcrs,sQuery,objDOM,Root,Node,sType
Dim sPackShortName,sPackName,sAltName,sNumberingType,sManualLotNumbering
Dim nPackingCode,sSerialWithinLot,sEnforce,nNoOfSubLevels,sManualSerialNumbering
Dim nSubLevelId,sSubLevelName
Dim sGrossLabel,sTare,sTareLabel

set dcrs  = server.CreateObject("adodb.recordset")
set objDOM = Server.CreateObject("Microsoft.XMLDOM")

objDOM.async =False
objDOM.Load(Request)

set Root = objDOM.documentElement

	con.BeginTrans

	sType = Root.getAttribute("Type")
	
	If sType ="C" Then
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISNULL(MAX(PackingCode)+1,1) FROM APP_M_PackingType"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.EOF then
			nPackingCode = dcrs(0)
		end if
		dcrs.Close
	Else
		nPackingCode = Root.getAttribute("PackCode")
		
		sQuery = "Delete From APP_M_PackingTypeSubLevel Where PackingCode="& nPackingCode &" "
		con.Execute sQuery
		
		sQuery = "Delete From APP_M_PackingType Where PackingCode="& nPackingCode &" "
		con.Execute sQuery
	End IF
	
	sPackShortName = Root.getAttribute("ShortName")
	sPackName = Root.getAttribute("Name")
	sAltName = Root.GetAttribute("AltLabel")
	sNumberingType = Root.getAttribute("ReceiptNumbering")
	sManualLotNumbering = Root.getAttribute("LotNoSelection")		'Manual or Auto
	sManualSerialNumbering = Root.getAttribute("SerialNoSelection")	'Manual or Auto
	sSerialWithinLot = Root.getattribute("SerialNoWithinLotCheck")	'Check or Not
	sEnforce = Root.getAttribute("LotNoEnforceCheck")				'Check Or Not
	nNoOfSubLevels = Root.getAttribute("NoOfSubLevel")
	
	if trim(nNoOfSubLevels)="" or IsNull(nNoOfSubLevels) then nNoOfSubLevels=0
	
	sGrossLabel = Root.getAttribute("GrossLabel")
	sTare = Root.getAttribute("Tare")
	sTareLabel = Root.getAttribute("TareLabel")
	
	if trim(sGrossLabel)="" or IsNull(sGrossLabel) then sGrossLabel="NULL"
	if trim(sGrossLabel)<>"NULL" then sGrossLabel=Pack(sGrossLabel)
	
	if trim(sTareLabel)="" or IsNull(sTareLabel) then sTareLabel="NULL"
	if trim(sTareLabel)<>"NULL" then sTareLabel=Pack(sTareLabel)
	
	
	sQuery = " Insert into APP_M_PackingType(PackingCode,PackingShortName,PackingName,AlternateName,NumberingType,ManualLotNumbering,ManualSerialNumbering,SerialWithinLot,NoOfSubLevels,Enforce,GrossPerPackLabel,TarePerPack,TarePerPackLabel) VALUES "&_
			 " ("& nPackingCode &",'"&sPackShortName &"','"& sPackName &"','"& sAltName &"','"& sNumberingType &"','"&sManualLotNumbering &"','"&sManualSerialNumbering &"','"& sSerialWithinLot &"',"& nNoOfSubLevels &",'"& sEnforce &"',"& sGrossLabel &","& pack(sTare) &","& sTareLabel &")"
	'Response.Write "<p>Query="&sQuery
	con.Execute sQuery
	
	If Root.hasChildNodes Then
	
		For Each Node in Root.childNodes
			nSubLevelId = Node.getAttribute("LevelNo")
			sSubLevelName = Node.getAttribute("LevelLabel")
			
			sQuery = "Insert into APP_M_PackingTypeSubLevel(PackingCode,SubLevelId,SubLevelName) Values "&_
					 " ( "& nPackingCode &","& nSubLevelId &",'"& sSubLevelName &"')"
			'Response.Write "<p>Query1="&sQuery
			con.Execute sQuery
		Next
	
	End IF	'If Root.hasChildNodes Then

If con.Errors.count <> "0" Then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & vbCrLf
	next
Else
	'con.RollbackTrans
	'Response.End 
	con.CommitTrans
End IF
%>