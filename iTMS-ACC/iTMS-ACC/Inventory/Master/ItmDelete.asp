<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmDelete.asp
	'Module Name				:	Inventory (Item Creation and Definition)
	'Author Name				:	S.MAHESWARI
	'Created On					:	April 05, 2008
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	ItmCreationDefinitionEntry.asp
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
<!--#include virtual="/include/populate.asp"-->
<%
Dim dcrs,dcrs1,iItemCode,iClassCode,iOrgCode
Dim sPara,sSql,sTransType,sTemp,i,k,sArrTransType
Set dcrs = Server.CreateObject("ADODB.Recordset")
sPara = trim(Request("hItemCode"))
'Response.Write "ORGID="& trim(Request("hOrgId"))
'Response.End 
con.beginTrans
If trim(sPara) <> "" then
	sTemp = split(sPara,",")
'	sSql = "Select distinct TransactionType from Inv_T_ItemLedger where TransactionType <> 'RO' "
'	Response.Write sSql
'	with dcrs
'		.CursorLocation = 3
'		.CursorType = 3
'		.Source = sSql
'		.ActiveConnection = con
''		.Open 
'	end with
'	Do while not dcrs.EOF 
'		sArrTransType = sArrTransType &":"& dcrs(0)
'		dcrs.MoveNext 
''	loop
'	dcrs.Close
'	sArrTransType = mid(sArrTransType,2)
'	'Response.Write "sArrTransType=" & sArrTransType
''	
	For i = 0 to UBOUND(sTemp)
		iItemCode = sTemp(i)
		iClassCode = trim(Request("hClassCode"))
		iOrgCode = trim(Request("hOrgId")) 
		 'Response.Write iItemCode		
			
	
			sSql = " Select TransactionType from Inv_T_ItemLedger where ItemCode = "& iItemCode &" and "&_
				   " ClassificationCode = "& iClassCode &" and OrganisationCode = '"& iOrgCode &"' "
				   'Response.Write "<p>"&sSql
			with dcrs
				.CursorLocation = 3
				.CursorType = 3
				.Source = sSql
				.ActiveConnection = con
				.Open 
			end with
	    if not dcrs.EOF then
			Do while Not dcrs.EOF 
	    		sArrTransType = sArrTransType & ":" & dcrs(0)
				'sTransType = dcrs(0)		 				 		
				dcrs.MoveNext 
			loop		
			
			sArrTransType = mid(sArrTransType,2)
	    end if 'if not dcrs.EOF then
	    dcrs.Close 
			IF trim(sArrTransType) <> "" then 
				sTemp = split(sArrTransType,":")
				For k = 0 to UBOUND(sTemp)
					if sTemp(k) <> "RO" then 
						Response.Write "This Item Cannot be Deleted"
						Response.End 
					end if
				Next
			End IF
	
			'IF len(sTransType) > 2 then 
				
			'Response.End 
			'IF sTransType = "RO"  then 
			'Delete only the item in Itemledger where transactiontype is only "RO" not othere than that
			
			sSql = "Delete from Inv_T_LocationLot where ItemCode = "& iItemCode &" and "&_
				   " ClassificationCode = "& iClassCode &" and OrganisationCode = '"& iOrgCode &"' "
			Response.Write "<BR>"&sSql &"<BR><BR>"
			Con.Execute sSql
	
			sSql = "Delete from Inv_T_ItemYearlyStock where ItemCode = "& iItemCode &" and "&_
				   " ClassificationCode = "& iClassCode &" and OrganisationCode = '"& iOrgCode &"' "
			Response.Write sSql&"<BR><BR>"
			Con.Execute sSql
	
			'sSql = "Delete from INV_T_ItemStockMatrix"
	
			sSql = "Delete from Inv_T_ItemLocationStock where ItemCode = "& iItemCode &" and "&_
				   " ClassificationCode = "& iClassCode &" and OrganisationCode = '"& iOrgCode &"' "
			Response.Write sSql&"<BR><BR>"
			Con.Execute sSql
	
			sSql = "Delete from Inv_T_ItemLedger where ItemCode = "& iItemCode &" and "&_
				   " ClassificationCode = "& iClassCode &" and OrganisationCode = '"& iOrgCode &"' "
			Response.Write sSql&"<BR><BR>"
			Con.Execute sSql
			
			sSql = "Delete from INV_M_ITEMMASTERBOM where ItemCode = "& iItemCode &" and "&_
			        "  ClassificationCode = "& iClassCode &" and Organisationcode = '"& iOrgCode &"'"
			Response.Write "<BR>"& sSql &"<BR><BR>"
			con.execute sSql
	
			sSql = "Delete from Inv_M_ItemOptionalUoM where ItemCode = "& iItemCode &" and "&_
				   " ClassificationCode = "& iClassCode &" and OrganisationCode = '"& iOrgCode &"' "
			Response.Write sSql&"<BR><BR>"
			Con.Execute sSql
	
			sSql = "Delete from Inv_M_ItemStorage where ItemCode = "& iItemCode &" and "&_
				   " ClassificationCode = "& iClassCode &" and OrganisationCode = '"& iOrgCode &"' "
			Response.Write sSql&"<BR><BR>"
			Con.Execute sSql
			
			sSql = "Delete from INV_M_ItemMasterAttributes where ItemCode = "& iItemCode &" and "&_
				   " ClassificationCode = "& iClassCode &" and OrganisationCode = '"& iOrgCode &"' "
			Response.Write sSql&"<BR><BR>"
			Con.Execute sSql
	
	
			sSql = "Delete from Inv_M_ItemMaster where ItemCode = "& iItemCode &" and "&_
				   " ClassificationCode = "& iClassCode &" and OrganisationCode = '"& iOrgCode &"' "
			Response.Write sSql&"<BR><BR>"
			Con.Execute sSql
	
	
	
	Next 'For i = 0 to UBOUND(sTemp)
End If 'If trim(sPara) <> "" then
Response.Clear

'Response.End 
	
	if con.Errors.count <> 0 then
		dim iErrCounter
		con.RollbackTrans
		for iErrCounter=0 to con.Errors.count
			Response.Write con.Errors(iErrCounter) & "<BR>"
		next
		'Redirect to Error Handling System
	else		
		'con.RollbackTrans
		'Response.End
		con.CommitTrans
	end if

	con.close
	set con = nothing
	Response.Redirect ("ITEMLISTENTRY.ASP")

%>

