<%@ Language=VBScript %>
<%	option explicit	
	Response.Expires=-10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ContactInsert.asp.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Kalaiselvi R
	'Created On					:	sep 28,2011
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
<!--#include file="../../include/sessionVerify.asp"-->
<%
dim sName,sDesignation,sAddress1,sAddress2,sPincode,sCity,sState,sCountry,sPhone
Dim sFax,sAction,sEmail,sWebsite,iContactNo,sMobile,iCreatedBy,sContactPersonFor
Dim oDOM,Root,newElem,newElem1,sCallTy,Temparr,ndChild,iPartyCode,iExistingPartyCode
Dim iExistingPartyContactNo,iGeneratePartyContactNo
dim sQuery,objRs,sInActive,rsTemp


Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = Server.CreateObject("ADODB.RecordSet")
Set rsTemp = Server.CreateObject("ADODB.RecordSet")



iCreatedBy = getUserid()

iContactNo = Request.Form("hContactNumber")
'Response.Write "iContactNo = "& iContactNo
sAction = Request.QueryString("Action")

'Response.Write "sAction = "& sAction

Con.BeginTrans

if trim(iContactNo)<>"" then

	iExistingPartyCode = 0
	iExistingPartyContactNo = 0
	
	sQuery = "Select PartyCode,PartyContactNo from APP_M_Contacts where ContactNumber="& iContactNo
	objRs.Open sQuery,con 
	if not objRs.EOF then
		iExistingPartyCode = objRs(0)
		iExistingPartyContactNo = objRs(1)
	end if 
	objRs.Close 
						
	if trim(iExistingPartyCode) <> "0" then
	
		sQuery ="Delete from APP_M_PartyContactPersons where PartyCode="& iExistingPartyCode & " and ContactNo = " & iExistingPartyContactNo & ""
		Response.Write "<p>"&sQuery
		con.execute sQuery
		
		sQuery ="Update APP_M_PartyContactPersons set ContactNo =  ContactNo -1 where PartyCode="& iExistingPartyCode & " and ContactNo > " & iExistingPartyContactNo  & ""
		Response.Write "<p>"&sQuery
		con.execute sQuery
			
	
	end if 'if trim(iExistingPartyCode) <> "0" then
	
	'sQuery ="Delete from APP_M_Contacts where ContactNumber="& iContactNo
	'Response.Write "<p>"&sQuery
	'con.execute sQuery
	
	
end if ' if trim(iContactNo)<>"" then



oDOM.load(server.MapPath("../temp/master/Contact_Master_"&Session.SessionID&".xml"))

	Set Root = oDOM.documentElement
	
	
	if Root.hasChildNodes() then
		for each ndChild in Root.childNodes
			if ndChild.nodeName="ParName" then
				sName = ndChild.text
			elseif ndChild.nodeName="Designation" then
				sDesignation = ndChild.text
			elseif ndChild.nodeName="ContactPersonFor" then
				sContactPersonFor= ndChild.text
			elseif ndChild.nodeName="PartyCode" then
				iPartyCode	= ndChild.text
			elseif ndChild.nodeName="Address1" then
				sAddress1 = ndChild.text
			elseif ndChild.nodeName="Address2" then
				sAddress2 = ndChild.text
			elseif ndChild.nodeName="PinCode" then
				sPincode = ndChild.text
			elseif ndChild.nodeName="City" then
				sCity = ndChild.text
			elseif ndChild.nodeName="State" then
				sState = ndChild.text
			elseif ndChild.nodeName="Country" then
				sCountry= ndChild.text
			elseif ndChild.nodeName="Phone" then
				sPhone = ndChild.text
			elseif ndChild.nodeName="Mobile" then
				sMobile = ndChild.text
			elseif ndChild.nodeName="Fax" then
				sFax = ndChild.text
			elseif ndChild.nodeName="Email" then
				sEmail = ndChild.text
			elseif ndChild.nodeName="Website" then
				sWebsite = ndChild.text
			elseif ndChild.nodeName="CreatedBy" then
				iCreatedBy = ndChild.text
			elseif ndChild.nodeName="Active" then
				sInActive = ndChild.getAttribute("Flag")
			end if
		next
	end if
	
sName = mid(sName,1,100)

if trim(iPartyCode) = "" then
	iPartyCode = 0
end if 

if trim(sInActive) = "" then sInActive = 0
	
if trim(sAction)="Edit" then

	iGeneratePartyContactNo = 0		
	if trim(iPartyCode) <> "0" then	
		InsertPartyContactData()	
	end if 'if trim(iPartyCode) <> "0" then
	
			
	sQuery= "Update APP_M_Contacts set ContactName='"&sName&"',Designation='"&sDesignation&"',"&_
			"ContactPersonFor='" & sContactPersonFor & "',"&_
			"AddressLine1='"&sAddress1&"',AddressLine2='"&sAddress2&"',City='"&sCity&"',"&_
			"State='"&sState&"',Country='"&sCountry&"',PhoneNos='"&sPhone&"',MobileNos='"&sMobile&"',"&_
			"FaxNos='"&sFax&"',Email='"&sEmail&"',WebsiteURL='"&sWebsite&"',Pincode='"&sPincode&"',"&_
			"Useable="& sInActive &",PartyCode=" & iPartyCode & ",PartyContactNo=" & iGeneratePartyContactNo & " where  ContactNumber = "& iContactNo
		
	Response.Write "<p>"&sQuery
	Con.Execute sQuery
	
else
	
	iGeneratePartyContactNo = 0
	if trim(iPartyCode) <> "0" then	
		InsertPartyContactData()
	end if 'if trim(iPartyCode) <> "0" then
			
	
	sQuery = "Select isNull(Max(ContactNumber),0)+1 from APP_M_Contacts"
	objRs.Open sQuery,con 
	if not objRs.EOF then
		iContactNo = objRs(0)
	end if 
	objRs.Close 
	
	
	sQuery = "Insert into APP_M_Contacts (ContactNumber,ContactName,Designation,ContactPersonFor,AddressLine1,AddressLine2,"&_
			"City,State,Country,PhoneNos,MobileNos,FaxNos,Email,WebsiteURL,Pincode,"&_
			"CreatedBy, CreatedOn,Useable,PartyCode,PartyContactNo)"&_
			" values ("&iContactNo&",'"&sName&"','"&sDesignation&"','" & sContactPersonFor & "','"&sAddress1&"','"&sAddress2&"'," &_
			"'"&sCity&"','"&sState&"', '"&sCountry&"','"&sPhone&"','"&sMobile&"','"&sFax&"','"&sEmail&"','"&sWebsite&"', '"&sPincode&"',"&_
			""& iCreatedBy &",getdate(),'"& sInActive &"'," & iPartyCode & ","  & iGeneratePartyContactNo & ")"
	
	'Response.Write "<p>"&sQuery
	Con.Execute sQuery
	
	
end if

	'Con.RollbackTrans
	'Response.End 

	Response.Clear 
	Con.CommitTrans

	Response.Redirect "ContactsList.asp"

%>

<%
Function InsertPartyContactData()

	iGeneratePartyContactNo = 1
	
	sQuery = "Select isNull(Max(ContactNo),0)+1 from APP_M_PartyContactPersons where PartyCode="& iPartyCode & ""
	objRs.Open sQuery,con 
	if not objRs.EOF then
		iGeneratePartyContactNo = objRs(0)
	end if 
	objRs.Close 
	
	
	sQuery = "Insert into APP_M_PartyContactPersons(PartyCode,ContactNo,ContactPersonName,Designation,ContactPersonFor,ContactMailID)" &_
			" values (" & iPartyCode & "," & iGeneratePartyContactNo & ",'"& sName &"','"& sDesignation &"','"& sContactPersonFor &"','" & sEmail & "')"
	Response.Write "<p>"&sQuery
	Con.Execute sQuery
	

End Function
%>