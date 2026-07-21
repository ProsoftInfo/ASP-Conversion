
<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	mrsIssueInsert.asp
	'Module Name				:	Inventory (Issue)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	APRIL 01,2010
	'Modified On				:	Dec 28,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	mrsIssueItemEntry.asp
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
<!--#include file="../../include/POInsertCommon.asp"-->

<%
Dim sAppCallFrom,rsIssObj,sIssQuery,sPONO
sAppCallFrom = Request("hCallFrom")
set rsIssObj= server.CreateObject("ADODB.Recordset")

    con.begintrans

    '' To Call the Issue Insert Common Function
    MrsIssueInsert
    
    
   
    if con.Errors.count <> 0 then
		con.RollbackTrans
		for iCounter=0 to con.Errors.count
			Response.Write con.Errors(iCounter) & vbCrLf
		next
		'Redirect to Error Handling System
	else
		'Response.Write "<p>sSalInvConfirm="&sSalInvConfirm
	'	Response.Clear 
		
		if sSInvType ="X" then sSInvType ="CB"
        if sSInvType ="Y" then sSInvType ="NEB"
        if sSInvType ="Z" then sSInvType ="EB"
        Response.Write "<p>sSinvType = "& sSinvType
        Response.Write "<p> sSalType = "& sSSalType
        Response.Write "<p> POS = "& sSALPOSID
        Response.Write "<p> sSinvType = "& sSinvTypeName
        Response.Write "<p>  sSalType = "& sSSalTypeName
        Response.Write "<p>  POS = "& sSSALPOSIDName
        
        sIssQuery = "Select AppRefNo from INV_T_MRSHeader where MRSNumber = "& sAppRefNo 
	    rsIssObj.Open sIssQuery,con
	    if not rsIssObj.EOF then
	        sPONO = rsIssObj(0)
	    end if 
	    rsIssObj.Close 
	    
	    Response.Write "<p>sAppCallFrom="& sAppCallFrom
        		
	'	con.RollbackTrans
	'	Response.End
	   Response.Clear
	   con.CommitTrans
	   
	 '   if trim(sAppCallFrom)="SUB" then
	'	    Response.Redirect "../../Purchase/Transaction/POLIST.ASP?ACTN=L&CallFrom=ISS&ISSNO="&iLedIssueNo&"&PONO="&sPONO
	 '   else
	  '      Response.Redirect(sRedirectTo)
	   ' end if 

	'	if trim(sIssType)<>"M" and trim(sCallFrom)<>"MR" then
			''Dont Delete or Block this Below Print It is the Agrgument form Direct Invoice Creation
			''Added by Ragav March 30,2010
			if trim(sType)="SUB" then
				if trim(sPOConfirm)="Y" then
		                con.begintrans
		                
		                POInsert()
		                
                        Response.Clear
                        if con.Errors.count <> 0 then
		                    con.RollbackTrans
		                    for iCounter=0 to con.Errors.count
			                    Response.Write con.Errors(iCounter) & vbCrLf
		                    next
		                    'Redirect to Error Handling System
	                    else

	                    '	con.RollbackTrans
	                    '	Response.End
		                    con.CommitTrans
		                    
		                    if trim(sProformaConfirm)="Y" then
								Response.Redirect("../../Sales/Transaction/SalTrProInvoice.asp?hInvNo="&sInvNo)
							else
								Response.Redirect(sRedirectTo)
							end if
		                end if 'if con.Errors.count <> 0 then
				else
					Response.Redirect(sRedirectTo)
				end if
			elseif trim(IssToCode)="DIS" then
				if trim(sSalInvConfirm)="Y" then
					if trim(sSelectedInvoice)="A" then
					     '   If sSInvType = "CB" or sSInvType = "NEB" or sSInvType = "EB" Then
		                        Response.Redirect("../../Sales/Transaction/SalInvoiceEntry_Trading.asp?hInvNo="&sInvNo&"&InvType="&sSInvType&":"&sSInvTypeName&"&TaxType="&sSSalType&":"&sSSalTypeName&"&POS="&sSALPOSID&":"&sSSALPOSIDName&"&CallFrom=ISSUE&AppRefNo="&iLedIssueNo&"&AppRefType=12")
	                     '   Else
		                 '       Response.Redirect("../../Sales/Transaction/SalInvoiceEntry.asp?hInvNo="&sInvNo&"&InvType="&sSInvType&"&CallFrom=ISSUE&AppRefNo="&iLedIssueNo&"&AppRefType=12")
	                     '   End IF
					elseif trim(sSelectedInvoice)="P" then
						Response.Redirect("../../Sales/Transaction/SalTrProInvoice.asp?hInvNo="&sInvNo&"&InvType="&sSInvType&"&CallFrom=ISSUE&AppRefNo="&iLedIssueNo&"&AppRefType=12")
					end if
				else
					Response.Redirect(sRedirectTo)
				end if
			elseif trim(IssToCode)="SER" or trim(IssToCode)="JWK" then
				if trim(sGatePassConfirm)="Y" then
					Response.Redirect("GatePassServiceEntryAmd.asp?GatePassNo="&sGatePassNo)
				else
					Response.Redirect(sRedirectTo)
				end if
			else
				Response.Redirect(sRedirectTo)
			end if
	'	else
	'	    Response.Redirect(sRedirectTo)
	'	end if ' if trim(sIssType)<>"M" and trim(sCallFrom)<>"MR" then
	end if 'if con.Errors.count <> 0 then
%>
