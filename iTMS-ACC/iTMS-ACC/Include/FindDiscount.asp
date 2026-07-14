
<%
'	Program : FindDiscount
'	Created BY: Ragavendarn
'	Created On : July 16,2011
%>
<!--#include virtual="/Include/DatabaseConnection.asp"-->
<%	
	Dim ItemCode,ClassCode,OrgCode,Qty,Value,PartyCode
	Dim sReturnValue,rsCusDis,rsItemDis,sQuery,sPrecedence,sEligible
    Dim iQtyDis,iQtyFrom,iQtyTo,iValDis,iValFrom,iValTo
    
    set rsCusDis = Server.CreateObject("ADODB.Recordset")
    set rsItemDis = Server.CreateObject("ADODB.Recordset")
    
    sEligible = false 
    
	ItemCode = Request.QueryString("ItemCode")
	ClassCode = Request.QueryString("ClassCode")
	OrgCode  = Request.QueryString("OrgCode")
	Qty = Request.QueryString("Qty")
	Value =Request.QueryString("Value")
	PartyCode= Request.QueryString("PartyCode")


    If PartyCode<>"" and PartyCode<>"0" Then 
       sQuery = "Select QtyDiscountOffered,QuantityFrom,QuantityTo,ValueDiscountOffered,ValueFrom,"&_
                " ValueTo,Precedence from INV_T_ItemSupplierDiscount where PartyCode = "& PartyCode &" and "&_
                " ItemCode = "& ItemCode &" and ClassificationCode = "& ClassCode &" and OrganisationCode = '"& OrgCode &"' "
    Else
       sQuery = "Select QtyDiscountOffered,QuantityFrom,QuantityTo,ValueDiscountOffered,ValueFrom,"&_
                " ValueTo,Precedence from Inv_M_ItemOrgSaleDiscount where ItemCode = "& ItemCode &" and "&_
                " ClassificationCode = "& ClassCode &" and OrganisationCode = '"& OrgCode &"' "
    End if
   ' Response.Write sQuery
    rsCusDis.Open sQuery,con
    If not rsCusDis.eof Then
       do while not rsCusDis.eof
           iQtyDis = rsCusDis(0)
           iQtyFrom = rsCusDis(1)
           iQtyTo = rsCusDis(2)
           iValDis = rsCusDis(3)
           iValFrom = rsCusDis(4)
           iValTo = rsCusDis(5)
           sPrecedence = rsCusDis(6)
           if sPrecedence = "Q" then
               if cdbl(Qty)>=cdbl(iQtyFrom) and cdbl(Qty) <=cdbl(iQtyTo) then
                   sEligible = true
                   sReturnValue = iQtyDis 
                   exit Do
               end if
           elseif sPrecedence = "V" then
               if cdbl(Value)>=cdbl(iValFrom) and cdbl(Value) <=cdbl(iValTo) then
                   sEligible = true
                   sReturnValue = iValDis 
                   exit Do
               end if
           end if 
           rsCusDis.movenext
       loop
    End If
    rsCusDis.close
    if (PartyCode<>"" and PartyCode<>"0") and sEligible = false then
       sQuery = "Select QtyDiscountOffered,QuantityFrom,QuantityTo,ValueDiscountOffered,ValueFrom,"&_
                " ValueTo,Precedence from Inv_M_ItemOrgSaleDiscount where ItemCode = "& ItemCode &" and "&_
                " ClassificationCode = "& ClassCode &" and OrganisationCode = "& OrgCode
       ' Response.Write "----"&sQuery
        rsCusDis.Open sQuery,con
        If not rsCusDis.eof Then
           do while not rsCusDis.eof
               iQtyDis = rsCusDis(0)
               iQtyFrom = rsCusDis(1)
               iQtyTo = rsCusDis(2)
               iValDis = rsCusDis(3)
               iValFrom = rsCusDis(4)
               iValTo = rsCusDis(5)
               sPrecedence = rsCusDis(6)
               if sPrecedence = "Q" then
                   if cdbl(Qty)>=cdbl(iQtyFrom) and cdbl(Qty) <=cdbl(iQtyTo) then
                       sEligible = true
                       sReturnValue = iQtyDis 
                       exit Do
                   end if
               elseif sPrecedence = "V" then
                   if cdbl(Value)>=cdbl(iValFrom) and cdbl(Value) <=cdbl(iValTo) then
                       sEligible = true
                       sReturnValue = iValDis 
                       exit Do
                   end if
               end if 
               rsCusDis.movenext
           loop
       End If
       rsCusDis.close
    End if ' if (PartyCode<>"" and PartyCode<>"0") and sEligible = false then
    Response.Write sReturnValue
%>
