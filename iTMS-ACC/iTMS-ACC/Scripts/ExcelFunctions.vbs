'--------Example for Text Display-----------------------------
'InsertCell oRow,1,"",iSno,"ExcelSerial","Center","",0,0,0,0,""

'--------Example for TextBox -----------------------------
'InsertCell oRow,2,"txtDocAmount"&CStr(sDocNo),"0","ExcelInputCell","","",12,10,0,0,""

Function InsertCell(oRow,iType,sName,sValue,sClass,sAlign,sValign,iSize,iMaxlen,iColspan,iRowspan,sOptions)
	dim objCell
	
	select case iType
		case 1:
				set objCell=oRow.insertCell()									
				objCell.innerHTML=sValue
		case 2:
				set objCell=oRow.insertCell()	
				if trim(sOptions)<>"" then								
					set objText = document.createElement("<input type=""text"" name=""" & sName & """ value=""" & sValue & """ size="""&iSize&""" maxlength="&iMaxlen&" class=""Formelem"" "&sOptions&">" )
				else
					set objText = document.createElement("<input type=""text"" name=""" & sName & """ value=""" & sValue & """ size="""&iSize&""" maxlength="&iMaxlen&" class=""Formelem"">" )
				end if					
				objCell.appendChild(objText)
		case 3:
				set objCell=oRow.insertCell()	
				if trim(sOptions)<>"" then								
					set objText = document.createElement("<input type=""checkbox"" name=""" & sName & """ value=""" & sValue & """ "&sOptions&">" )
				else
					set objText = document.createElement("<input type=""checkbox"" name=""" & sName & """ value=""" & sValue & """ >" )
				end if					
				objCell.appendChild(objText)
		case 4:
				set objCell=oRow.insertCell()	
				if trim(sOptions)<>"" then								
					set objText = document.createElement("<input type=""text"" name=""" & sName & """ value=""" & sValue & """ size="""&iSize&""" maxlength="&iMaxlen&" class=""FormelemRead"" "&sOptions&">" )
				else
					set objText = document.createElement("<input type=""text"" name=""" & sName & """ value=""" & sValue & """ size="""&iSize&""" maxlength="&iMaxlen&" class=""FormelemRead"">" )
				end if					
				objCell.appendChild(objText)
						
	end select

	objCell.className=sClass
	if trim (sAlign)<>"" then	objCell.align=sAlign
	if trim (sValign)<>"" then  objCell.valign=sValign
	if iColspan<>0 then  objCell.colspan=iColspan			
	if iRowspan<>0 then  objCell.rowspan=iRowspan						
	
end Function
'---------------------End Of Function InsertCell--------------------------

Function ClearTable(objTable,startlen,Count)
	dim i	
	for i=startlen to eval("document.all."&objTable).rows.length - Count
		eval("document.all."&objTable).deleteRow(startlen) 
	next
end Function
'---------------------End Of Function ClearTable--------------------------
