Function GetAttributes(gName,gPath,gKey)
    Dim gArr, gcode, sTemp
    gArr = Split(gKey, ":")

    If UBound(gArr) < 3 Then 
		alert("Select Classification Created")
		document.formname.pGroup.value = ""
		document.formname.action = "MasClassificationAttributeEntry.asp"
		document.formname.submit()
		Exit function
    end if
    sTemp = "C" & gArr(0) & ":" & gArr(UBound(gArr))

	document.formname.pGroup.value = sTemp
	document.formname.pName.value = gName
	document.formname.gPath.value = gPath
	document.formname.action = "MasClassificationAttributeEntry.asp"
	document.formname.submit()
End Function

Function NewGroupValidate(gKey)
Dim sFlag
    If Not gKey = "GRP" Then
        Dim gArr, gcode, sTemp
        
        gArr = Split(gKey, ":")
        
        '''''''''''''''''''''''''''''''''''''''''
        If UBound(gArr) < 3 Then 
        	alert("Cannot create Classification from here. Select Classification or Existing Classification.")
        	Exit function
        end if

        gcode = gArr(UBound(gArr))
        sTemp = "C" & gArr(0) & ":" & gArr(UBound(gArr))
        '''''''''''''''''''''''''''''''''''''''''
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		objhttp.Open "GET","../components/NewGroupVal.asp?Groupcode="&gcode, false
		objhttp.send 

		if objhttp.responseXML.xml <> "" then
			Data.loadXML objhttp.responseXML.xml
			Set Root = Data.documentElement
            If Root.hasChildNodes() Then
                For Each pageNode In Root.childNodes
                    rec0 = pageNode.attributes.Item(0).nodeValue
                    
                    If Trim(rec0) = "-1" Then
						document.formname.pGroup.value = sTemp
						document.formname.action = "MasClassificationNameEntry.asp"
						document.formname.submit()
                    Else
                        alert("Classification cannot be created since Item has been Created under this Classification")
						document.formname.pGroup.value = ""
						document.formname.action = "MasClassificationNameEntry.asp"
						document.formname.submit()
                    End If
                Next
            End If
        Else
            alert("No Data")
            Exit function
        End If
    Else
		document.formname.pGroup.value = gKey
		document.formname.action = "MasClassificationNameEntry.asp"
		document.formname.submit()
    End If
End function
