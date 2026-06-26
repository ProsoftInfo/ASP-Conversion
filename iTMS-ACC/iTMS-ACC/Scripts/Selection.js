
function enableButtonAdd()
{
	if (document.forms[0].add.disabled)
	{
		document.forms[0].add.disabled = false;
	}
}

function enableButtonRemove()
{
	if (document.forms[0].remove.disabled)
	{
		document.forms[0].remove.disabled = false;
	}
}

function removeclick(rightCombo , leftCombo  , removeBtn)
{
	var thelist = document.forms[0].elements[leftCombo]; //"select Item"
	var tmplist = document.forms[0].elements[rightCombo]; //"selected Item"

	tmplength = tmplist.options.length;

    	for(i = tmplist.options.length -1; i >=0 ; i--)
    	{
        	if(tmplist.options[i].selected)
        	{
            		thelist.options[thelist.options.length] = new Option(tmplist.options[i].text,tmplist.options[i].value);
            		tmplist.options[i] = null;
	 	}
    	}
    
    	if(tmplist.options.length == 0) 
    	{
	  document.forms[0].elements[removeBtn].disabled = true;
	  //document.forms[0].next.disabled = true;
    	}
		  
    	//tmplist.size = 5;
    	//thelist.size = 5;
}


function addclick(rightCombo , leftCombo  , removeBtn)
{
	var thelist = document.forms[0].elements[leftCombo]; //"selectItem"
	var tmplist = document.forms[0].elements[rightCombo]; //"selectedItem"
	
  	tmplength = tmplist.options.length;
       
	//tmplist.size=5;
        if(tmplength == -25)
		alert("Only 25 Items can be selected");
	else
	{
    		for( i = thelist.options.length -1; i >=0 ; i--)
    		{
        		if(thelist.options[i].selected)
        		{
            			newname = thelist.options[i].text;
            			newname1 = thelist.options[i].value;
            			addit = true;
            			for(j = tmplength-1; j >=0 ; j--)
            			{
				 	//if(tmplist.options[j].text == newname)
				 	if(tmplist.options[j].value == newname1)
				 	{
						alert("Selected item already  in the list");
                    				addit = false;
                    				break;
					}
	    			}
	    			if(addit)
	    			{
					tmplist.options.length = ++tmplength;
					tmplist.options[tmplength-1].text = newname;
					tmplist.options[tmplength-1].value = newname1;
					document.forms[0].elements[removeBtn].disabled = false;
					//document.forms[0].next.disabled = false;
					thelist.options[i] = null;
	    			}
			}
		}// end of for loop
	}// end of else

    //tmplist.size = 5;
    //thelist.size = 5;
}


function finaldone(rightCombo,hiddenFieldName)
{
	var i,par
	par="";
		
	for (i=0;i<(document.forms[0].elements[rightCombo].options.length)-1;i++)
	{
		par= par+document.forms[0].elements[rightCombo].options[i].value+":";
	}
	if(document.forms[0].elements[rightCombo].options.length==0)
	{
		//par="m-1"
		alert("Select the item");
		return false; 
	}
	else
	{
		par= par+document.forms[0].elements[rightCombo].options[i].value;
	}	
	
	document.forms[0].elements[hiddenFieldName].value=par;
	document.forms[0].submit();
	return true;
}

function  selectTheItem(obj,srcCombo)
{

		objSel = document.forms[0].elements[srcCombo];	

		var arr = new Array(objSel.length);
		str2 = obj.value;
				
		for(i=0; i < objSel.options.length; i++)
		{
				objSel.options[i].selected = false;
		}
		
		if (obj.value == "") 
		{
			objSel.selectedIndex = -1;
		}

		for (i=0;i<arr.length;i++)
		{
			arr[i] = objSel.options[i].text;
		}

		if(str2!="")
		{
			str2=str2.toUpperCase();
			for(i=0; i < arr.length; i++)
			{
				str = arr[i].toUpperCase();
				if (str.indexOf(str2)==0)
				{
					
					objSel.options[i].selected = true;
					break;
				}
			}
		}
}
