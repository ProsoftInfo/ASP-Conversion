function  selectTheItem(obj,srcCombo){
var i;		
		objSel = document.forms[0].elements[srcCombo];	
		i = 0;
		if (obj.value == "") {
			for(i=0; i < objSel.options.length; i++){
				objSel.options[i].selected = false;
			}
		}
		i = 0;
		for(i=0; i < objSel.options.length; i++){
			if (obj.value != "" && objSel.options[i].text.toUpperCase().indexOf(obj.value.toUpperCase()) >=0 ){
				objSel.options[i].selected = true;
				return;
			}
		}
		if (obj.value == "") {
			objSel.selectedIndex = -1;
		}
}
