function GetParentWithTag( oStart, sTag ) {
	var oTag = oStart;
	while( oTag && (oTag.tagName.toUpperCase() != sTag) && (oTag != oTag.parentElement))
		oTag = oTag.parentElement;
	return oTag;
}

function itmsCancelDivEvent(evt) {
	if (evt && evt.preventDefault) {
		evt.preventDefault();
	}
	if (evt && evt.stopPropagation) {
		evt.stopPropagation();
	}
}

function Div_OnClick(objDiv, evt) {
	var divarr, divwidth, divid, temp;
	divarr = document.body.getElementsByTagName('IMG');
	divarrlength = divarr.length;

	itmsCancelDivEvent(evt);
	var oAnchor = GetParentWithTag((evt && evt.target) || objDiv, 'A' );
	if (!oAnchor) {
		return false;
	}
	var oImage  = oAnchor.children[0];

	for(i=0;i<divarrlength;i++){
		if(divarr[i].style.width != "") {
			oOthAnchor = divarr[i].parentElement;
			//alert(oOthAnchor.title);
			if(oOthAnchor.title != oAnchor.title) {
				//alert(oOthAnchor.title+'>>>'+oAnchor.title);
				oOthImage  = oOthAnchor.children[0];
				oOthImage.src = '../../assets/images/plus.gif';
				oOthAnchor.iTMS_State = '0';
				oOthImage.alt = "Expands this section.";
			}
		}
	}

	divarr = document.body.getElementsByTagName('DIV');
	divarrlength = divarr.length;

	for(i=0;i<divarrlength;i++){
		//alert(divarr[i].id+">>>"+oAnchor.id);
		if(divarr[i].id != "") {
			divarr[i].style.display = 'none';
		}
	}

	if (oAnchor.iTMS_State == '1') {
		oImage.src = '../../assets/images/plus.gif';
		oAnchor.iTMS_State = '0';
		objDiv.style.display = 'none';
		oImage.alt = "Expands this section.";
	}
	else {
		oImage.src = '../../assets/images/minus.gif';
		oAnchor.iTMS_State = '1';
		objDiv.style.display = 'block';
		oImage.alt = "Collapses this section.";
	}
	return false;
}
