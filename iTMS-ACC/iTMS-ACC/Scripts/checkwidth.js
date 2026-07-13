cc = 0

function itmsElement(id) {
	return document.getElementById(id) || window[id];
}

function itmsFrameDocument(index) {
	var frame = window.frames && window.frames[index];
	return frame && frame.document ? frame.document : document;
}

function Home() {
	if (cc==0) {
		hideMenu()
		cc = 1
	}
	else {
		showMenu()
		cc = 0
	}
}

function showMenu() {
	var i,divlen
	var tblMenuHead = itmsElement("tblMenuHead");
	var tblBody = itmsElement("tblBody");
	var Menu = itmsElement("Menu");
	var frameDocument = itmsFrameDocument(1);
	var aReturn= frameDocument.body.getElementsByTagName("DIV");

//-----------------
	var divarr, divwidth, divid, temp;
	divarr = frameDocument.body.getElementsByTagName('DIV');
	divarrlength = divarr.length;

/*	for(i=0;i<divarrlength;i++){
		if(divarr[i].id != ""){
			//alert( "index is " + i + " "+ divarr[i].id + " - " + divarr[i].style.width );

			temp=divarr[i].style.width;
			// width of the div	alert(temp.substring(0,temp.length-2));
			divwidth = temp.substring(0,temp.length-2);
			//alert(divwidth);

			if(divwidth > 585)
			{

				for (i=0;i<aReturn.length;i++)
				{
					if (aReturn[i].id!=null && aReturn[i].id!="")
					{
						divlen = new String(aReturn[i].style.width).replace('px','');
						divlen = Math.round((parseInt(divlen) * 0.809));
						aReturn[i].style.width=divlen+'px';
					}
				}
			}
			else if(divwidth < 574)
			{
				for (i=0;i<aReturn.length;i++)
				{
					if (aReturn[i].id!=null && aReturn[i].id!="")
					{
						divlen = new String(aReturn[i].style.width).replace('px','');
						divlen = Math.round((parseInt(divlen) * 0.809));
						aReturn[i].style.width=divlen+'px';
					}
				}
			}
			else
			{
				for (i=0;i<aReturn.length;i++)
				{
					if (aReturn[i].id!=null && aReturn[i].id!="")
					{
						divlen = new String(aReturn[i].style.width).replace('px','');
						divlen = Math.round((parseInt(divlen) * 1.0));
						aReturn[i].style.width=divlen+'px';
					}
				}
			}
		}
	}*/
//-----------------

	tblMenuHead.deleteRow(0);
	oRow = tblMenuHead.insertRow(0);
	headerCell=oRow.insertCell();

	headerCell.innerHTML="&nbsp;Menu";
	headerCell.className="NavTitleText";
	headerCell.width="50%";

	headerCell=oRow.insertCell();
	headerCell.className="NavTitleImg";
	headerCell.width="50%";
	headerCell.align="right"
	headerCell.innerHTML="<span style=\"cursor: pointer\"><IMG id=\"imgEC\" onclick=\"Home()\" Title=Collapse src=\"../assets/images/CollapseButton.gif\"  border=2 width=\"17\" height=\"14\" style=\"border-style: solid; border-color: #999999; background-color: #ffffff;\"></span>";

	tblMenuHead.rows[1].cells[0].bgColor = "#ffffff"

	//alert(tblBody.rows[0].width);
	//alert("Hello");

	tblBody.rows[0].cells[0].width="20%";
	tblBody.rows[0].cells[1].width="80%";



	Menu.style.visibility="visible";
}

function hideMenu() {
	var i,divlen
	var tblMenuHead = itmsElement("tblMenuHead");
	var tblBody = itmsElement("tblBody");
	var Menu = itmsElement("Menu");
	var aReturn= itmsFrameDocument(1).body.getElementsByTagName("DIV");
	
/*	for (i=0;i<aReturn.length;i++)
	{
	    if (aReturn[i].id!=null && aReturn[i].id!="")
		{
			divlen = new String(aReturn[i].style.width).replace('px','');
			divlen = Math.round((parseInt(divlen) /0.809));
			alert("113");
			aReturn[i].style.width=divlen+'px';
		}
	}*/
	tblMenuHead.deleteRow(0);
	oRow = tblMenuHead.insertRow(0);
	headerCell=oRow.insertCell();
	headerCell.height=10;
	headerCell.width=20;
	headerCell.colSpan=2;

	tblMenuHead.rows[1].cells[0].width=20;
	tblMenuHead.rows[1].cells[0].bgColor = "#cccccc"

	tblBody.rows[0].cells[0].width="10";
	tblBody.rows[0].cells[1].width="100%";

	headerCell.innerHTML="<span style=\"cursor: pointer\"><IMG id=\"imgEC\" onclick=\"Home()\" Title=Expand src=\"../assets/images/ExpandButton.gif\"  border=2 width=\"17\" height=\"14\" style=\"border-style: solid; border-color: #999999; background-color: #ffffff;\"></span>";

	Menu.style.visibility="hidden";
}



// The following code checks whether the left side menu is collapsed and adjusts the width
// of the target page
function checkWidth(){
var parentflag = cc;
var frm1 = document.getElementById("frm1");
if (!frm1) {
	return;
}
if(parentflag == 0)
	{
		frm1.style.width="603px";
		frm1.style.height="377px";
	}
	else
	{
		frm1.style.width="739px";
		
	}
}
