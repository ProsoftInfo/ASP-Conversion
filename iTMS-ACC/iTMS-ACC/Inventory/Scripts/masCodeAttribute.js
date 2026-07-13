function ViewAttr(sAddDesc)
{
    var oPopup = window.createPopup();
    var oPopBody = oPopup.document.body;
	oPopBody.style.backgroundColor = "lightyellow";
	oPopBody.style.border = "solid black 1px";
    oPopBody.innerHTML = sAddDesc;
    oPopup.show(400, 75, 200, 75, document.body);
}
