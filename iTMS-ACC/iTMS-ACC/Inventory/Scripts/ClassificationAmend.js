function treeFrame() {
	return window.frames["main"] || window.frames[0] || null;
}

function treeControl() {
	var frame = treeFrame();
	return frame && (frame.ctlClassificationTree || frame.document && frame.document.getElementById("ctlClassificationTree"));
}

function Refresh() {
	var tree = treeControl();
	if (tree && tree.populateTree) {
		tree.populateTree();
	}
}

function Amend(sVal) {
	var tree = treeControl();
	var gName = tree ? tree.GetText : "";
	var gPath = tree ? tree.GetFullPath : "";
	var gKey = tree ? tree.GetKey : "";
	AmendClass(gName, gPath, gKey, sVal);
}

function Delete(sStr) {
	var tree = treeControl();
	var gKey = tree ? tree.GetKey : "";
	var sOrgID;
	if (sStr == "I") {
		sOrgID = document.forms[0].selOrgUnit.value;
		DeleteItem(gKey, sOrgID);
	}
	else if (sStr == "C") {
		DeleteClass(gKey);
	}
}

function AmendClass(gName, gPath, gKey, sVal) {
	var normalized = window.normalizeClassificationGroupKey ? window.normalizeClassificationGroupKey(gKey, false) : String(gKey || "");
	var sTemp;
	var gArr = normalized.split(":");
	if (normalized === "" || normalized === "GRP" || gArr.length < 2) {
		alert("Select Classification Created");
		return false;
	}
	sTemp = normalized;
	document.formname.pGroup.value = sTemp;
	document.formname.pName.value = gName;
	document.formname.hPara.value = sVal;
	document.formname.action = "MasClassificationNameAmendEntry.asp";
	document.formname.submit();
	return true;
}

function DeleteItem(gKey, sOrgID) {
	if (sOrgID === "select") {
		alert("Select Organization");
		return false;
	}
	if (gKey === "GRP") {
		alert("Select an Category or an Classification");
		return false;
	}
	document.formname.pGroup.value = gKey;
	document.formname.action = "ItmDeletionDetailsEntry.asp";
	document.formname.submit();
	return true;
}

function DeleteClass(gKey) {
	var normalized = window.normalizeClassificationGroupKey ? window.normalizeClassificationGroupKey(gKey, false) : String(gKey || "");
	var gArr = normalized.split(":");
	if (normalized === "" || normalized === "GRP" || gArr.length < 2) {
		alert("Select Classification Created");
		return false;
	}
	document.formname.pGroup.value = normalized;
	document.formname.action = "ClassDeletionDetailsEntry.asp";
	document.formname.submit();
	return true;
}
