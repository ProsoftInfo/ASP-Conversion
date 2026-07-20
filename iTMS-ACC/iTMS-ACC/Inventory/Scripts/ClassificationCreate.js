function treeFrame() {
	return window.frames["main"] || window.frames[0] || null;
}

function treeControl() {
	var frame = treeFrame();
	return frame && (frame.ctlClassificationTree || frame.document && frame.document.getElementById("ctlClassificationTree"));
}

function trimValue(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function selectedTreeKey() {
	var tree = treeControl();
	return trimValue(tree && (tree.GetKey || tree.value) || "");
}

function selectedTreeText() {
	var tree = treeControl();
	return trimValue(tree && tree.GetText || "");
}

function selectedTreePath() {
	var tree = treeControl();
	return trimValue(tree && tree.GetFullPath || "");
}

function normalizeClassificationGroupKey(gKey, allowRoot) {
	var key = trimValue(gKey);
	var parts;
	var last;
	if (key === "" || key === "0") {
		return allowRoot ? "GRP" : "";
	}
	if (key === "GRP" || /^CAT/i.test(key)) {
		return allowRoot ? "GRP" : "";
	}
	parts = key.split(":");
	last = trimValue(parts[parts.length - 1]);
	if (last === "" || last === "GRP" || /^CAT/i.test(last)) {
		return allowRoot ? "GRP" : "";
	}
	if (parts.length >= 4) {
		return "C" + trimValue(parts[0]) + ":" + last;
	}
	return "C" + last + ":" + last;
}

function classificationCodeFromKey(gKey) {
	var normalized = normalizeClassificationGroupKey(gKey, false);
	var parts;
	if (normalized === "") {
		return "";
	}
	parts = normalized.split(":");
	return trimValue(parts[parts.length - 1]);
}

function firstItemCode(xmlText) {
	var doc;
	var item;
	if (trimValue(xmlText) === "") {
		return "";
	}
	doc = new DOMParser().parseFromString(xmlText, "text/xml");
	if (doc.getElementsByTagName("parsererror").length) {
		return "";
	}
	item = doc.getElementsByTagName("Item")[0];
	return item ? trimValue(item.getAttribute("Code")) : "";
}

function Refresh() {
	var tree = treeControl();
	if (tree && tree.populateTree) {
		tree.populateTree();
	} else if (tree && tree.Refresh) {
		tree.Refresh();
	}
}

function GetAttributes(gName, gPath, gKey) {
	var normalized = normalizeClassificationGroupKey(gKey, false);
	if (normalized === "") {
		alert("Select Classification Created");
		document.formname.pGroup.value = "";
		document.formname.action = "MasClassificationAttributeEntry.asp";
		document.formname.submit();
		return false;
	}
	document.formname.pGroup.value = normalized;
	document.formname.pName.value = gName;
	document.formname.gPath.value = gPath;
	document.formname.action = "MasClassificationAttributeEntry.asp";
	document.formname.submit();
	return true;
}

function Attributes() {
	GetAttributes(selectedTreeText(), selectedTreePath(), selectedTreeKey());
}

function NewGroupValidate(gKey) {
	var normalized = normalizeClassificationGroupKey(gKey, true);
	var groupCode;
	var request;
	var itemCode;
	if (normalized === "GRP") {
		document.formname.pGroup.value = normalized;
		document.formname.action = "MasClassificationNameEntry.asp";
		document.formname.submit();
		return true;
	}
	groupCode = classificationCodeFromKey(normalized);
	if (groupCode === "") {
		alert("Cannot create Classification from here. Select Classification or Existing Classification.");
		return false;
	}
	request = new XMLHttpRequest();
	request.open("GET", "../components/NewGroupVal.asp?Groupcode=" + encodeURIComponent(groupCode), false);
	request.send(null);
	itemCode = firstItemCode(request.responseText);
	if (itemCode === "") {
		alert("No Data");
		return false;
	}
	if (itemCode === "-1") {
		document.formname.pGroup.value = normalized;
		document.formname.action = "MasClassificationNameEntry.asp";
		document.formname.submit();
	} else {
		alert("Classification cannot be created since Item has been Created under this Classification");
		document.formname.pGroup.value = "";
		document.formname.action = "MasClassificationNameEntry.asp";
		document.formname.submit();
	}
	return true;
}

function NewGroup() {
	NewGroupValidate(selectedTreeKey() || "GRP");
}
