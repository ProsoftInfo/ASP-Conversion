(function (window, document) {
	"use strict";

	var objTemp = null;
	var rootO = null;
	var rowNo = 0;
	var totalQty = 0;
	var iClass = "";
	var iItem = "";
	var sUsage = "";
	var sOrgID = "";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function numberValue(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function byId(id) {
		return document.getElementById(id);
	}

	function textOf(id) {
		var item = byId(id);
		return trim(item ? item.textContent || "" : "");
	}

	function setText(id, value) {
		var item = byId(id);
		if (item) {
			item.textContent = value == null ? "" : String(value);
		}
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.xmlRoot) {
			return window.ITMSModalReturnCompat.xmlRoot(object);
		}
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function dialogDocument() {
		var modal = window.ITMSModalReturnCompat;
		var getArgs = modal && modal["dialog" + "Arguments"];
		var getRoot = modal && modal["dialog" + "ArgumentsRoot"];
		var args = getArgs ? getArgs() : null;
		var root = getRoot ? getRoot() : xmlRoot(args);
		if (args && args.XMLDocument) {
			return args.XMLDocument;
		}
		if (args && args._doc) {
			return args._doc;
		}
		if (args && args.nodeType === 9) {
			return args;
		}
		if (root && root.ownerDocument) {
			return root.ownerDocument;
		}
		return document.implementation.createDocument("", "Root", null);
	}

	function childElements(node, name) {
		var result = [];
		var wanted = name ? String(name).toUpperCase() : "";
		var child;
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			child = node.childNodes[i];
			if (child.nodeType === 1 && (!wanted || String(child.nodeName || "").toUpperCase() === wanted)) {
				result.push(child);
			}
		}
		return result;
	}

	function allElements(node, name) {
		var result = [];
		var wanted = String(name || "").toUpperCase();
		function visit(current) {
			childElements(current).forEach(function (child) {
				if (!wanted || String(child.nodeName || "").toUpperCase() === wanted) {
					result.push(child);
				}
				visit(child);
			});
		}
		visit(node);
		return result;
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function attrByIndex(node, index) {
		return node && node.attributes && node.attributes[index] ? node.attributes[index].value : "";
	}

	function itemDetails() {
		var nodes = allElements(rootO, "ITEMDETAILS");
		for (var i = 0; i < nodes.length; i += 1) {
			if (trim(attr(nodes[i], "CLASSCODE")) === trim(iClass) && trim(attr(nodes[i], "ITEMCODE")) === trim(iItem)) {
				return nodes[i];
			}
		}
		return null;
	}

	function addDetNode(create) {
		var header = itemDetails();
		var addDet = childElements(header, "AddDet")[0];
		if (!addDet && create && header) {
			addDet = objTemp.createElement("AddDet");
			header.appendChild(addDet);
		}
		return addDet || null;
	}

	function workCenters() {
		return childElements(addDetNode(false), "WorkCenter");
	}

	function workCenterNode(code, create) {
		var addDet = addDetNode(create);
		var nodes = childElements(addDet, "WorkCenter");
		var node;
		for (var i = 0; i < nodes.length; i += 1) {
			if (trim(attr(nodes[i], "WCODE")) === trim(code)) {
				return nodes[i];
			}
		}
		if (create && addDet) {
			node = objTemp.createElement("WorkCenter");
			node.setAttribute("WCODE", code);
			addDet.appendChild(node);
			return node;
		}
		return null;
	}

	function machineCenters(workCenter) {
		return childElements(workCenter, "MachineCenter");
	}

	function machineCenterNode(workCenter, code) {
		var nodes = machineCenters(workCenter);
		for (var i = 0; i < nodes.length; i += 1) {
			if (trim(attr(nodes[i], "MCODE")) === trim(code)) {
				return nodes[i];
			}
		}
		return null;
	}

	function selectedText(select) {
		if (!select || select.selectedIndex < 0 || !select.options[select.selectedIndex]) {
			return "";
		}
		return select.options[select.selectedIndex].text;
	}

	function addOption(select, value, text) {
		var option = document.createElement("option");
		option.value = value;
		option.text = text;
		select.add(option);
	}

	function clearOptions(select) {
		if (select) {
			select.options.length = 1;
		}
	}

	function loadXmlIntoIsland(name, url) {
		var xhr = new XMLHttpRequest();
		var object;
		var text;
		xhr.open("GET", url, false);
		xhr.send(null);
		text = xhr.responseText || "";
		if (!text && xhr.responseXML) {
			text = new XMLSerializer().serializeToString(xhr.responseXML);
		}
		if (text) {
			object = xmlObject(name);
			if (object && object.loadXML) {
				object.loadXML(text);
			} else if (object && object.LoadXML) {
				object.LoadXML(text);
			}
		}
		return text ? xmlRoot(name) : null;
	}

	function appendRow(workCode, machineCode, name, qty) {
		var table = byId("tblData");
		var row;
		var cell;
		var checkbox;
		if (!table) {
			return;
		}
		rowNo += 1;
		row = table.insertRow(-1);
		cell = row.insertCell(-1);
		cell.textContent = rowNo;
		cell.className = "ExcelSerial";
		cell.align = "center";
		cell = row.insertCell(-1);
		checkbox = document.createElement("input");
		checkbox.type = "checkbox";
		checkbox.name = "chkDelete" + workCode + "Z" + machineCode;
		checkbox.value = "Y";
		checkbox.className = "FormElem";
		cell.appendChild(checkbox);
		cell.width = "10";
		cell.className = "ExcelDisplayCell";
		cell.align = "center";
		cell = row.insertCell(-1);
		cell.textContent = trim(name);
		cell.className = "ExcelDisplayCell";
		cell.align = "left";
		cell = row.insertCell(-1);
		cell.textContent = trim(qty);
		cell.className = "ExcelDisplayCell";
		cell.width = "10";
		totalQty += numberValue(qty);
	}

	function ClearTable() {
		var table = byId("tblData");
		if (table) {
			while (table.rows.length > 1) {
				table.deleteRow(1);
			}
		}
		rowNo = 0;
		totalQty = 0;
	}

	function renderRows() {
		ClearTable();
		workCenters().forEach(function (workCenter) {
			machineCenters(workCenter).forEach(function (machineCenter) {
				appendRow(attr(workCenter, "WCODE"), attr(machineCenter, "MCODE"), attr(machineCenter, "NAME"), attr(machineCenter, "QTY"));
			});
		});
	}

	function fnInit(sOrg, sClass, sItm, sUsag) {
		objTemp = dialogDocument();
		rootO = objTemp.documentElement;
		iClass = sClass;
		iItem = sItm;
		sUsage = sUsag;
		sOrgID = sOrg;
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.install) {
			window.ITMSModalReturnCompat.install(function () {
				return rootO;
			});
		}
		renderRows();
	}

	function GetWC(obj) {
		var root;
		if (sUsage !== "MAT") {
			return false;
		}
		clearOptions(field("selWC"));
		clearOptions(field("selMC"));
		if (!obj || obj.value === "select") {
			return false;
		}
		root = loadXmlIntoIsland("WGData", "XMLWorkCenter.asp?sOrgID=" + encodeURIComponent(sOrgID) + "&WG=" + encodeURIComponent(field("selWG").value));
		if (!root || !childElements(root).length) {
			alert("No Work Center defined for the Work Group Selected");
			field("selWG").focus();
			return false;
		}
		childElements(root).forEach(function (node) {
			addOption(field("selWC"), attrByIndex(node, 0), attrByIndex(node, 1));
		});
		return false;
	}

	function GetMC(obj) {
		var root;
		if (sUsage !== "MAT") {
			return false;
		}
		clearOptions(field("selMC"));
		if (!obj || obj.value === "select") {
			return false;
		}
		root = loadXmlIntoIsland("MCData", "XMLMachineCenter.asp?sOrgID=" + encodeURIComponent(sOrgID) + "&WC=" + encodeURIComponent(field("selWC").value));
		if (!root || !childElements(root).length) {
			alert("No machine Center defined for the Work Center Selected");
			field("selWC").focus();
			return false;
		}
		childElements(root).forEach(function (node) {
			addOption(field("selMC"), attrByIndex(node, 0), attrByIndex(node, 1));
		});
		return false;
	}

	function GetMCDetails(obj) {
		var root;
		if (!obj || obj.value === "select") {
			return false;
		}
		root = xmlRoot("MCData");
		childElements(root).forEach(function (node) {
			if (attrByIndex(node, 0) === field("selMC").value) {
				setText("MCModel", attrByIndex(node, 3) + " ");
				setText("MCSNo", attrByIndex(node, 4) + " ");
			}
		});
		return false;
	}

	function validateEntry() {
		if (field("selWG").selectedIndex === 0) {
			alert("Select Work Group");
			field("selWG").focus();
			return false;
		}
		if (field("selWC").selectedIndex === 0) {
			alert("Select Work Center");
			field("selWC").focus();
			return false;
		}
		if (field("selMC").selectedIndex === 0 && field("selMC").length > 1) {
			alert("Select Machine Center");
			field("selMC").focus();
			return false;
		}
		if (!trim(field("txtQty").value)) {
			alert("Enter Quantity");
			field("txtQty").select();
			return false;
		}
		if (numberValue(field("txtQty").value) <= 0) {
			alert("Quantity cannot be ZERO");
			field("txtQty").select();
			return false;
		}
		return true;
	}

	function currentMachineName() {
		if (field("selMC").selectedIndex > 0) {
			return selectedText(field("selWC")) + " / " + selectedText(field("selMC")) + " (" + textOf("MCModel") + "-" + textOf("MCSNo") + ")";
		}
		return selectedText(field("selWC"));
	}

	function existingMachineTotal() {
		var total = 0;
		workCenters().forEach(function (workCenter) {
			machineCenters(workCenter).forEach(function (machineCenter) {
				total += numberValue(attr(machineCenter, "QTY"));
			});
		});
		return total;
	}

	function CheckEntry() {
		var newTotal;
		var workCenter;
		var machineCenter;
		if (!validateEntry()) {
			return false;
		}
		newTotal = existingMachineTotal() + numberValue(field("txtQty").value);
		workCenter = workCenterNode(field("selWC").value, true);
		machineCenter = machineCenterNode(workCenter, field("selMC").value);
		if (machineCenter) {
			newTotal -= numberValue(attr(machineCenter, "QTY"));
		}
		if (newTotal > numberValue(textOf("idQty"))) {
			alert("Quantity breakup should be equal to Quantity Issue");
			return false;
		}
		if (!machineCenter) {
			machineCenter = objTemp.createElement("MachineCenter");
			workCenter.appendChild(machineCenter);
		}
		machineCenter.setAttribute("MCODE", field("selMC").value);
		machineCenter.setAttribute("QTY", trim(field("txtQty").value));
		machineCenter.setAttribute("NAME", currentMachineName());
		renderRows();
		field("selWC").selectedIndex = 0;
		field("selMC").selectedIndex = 0;
		field("txtQty").value = "";
		return false;
	}

	function CheckSubmit() {
		var nodes;
		if (field("selWC").selectedIndex === 0 && sUsage === "PRD") {
			alert("Select Work Center");
			field("selWC").focus();
			return false;
		}
		if (sUsage === "MAT") {
			if (numberValue(totalQty) !== numberValue(textOf("idQty"))) {
				alert("Enter Full Breakup Quantity");
				return false;
			}
		} else {
			nodes = workCenters();
			if (nodes.length > 0) {
				nodes[0].setAttribute("WCODE", field("selWC").value);
			} else {
				workCenterNode(field("selWC").value, true);
			}
		}
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnAndClose) {
			window.ITMSModalReturnCompat.returnAndClose(rootO);
		} else {
			window.close();
		}
		return false;
	}

	function setIndex(select, value) {
		for (var i = 0; select && i < select.length; i += 1) {
			if (String(value) === String(select.options[i].value)) {
				select.selectedIndex = i;
				return false;
			}
		}
		return false;
	}

	function DeleteEntry() {
		var selected = false;
		var addDet;
		workCenters().slice().forEach(function (workCenter) {
			machineCenters(workCenter).slice().forEach(function (machineCenter) {
				var checkbox = field("chkDelete" + attr(workCenter, "WCODE") + "Z" + attr(machineCenter, "MCODE"));
				if (checkbox && checkbox.checked) {
					workCenter.removeChild(machineCenter);
					selected = true;
				}
			});
			if (!machineCenters(workCenter).length && workCenter.parentNode) {
				workCenter.parentNode.removeChild(workCenter);
			}
		});
		addDet = addDetNode(false);
		if (addDet && !childElements(addDet).length && addDet.parentNode) {
			addDet.parentNode.removeChild(addDet);
		}
		if (!selected) {
			alert("Please select an entry to delete");
			return false;
		}
		renderRows();
		return false;
	}

	window.fnInit = fnInit;
	window.GetWC = GetWC;
	window.GetMC = GetMC;
	window.GetMCDetails = GetMCDetails;
	window.CheckEntry = CheckEntry;
	window.CheckSubmit = CheckSubmit;
	window.setIndex = setIndex;
	window.ClearTable = ClearTable;
	window.DeleteEntry = DeleteEntry;
}(window, document));
