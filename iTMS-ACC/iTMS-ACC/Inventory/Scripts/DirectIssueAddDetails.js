(function (window, document) {
	"use strict";

	var objTemp = null;
	var root = null;
	var rootDoc = null;
	var rowCount = 0;
	var classCode = "";
	var itemCode = "";
	var usage = "";
	var orgId = "";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function byId(id) {
		return document.getElementById(id) || window[id] || null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function textOf(id) {
		var element = byId(id);
		return trim(element ? element.innerText || element.textContent || "" : "");
	}

	function setText(id, value) {
		var element = byId(id);
		if (element) {
			element.textContent = String(value == null ? "" : value);
			element.innerText = String(value == null ? "" : value);
		}
	}

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(String(value || ""));
	}

	function xmlDocument(value) {
		ensureCompat();
		if (!value) {
			return null;
		}
		if (value.XMLDocument) {
			return value.XMLDocument;
		}
		if (value._doc) {
			return value._doc;
		}
		if (value.nodeType === 9) {
			return value;
		}
		return value.ownerDocument || null;
	}

	function xmlRoot(value) {
		if (!value) {
			return null;
		}
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || value;
	}

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
	}

	function xmlIsland(name) {
		ensureCompat();
		return window[name] || document[name] || byId(name);
	}

	function elementChildren(node, name) {
		var result = [];
		var wanted = name && String(name).toLowerCase();
		var children = node && node.childNodes || [];
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1 && (!wanted || String(children[i].nodeName).toLowerCase() === wanted)) {
				result.push(children[i]);
			}
		}
		return result;
	}

	function attr(node, name) {
		return node && node.getAttribute ? trim(node.getAttribute(name)) : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, String(value == null ? "" : value));
		}
	}

	function attrAt(node, index) {
		var item = node && node.attributes && node.attributes.item(index);
		return trim(item ? item.nodeValue || item.value || "" : "");
	}

	function parseXml(text) {
		if (!trim(text)) {
			return null;
		}
		return new window.DOMParser().parseFromString(text, "application/xml");
	}

	function responseRoot(xhr) {
		var doc = xhr.responseXML && xhr.responseXML.documentElement ? xhr.responseXML : parseXml(xhr.responseText);
		return doc && doc.documentElement ? doc.documentElement : null;
	}

	function loadXmlIsland(name, text) {
		var island = xmlIsland(name);
		if (island && typeof island.loadXML === "function") {
			island.loadXML(text);
			return xmlRoot(island);
		}
		return responseRoot({ responseText: text, responseXML: null });
	}

	function selectText(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function selectedUsageRequiresMachineBreakup() {
		return usage === "MAT" || usage === "CON";
	}

	function findItemDetails() {
		var items = elementChildren(root, "ITEMDETAILS");
		for (var i = 0; i < items.length; i += 1) {
			if (attr(items[i], "CLASSCODE") === trim(classCode) && attr(items[i], "ITEMCODE") === trim(itemCode)) {
				return items[i];
			}
		}
		return null;
	}

	function addDetNode(itemNode, createIfMissing) {
		var addDet = elementChildren(itemNode, "AddDet")[0];
		if (!addDet && createIfMissing) {
			addDet = rootDoc.createElement("AddDet");
			itemNode.appendChild(addDet);
		}
		return addDet || null;
	}

	function workCenters(itemNode) {
		var addDet = addDetNode(itemNode, false);
		return elementChildren(addDet, "WorkCenter");
	}

	function findWorkCenter(itemNode, workCode) {
		var nodes = workCenters(itemNode);
		for (var i = 0; i < nodes.length; i += 1) {
			if (attr(nodes[i], "WCODE") === trim(workCode)) {
				return nodes[i];
			}
		}
		return null;
	}

	function machineCenters(workCenterNode) {
		return elementChildren(workCenterNode, "MachineCenter");
	}

	function findMachineCenter(workCenterNode, machineCode) {
		var nodes = machineCenters(workCenterNode);
		for (var i = 0; i < nodes.length; i += 1) {
			if (attr(nodes[i], "MCODE") === trim(machineCode)) {
				return nodes[i];
			}
		}
		return null;
	}

	function currentMachineTotal() {
		var itemNode = findItemDetails();
		var total = 0;
		var wcs = workCenters(itemNode);
		var mcs;
		for (var i = 0; i < wcs.length; i += 1) {
			mcs = machineCenters(wcs[i]);
			for (var m = 0; m < mcs.length; m += 1) {
				total += toNumber(attr(mcs[m], "QTY"));
			}
		}
		return total;
	}

	function projectedMachineTotal(workCode, machineCode, newQty) {
		var itemNode = findItemDetails();
		var workCenter = itemNode ? findWorkCenter(itemNode, workCode) : null;
		var machineCenter = workCenter ? findMachineCenter(workCenter, machineCode) : null;
		return currentMachineTotal() - toNumber(attr(machineCenter, "QTY")) + toNumber(newQty);
	}

	function makeCheckbox(name) {
		var input = document.createElement("input");
		input.type = "checkbox";
		input.name = name;
		input.value = "Y";
		input.className = "FormElem";
		return input;
	}

	function appendCell(row, className, align, content) {
		var cell = row.insertCell(row.cells.length);
		cell.className = className;
		if (align) {
			cell.align = align;
		}
		if (typeof content === "string" || typeof content === "number") {
			cell.textContent = String(content);
		} else if (content) {
			cell.appendChild(content);
		}
		return cell;
	}

	function appendMachineRow(workCode, machineCode, name, quantity) {
		var table = byId("tblData");
		var row = table.insertRow(table.rows.length);
		rowCount += 1;
		appendCell(row, "ExcelSerial", "center", rowCount);
		appendCell(row, "ExcelDisplayCell", "center", makeCheckbox("chkDelete" + workCode + "Z" + machineCode)).width = "10";
		appendCell(row, "ExcelDisplayCell", "left", name);
		appendCell(row, "ExcelDisplayCell", null, quantity).width = "10";
	}

	function renderTable() {
		var itemNode = findItemDetails();
		var wcs = workCenters(itemNode);
		var mcs;
		window.ClearTable();
		for (var i = 0; i < wcs.length; i += 1) {
			mcs = machineCenters(wcs[i]);
			for (var m = 0; m < mcs.length; m += 1) {
				appendMachineRow(attr(wcs[i], "WCODE"), attr(mcs[m], "MCODE"), attr(mcs[m], "NAME"), attr(mcs[m], "QTY"));
			}
		}
	}

	function setReturnValue() {
		if (!root) {
			root = xmlRoot(objTemp);
		}
		if (!root) {
			return;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		} else {
			window["return" + "Value"] = root;
			window.returnvalue = root;
		}
	}

	function closeWithReturn() {
		setReturnValue();
		window.close();
	}

	function populateSelectFromXml(select, xmlRootNode, emptyMessage, focusTarget) {
		var nodes = elementChildren(xmlRootNode);
		select.options.length = 1;
		if (!nodes.length) {
			alert(emptyMessage);
			if (focusTarget && focusTarget.focus) {
				focusTarget.focus();
			}
			return false;
		}
		for (var i = 0; i < nodes.length; i += 1) {
			select.options[select.options.length] = new Option(attrAt(nodes[i], 1), attrAt(nodes[i], 0));
		}
		return true;
	}

	function getXml(url, islandName, callback) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		if (!(xhr.status === 0 || xhr.status >= 200 && xhr.status < 300) || !trim(xhr.responseText)) {
			callback(null);
			return;
		}
		callback(loadXmlIsland(islandName, xhr.responseText));
	}

	function machineDisplayName() {
		var wc = field("selWC");
		var mc = field("selMC");
		if (mc && mc.selectedIndex > 0) {
			return selectText(wc) + " / " + selectText(mc) + " (" + textOf("MCModel") + "-" + textOf("MCSNo") + ")";
		}
		return selectText(wc);
	}

	function resetEntryFields() {
		if (field("selWC")) {
			field("selWC").selectedIndex = 0;
		}
		if (field("selMC")) {
			field("selMC").selectedIndex = 0;
		}
		if (field("txtQty")) {
			field("txtQty").value = "";
		}
		setText("MCModel", "");
		setText("MCSNo", "");
	}

	window.fnInit = function (sOrg, sClass, sItm, sUsag) {
		objTemp = modalArgs();
		root = xmlRoot(objTemp);
		rootDoc = xmlDocument(objTemp);
		orgId = sOrg;
		classCode = sClass;
		itemCode = sItm;
		usage = sUsag;
		if (!rootDoc && root) {
			rootDoc = root.ownerDocument;
		}
		renderTable();
	};

	window.GetWC = function (obj) {
		var wc = field("selWC");
		var workGroup = field("selWG");
		if (!selectedUsageRequiresMachineBreakup()) {
			return;
		}
		if (wc) {
			wc.options.length = 1;
		}
		if (!obj || obj.value === "select") {
			return;
		}
		getXml("XMLWorkCenter.asp?sOrgID=" + encodeURIComponent(orgId) + "&WG=" + encodeURIComponent(workGroup ? workGroup.value : ""), "WGData", function (xmlRootNode) {
			if (wc) {
				populateSelectFromXml(wc, xmlRootNode, "No Work Center defined for the Work Group Selected", workGroup);
			}
		});
	};

	window.GetMC = function (obj) {
		var mc = field("selMC");
		var wc = field("selWC");
		if (!selectedUsageRequiresMachineBreakup()) {
			return;
		}
		if (mc) {
			mc.options.length = 1;
		}
		if (!obj || obj.value === "select") {
			return;
		}
		getXml("XMLMachineCenter.asp?sOrgID=" + encodeURIComponent(orgId) + "&WC=" + encodeURIComponent(wc ? wc.value : ""), "MCData", function (xmlRootNode) {
			if (mc) {
				populateSelectFromXml(mc, xmlRootNode, "No machine Center defined for the Work Center Selected", wc);
			}
		});
	};

	window.GetMCDetails = function (obj) {
		var rootNode = xmlRoot(xmlIsland("MCData"));
		var nodes = elementChildren(rootNode);
		if (!obj || obj.value === "select") {
			setText("MCModel", "");
			setText("MCSNo", "");
			return;
		}
		for (var i = 0; i < nodes.length; i += 1) {
			if (attrAt(nodes[i], 0) === obj.value) {
				setText("MCModel", attrAt(nodes[i], 3) + "\u00a0");
				setText("MCSNo", attrAt(nodes[i], 4) + "\u00a0");
				return;
			}
		}
	};

	window.CheckEntry = function () {
		var itemNode = findItemDetails();
		var addDet;
		var workCenter;
		var machineCenter;
		var wc = field("selWC");
		var mc = field("selMC");
		var qty = field("txtQty");
		var total;

		if (field("selWG") && field("selWG").selectedIndex === 0) {
			alert("Select Work Group");
			field("selWG").focus();
			return;
		}
		if (wc && wc.selectedIndex === 0) {
			alert("Select Work Center");
			wc.focus();
			return;
		}
		if (mc && mc.selectedIndex === 0 && mc.length > 1 && usage !== "CON") {
			alert("Select Machine Center");
			mc.focus();
			return;
		}
		if (!qty || trim(qty.value) === "") {
			alert("Enter Quantity");
			if (qty) {
				qty.select();
			}
			return;
		}
		if (!checkNumbers(qty.value) || toNumber(qty.value) <= 0) {
			alert(toNumber(qty.value) <= 0 ? "Quantity cannot be ZERO" : "Enter Numerals Only");
			qty.select();
			return;
		}

		total = projectedMachineTotal(wc.value, mc ? mc.value : "", qty.value);
		if (total > toNumber(textOf("idQty"))) {
			alert("Quantity breakup should be equal to Quantity Issue");
			return;
		}

		if (!itemNode || !rootDoc) {
			return;
		}
		addDet = addDetNode(itemNode, true);
		workCenter = findWorkCenter(itemNode, wc.value);
		if (!workCenter) {
			workCenter = rootDoc.createElement("WorkCenter");
			setAttr(workCenter, "WCODE", wc.value);
			addDet.appendChild(workCenter);
		} else {
			setAttr(workCenter, "WCODE", wc.value);
		}
		machineCenter = findMachineCenter(workCenter, mc ? mc.value : "");
		if (!machineCenter) {
			machineCenter = rootDoc.createElement("MachineCenter");
			workCenter.appendChild(machineCenter);
		}
		setAttr(machineCenter, "MCODE", mc ? mc.value : "");
		setAttr(machineCenter, "QTY", trim(qty.value));
		setAttr(machineCenter, "NAME", machineDisplayName());
		renderTable();
		resetEntryFields();
	};

	window.CheckSubmit = function () {
		var itemNode = findItemDetails();
		var workCenter;
		var wc = field("selWC");

		if (wc && wc.selectedIndex === 0 && usage === "PRD") {
			alert("Select Work Center");
			wc.focus();
			return;
		}

		if (selectedUsageRequiresMachineBreakup()) {
			if (currentMachineTotal() > toNumber(textOf("idQty"))) {
				alert("Enter Full Breakup Quantity");
				return;
			}
		} else if (itemNode && rootDoc) {
			workCenter = workCenters(itemNode)[0];
			if (workCenter) {
				setAttr(workCenter, "WCODE", wc ? wc.value : "");
			} else {
				workCenter = rootDoc.createElement("WorkCenter");
				setAttr(workCenter, "WCODE", wc ? wc.value : "");
				addDetNode(itemNode, true).appendChild(workCenter);
			}
		}
		closeWithReturn();
	};

	window.setIndex = function (select, value) {
		if (!select || !select.options) {
			return;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (select.options[i].value === value) {
				select.selectedIndex = i;
				return;
			}
		}
	};

	window.ClearTable = function () {
		var table = byId("tblData");
		if (!table) {
			return;
		}
		while (table.rows.length > 1) {
			table.deleteRow(1);
		}
		rowCount = 0;
	};

	window.DeleteEntry = function () {
		var itemNode = findItemDetails();
		var addDet = addDetNode(itemNode, false);
		var wcs = workCenters(itemNode).slice();
		var selected = false;
		var w;
		var m;
		var mcs;
		var checkbox;

		for (w = 0; w < wcs.length; w += 1) {
			mcs = machineCenters(wcs[w]).slice();
			for (m = 0; m < mcs.length; m += 1) {
				checkbox = field("chkDelete" + attr(wcs[w], "WCODE") + "Z" + attr(mcs[m], "MCODE"));
				if (checkbox && checkbox.checked) {
					wcs[w].removeChild(mcs[m]);
					selected = true;
				}
			}
			if (!elementChildren(wcs[w]).length && addDet) {
				addDet.removeChild(wcs[w]);
			}
		}
		if (addDet && !elementChildren(addDet).length && itemNode) {
			itemNode.removeChild(addDet);
		}
		if (!selected) {
			alert("Please select an entry to delete");
			return;
		}
		renderTable();
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
