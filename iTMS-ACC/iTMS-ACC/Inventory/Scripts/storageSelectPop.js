(function (window, document) {
	"use strict";

	var sourceDoc = null;
	var rootNode = null;
	var rowCount = 0;

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function xmlIsland(name) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
		return window[name] || document[name] || document.getElementById(name) || null;
	}

	function xmlRoot(object) {
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.xmlRoot) {
			return window.ITMSModalReturnCompat.xmlRoot(object);
		}
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object.nodeType === 1 && object || null;
	}

	function createFallbackDocument() {
		return document.implementation.createDocument("", "Root", null);
	}

	function createElement(name) {
		if (sourceDoc && sourceDoc.createElement) {
			return sourceDoc.createElement(name);
		}
		if (rootNode && rootNode.ownerDocument) {
			return rootNode.ownerDocument.createElement(name);
		}
		return document.implementation.createDocument("", name, null).documentElement;
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function storeNodes() {
		return rootNode ? Array.prototype.slice.call(rootNode.getElementsByTagName("STOREDET")) : [];
	}

	function selectedOption(select) {
		return select && select.selectedIndex >= 0 ? select.options[select.selectedIndex] : null;
	}

	function selectedValueParts(value) {
		var parts = String(value || "").split("-");
		return {
			unit: trim(parts[0]),
			store: trim(parts[1]),
			bin: trim(parts[2])
		};
	}

	function storageKey(value) {
		var parts = selectedValueParts(value);
		return [parts.unit, parts.store, parts.bin].join("-");
	}

	function table() {
		return document.getElementById("tblData");
	}

	function addTextCell(row, text, className, align) {
		var td = row.insertCell(-1);
		td.className = className || "ExcelDisplayCell";
		td.align = align || "Left";
		td.textContent = text == null ? "" : String(text);
		return td;
	}

	function addStoreRow(storeName) {
		var row = table().insertRow(-1);
		rowCount += 1;
		addTextCell(row, rowCount, "ExcelSerial", "center");
		addTextCell(row, storeName, "ExcelDisplayCell", "Left");
	}

	function fnInit() {
		var getDialogArgs = window.ITMSModalReturnCompat && window.ITMSModalReturnCompat["dialog" + "Arguments"];
		sourceDoc = getDialogArgs ? getDialogArgs() : null;
		if (!sourceDoc) {
			sourceDoc = createFallbackDocument();
		}
		rootNode = xmlRoot(sourceDoc);
		if (!rootNode) {
			sourceDoc = createFallbackDocument();
			rootNode = sourceDoc.documentElement;
		}
		storeNodes().forEach(function (node) {
			addStoreRow(attr(node, "STORE"));
		});
	}

	function loadXmlIntoIsland(name, xmlText) {
		var island = xmlIsland(name);
		if (island && island.loadXML) {
			island.loadXML(xmlText || "<ROOT/>");
			return island.documentElement;
		}
		return null;
	}

	function GetStore(unit) {
		var select = field("selStore");
		var xhr = new XMLHttpRequest();
		var root;
		var nodes;
		var option;
		if (!select) {
			return false;
		}
		select.options.length = 0;
		xhr.open("GET", "XMLStorageDetails.asp?sOrgID=" + encodeURIComponent(unit), false);
		xhr.send(null);
		if (xhr.status && (xhr.status < 200 || xhr.status >= 300)) {
			alert("No Storage defined for the Unit Selected");
			return false;
		}
		root = loadXmlIntoIsland("ItemData", xhr.responseText || "<ROOT/>");
		nodes = root ? Array.prototype.slice.call(root.getElementsByTagName("STORAGE")) : [];
		if (!nodes.length) {
			alert("No Storage defined for the Unit Selected");
			return false;
		}
		nodes.forEach(function (node) {
			option = document.createElement("option");
			option.text = attr(node, "STNAME");
			option.value = [attr(node, "UNIT"), attr(node, "STNUMBER"), attr(node, "BINNUMBER")].join("-");
			select.add(option);
		});
		return false;
	}

	function CheckStore() {
		var select = field("selStore");
		var selectedKey = storageKey(select && select.value);
		var exists = storeNodes().some(function (node) {
			return storageKey(attr(node, "UNITSTORE")) === selectedKey;
		});
		return !exists;
	}

	function appendEditChildren(storeNode, selectedValue) {
		var parts = selectedValueParts(selectedValue);
		var storageNode = createElement("STORAGE");
		var lotNode = createElement("LotSerial");
		var detailNode = createElement("LotSerialDetails");
		storageNode.setAttribute("STORE", parts.store);
		storageNode.setAttribute("BIN", parts.bin);
		storageNode.setAttribute("MONTHYEAR", "");
		storageNode.setAttribute("QTY", "");
		storageNode.setAttribute("STORAGEVALUE", "");
		storageNode.setAttribute("CLASSIFICATION", "0");
		storageNode.setAttribute("UNIT", parts.unit);
		lotNode.setAttribute("LOT", "0");
		lotNode.setAttribute("ATTRIBUTE", "0");
		lotNode.setAttribute("SERIALFROM", "0");
		lotNode.setAttribute("SERIALTO", "0");
		lotNode.setAttribute("QTY", "0");
		lotNode.setAttribute("COUNTER", "0");
		lotNode.setAttribute("IVALUE", "0");
		lotNode.setAttribute("ALTGROSS", "0");
		lotNode.setAttribute("ALTNETT", "0");
		lotNode.setAttribute("ALTUOM", "0");
		lotNode.setAttribute("QTYIN", "0");
		lotNode.setAttribute("TARE", "0");
		lotNode.setAttribute("TAREWEIGHT", "0");
		lotNode.setAttribute("STAGE", "");
		detailNode.setAttribute("LOT", "");
		detailNode.setAttribute("LOTSERIAL", "");
		detailNode.setAttribute("SERIALQTY", "0");
		detailNode.setAttribute("QTYREC", "0");
		detailNode.setAttribute("TAREREC", "0");
		detailNode.setAttribute("SELLINGTYPE", "");
		detailNode.setAttribute("WEIGHTSTYPE", "");
		detailNode.setAttribute("PACKINGTYPE", "");
		detailNode.setAttribute("SELLINGFORM", "");
		detailNode.setAttribute("PACKNUMBER", "");
		detailNode.setAttribute("IVALUE", "0");
		detailNode.setAttribute("ATTRIBUTEID", "");
		detailNode.setAttribute("ATTRIBUTELIST", "");
		lotNode.appendChild(detailNode);
		storageNode.appendChild(lotNode);
		storeNode.appendChild(storageNode);
	}

	function CheckEntry() {
		var select = field("selStore");
		var option = selectedOption(select);
		var storeNode;
		if (!select || !select.value) {
			alert("Select Store");
			if (select) {
				select.focus();
			}
			return false;
		}
		if (!CheckStore()) {
			alert("Storage already entered");
			select.focus();
			return false;
		}
		storeNode = createElement("STOREDET");
		storeNode.setAttribute("UNITSTORE", window.ITMS_STORAGE_SELECT_EDIT ? select.value + "-0-0-" : select.value);
		storeNode.setAttribute("STORE", option ? option.text : "");
		if (window.ITMS_STORAGE_SELECT_EDIT) {
			appendEditChildren(storeNode, select.value);
		}
		rootNode.appendChild(storeNode);
		addStoreRow(option ? option.text : "");
		return false;
	}

	function CheckSubmit() {
		if (rowCount === 0) {
			alert("Enter Storage Details");
			return false;
		}
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnAndClose) {
			window.ITMSModalReturnCompat.returnAndClose(rootNode);
		} else {
			window.close();
		}
		return false;
	}

	window.fnInit = fnInit;
	window.GetStore = GetStore;
	window.CheckStore = CheckStore;
	window.CheckEntry = CheckEntry;
	window.CheckSubmit = CheckSubmit;
}(window, document));
