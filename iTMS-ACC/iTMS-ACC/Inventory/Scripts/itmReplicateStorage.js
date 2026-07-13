(function (window, document) {
	"use strict";

	var objTemp = null;
	var rootO = null;
	var rowNo = 0;

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

	function byId(id) {
		return document.getElementById(id);
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || byId(name) || null;
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

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function attrByIndex(node, index) {
		return node && node.attributes && node.attributes[index] ? node.attributes[index].value : "";
	}

	function selectedText(select) {
		if (!select || select.selectedIndex < 0 || !select.options[select.selectedIndex]) {
			return "";
		}
		return select.options[select.selectedIndex].text;
	}

	function appendRow(storeName) {
		var table = byId("tblData");
		var row;
		var cell;
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
		cell.textContent = trim(storeName);
		cell.className = "ExcelDisplayCell";
	}

	function renderRows() {
		var table = byId("tblData");
		if (table) {
			while (table.rows.length > 1) {
				table.deleteRow(1);
			}
		}
		rowNo = 0;
		childElements(rootO, "STOREDET").forEach(function (node) {
			appendRow(attr(node, "STORE"));
		});
	}

	function fnInit() {
		objTemp = dialogDocument();
		rootO = objTemp.documentElement;
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.install) {
			window.ITMSModalReturnCompat.install(function () {
				return rootO;
			});
		}
		renderRows();
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

	function addOption(select, value, text) {
		var option = document.createElement("option");
		option.value = value;
		option.text = text;
		select.add(option);
	}

	function GetStore(obj) {
		var root;
		var select = field("selStore");
		if (!select) {
			return false;
		}
		select.options.length = 0;
		root = loadXmlIntoIsland("ItemData", "XMLStorageDetails.asp?sOrgID=" + encodeURIComponent(obj.value));
		if (!root || !childElements(root).length) {
			alert("No Storage defined for the Unit Selected");
			return false;
		}
		childElements(root).forEach(function (node) {
			addOption(select, attrByIndex(node, 0) + "-" + attrByIndex(node, 1) + "-" + attrByIndex(node, 4), attrByIndex(node, 3));
		});
		return false;
	}

	function CheckStore() {
		var store = field("selStore") ? field("selStore").value : "";
		var exists = false;
		childElements(rootO, "STOREDET").forEach(function (node) {
			var existing = attr(node, "UNITSTORE");
			if (existing === store || existing.indexOf(store + "-") === 0) {
				exists = true;
			}
		});
		return !exists;
	}

	function CheckEntry() {
		var store = field("selStore");
		var parts;
		var storeNode;
		var storageNode;
		if (!store || !store.value) {
			alert("Select Store");
			if (store) {
				store.focus();
			}
			return false;
		}
		if (!CheckStore()) {
			alert("Storage already entered");
			store.focus();
			return false;
		}
		parts = store.value.split("-");
		storeNode = objTemp.createElement("STOREDET");
		storeNode.setAttribute("UNITSTORE", store.value + "-0-0-" + (field("hMYr") ? field("hMYr").value : ""));
		storeNode.setAttribute("STORE", selectedText(store));
		storageNode = objTemp.createElement("STORAGE");
		storageNode.setAttribute("STORE", parts[1] || "");
		storageNode.setAttribute("BIN", parts[2] || "");
		storageNode.setAttribute("MONTHYEAR", "");
		storageNode.setAttribute("QTY", "");
		storageNode.setAttribute("STORAGEVALUE", "");
		storageNode.setAttribute("CLASSIFICATION", "0");
		storageNode.setAttribute("UNIT", parts[0] || "");
		storeNode.appendChild(storageNode);
		rootO.appendChild(storeNode);
		appendRow(selectedText(store));
		return false;
	}

	function CheckSubmit() {
		var unit = field("selUnit");
		if (unit && unit.selectedIndex === -1) {
			alert("Select Unit");
			unit.focus();
			return false;
		}
		if (rowNo === 0) {
			alert("Enter Storage Details");
			return false;
		}
		if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnAndClose) {
			window.ITMSModalReturnCompat.returnAndClose(rootO);
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
