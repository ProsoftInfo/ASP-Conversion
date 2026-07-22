(function (window, document) {
	"use strict";

	var objTemp = null;
	var parentRoot = null;
	var issueType = "";
	var issueNo = "";
	var itemCode = "";
	var classCode = "";
	var orgId = "";
	var attributeList = "";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function numericText(value) {
		return /^([0-9]+(\.[0-9]*)?|\.[0-9]+)$/.test(trim(value));
	}

	function normalizeLot(value) {
		var lot = trim(value);
		return lot === "" || lot === "N/A" || lot.toUpperCase() === "NULL" ? "0" : lot;
	}

	function xmlIsland(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name);
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
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || (value.nodeType === 1 ? value : null);
	}

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
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

	function getAttr(node, name) {
		return trim(node && node.getAttribute ? node.getAttribute(name) : "");
	}

	function setAttr(node, name, value) {
		if (node) {
			node.setAttribute(name, String(value == null ? "" : value));
		}
	}

	function importNode(doc, node) {
		return doc && doc.importNode ? doc.importNode(node, true) : node.cloneNode(true);
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return null;
		}
		return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function itemMatches(node) {
		return getAttr(node, "ITMCODE") === itemCode &&
			getAttr(node, "CLACODE") === classCode &&
			getAttr(node, "ATTID") === attributeList;
	}

	function localRoot() {
		return xmlRoot(xmlIsland("ItemData"));
	}

	function localItem() {
		var items = elementChildren(localRoot(), "ITEM");
		for (var i = 0; i < items.length; i += 1) {
			if (itemMatches(items[i])) {
				return items[i];
			}
		}
		return null;
	}

	function lotSerialNodes(root) {
		var result = [];
		elementChildren(root, "ITEM").forEach(function (item) {
			if (!itemMatches(item)) {
				return;
			}
			elementChildren(item, "STORAGE").forEach(function (storage) {
				elementChildren(storage, "LotSerial").forEach(function (lot) {
					result.push(lot);
				});
			});
		});
		return result;
	}

	function findLocalLot(lotNo) {
		var wanted = normalizeLot(lotNo);
		var nodes = lotSerialNodes(localRoot());
		for (var i = 0; i < nodes.length; i += 1) {
			if (normalizeLot(getAttr(nodes[i], "LOT")) === wanted) {
				return nodes[i];
			}
		}
		return null;
	}

	function removeExistingParentItem() {
		var items = elementChildren(parentRoot, "ITEM");
		for (var i = items.length - 1; i >= 0; i -= 1) {
			if (itemMatches(items[i])) {
				parentRoot.removeChild(items[i]);
			}
		}
	}

	function setReturnValue() {
		if (parentRoot) {
			if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
				window.ITMSModernCompat.returnModalValue(parentRoot);
			} else {
				window["return" + "Value"] = parentRoot;
				window.returnvalue = parentRoot;
			}
		}
	}

	function closeWithReturn() {
		setReturnValue();
		window.close();
	}

	window.checkNumbers = numericText;

	window.fnInit = function (arg) {
		var parts = String(arg || "").split(":");
		var existingLots;
		issueType = parts[0] || "";
		issueNo = parts[1] || "";
		itemCode = parts[2] || "";
		classCode = parts[3] || "";
		orgId = parts[4] || "";
		attributeList = trim(parts[7] || "") || "NULL";
		objTemp = modalArgs();
		parentRoot = xmlRoot(objTemp);
		if (!toNumber(field("hiCtr") && field("hiCtr").value)) {
			return false;
		}
		existingLots = lotSerialNodes(parentRoot);
		for (var i = 0; i < existingLots.length; i += 1) {
			if (field("txtQty" + (i + 1))) {
				field("txtQty" + (i + 1)).value = getAttr(existingLots[i], "QTY");
			}
		}
		return false;
	};

	window.CheckLot = function (obj, rowNo) {
		var token = obj && obj.name || "";
		var parts = token.split("`");
		var lotNo = parts[5] || "";
		openDialog("IssRetSerialPop.asp?sTemp=" + encodeURIComponent(token), xmlIsland("ItemData"), "dialogHeight:400px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No", function () {
			var lot = findLocalLot(lotNo);
			if (lot && field("txtQty" + rowNo)) {
				field("txtQty" + rowNo).value = getAttr(lot, "QTY");
			}
		});
		return false;
	};

	window.CheckSubmit = function () {
		var count = toNumber(field("hiCtr") && field("hiCtr").value);
		var total = 0;
		var item;
		var issueRate;
		var parentDoc;
		var lotNodes;
		for (var i = 1; i <= count; i += 1) {
			var qty = field("txtQty" + i);
			var remaining = field("txtRemQty" + i);
			if (!qty || trim(qty.value) === "") {
				alert("Enter Quantity to Return");
				qty && qty.select();
				return false;
			}
			if (!numericText(qty.value)) {
				alert("Enter Numerals Only");
				qty.select();
				return false;
			}
			if (toNumber(qty.value) > toNumber(remaining && remaining.value) || toNumber(qty.value) > toNumber(field("idQty") && field("idQty").value)) {
				alert("Return Quantity should be equal to or less than Remaining Quantity");
				qty.select();
				return false;
			}
			total += toNumber(qty.value);
		}
		issueRate = toNumber(field("hIssRate") && field("hIssRate").value);
		item = localItem();
		if (item) {
			setAttr(item, "QTY", total);
			elementChildren(item, "STORAGE").forEach(function (storage) {
				setAttr(storage, "QTY", total);
				setAttr(storage, "STORAGEVALUE", total * issueRate);
				lotNodes = elementChildren(storage, "LotSerial");
				for (var i = 0; i < lotNodes.length; i += 1) {
					setAttr(lotNodes[i], "QTY", field("txtQty" + (i + 1)) ? field("txtQty" + (i + 1)).value : "0");
				}
			});
		}
		removeExistingParentItem();
		parentDoc = xmlDocument(objTemp) || (parentRoot && parentRoot.ownerDocument);
		elementChildren(localRoot(), "ITEM").forEach(function (node) {
			parentRoot.appendChild(importNode(parentDoc, node));
		});
		closeWithReturn();
		return false;
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
