(function (window, document) {
	"use strict";

	var objTemp = null;
	var root = null;
	var rootDoc = null;
	var itemCode = "";
	var classCode = "";
	var lotNo = "";
	var issueDate = "";
	var attrId = "";
	var qtyTotal = 0;

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
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function isNumericText(value) {
		return /^([0-9]+(\.[0-9]*)?|\.[0-9]+)$/.test(trim(value));
	}

	function normalizeLot(value) {
		var lot = trim(value);
		return lot.toUpperCase() === "NULL" || lot.toUpperCase() === "N/A" ? "0" : lot;
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

	function descendantsByName(node, name) {
		var result = [];
		var wanted = String(name || "").toLowerCase();
		var children = elementChildren(node);
		for (var i = 0; i < children.length; i += 1) {
			if (String(children[i].nodeName).toLowerCase() === wanted) {
				result.push(children[i]);
			}
			result = result.concat(descendantsByName(children[i], name));
		}
		return result;
	}

	function getAttr(node, name) {
		return trim(node && node.getAttribute ? node.getAttribute(name) : "");
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, trim(value));
		}
	}

	function fieldValue(name) {
		var item = field(name);
		return trim(item ? item.value : "");
	}

	function createElement(name) {
		if (rootDoc && rootDoc.createElement) {
			return rootDoc.createElement(name);
		}
		return document.implementation.createDocument("", name, null).documentElement;
	}

	function rowCount() {
		return toNumber(field("hiCtr") && field("hiCtr").value);
	}

	function currentItemNode() {
		var nodes = descendantsByName(root, "ITEM");
		for (var i = 0; i < nodes.length; i += 1) {
			if (getAttr(nodes[i], "ITMCODE") === trim(itemCode) &&
					getAttr(nodes[i], "CLACODE") === trim(classCode) &&
					getAttr(nodes[i], "ATTID") === trim(attrId)) {
				return nodes[i];
			}
		}
		return null;
	}

	function currentLotNode() {
		var item = currentItemNode();
		var storageNodes = elementChildren(item, "STORAGE");
		var lotNodes;
		for (var i = 0; i < storageNodes.length; i += 1) {
			lotNodes = elementChildren(storageNodes[i], "LotSerial");
			for (var j = 0; j < lotNodes.length; j += 1) {
				if (getAttr(lotNodes[j], "LOT") === lotNo) {
					return lotNodes[j];
				}
			}
		}
		return null;
	}

	function removeLotSerialDetails(lotNode) {
		var details = elementChildren(lotNode, "LotSerialDetails");
		for (var i = 0; i < details.length; i += 1) {
			lotNode.removeChild(details[i]);
		}
	}

	function setReturnValue() {
		if (!root) {
			root = xmlRoot(objTemp);
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

	window.fnInit = function (arg) {
		var parts = String(arg || "").split("`");
		var lotNode;
		var details;
		var qty;

		objTemp = modalArgs();
		root = xmlRoot(objTemp);
		rootDoc = xmlDocument(objTemp);
		if (!rootDoc && root) {
			rootDoc = root.ownerDocument;
		}

		itemCode = parts[3] || "";
		classCode = parts[4] || "";
		lotNo = normalizeLot(parts[5] || "");
		issueDate = parts[9] || "";
		attrId = parts[10] || "";

		if (!rowCount() || !root) {
			return;
		}

		lotNode = currentLotNode();
		details = elementChildren(lotNode, "LotSerialDetails");
		for (var i = 0; i < details.length; i += 1) {
			qty = field("txtQty" + (i + 1));
			if (qty) {
				qty.value = getAttr(details[i], "QTYREC");
			}
		}
	};

	window.CheckSubmit = function () {
		var rows = rowCount();
		var lotNode;
		var detail;
		var qty;
		var stock;

		qtyTotal = 0;
		if (!rows) {
			return false;
		}

		for (var i = 1; i <= rows; i += 1) {
			qty = field("txtQty" + i);
			stock = field("txtStQty" + i);
			if (!qty || trim(qty.value) === "") {
				alert("Enter Quantity");
				if (qty) {
					qty.select();
				}
				return false;
			}
			if (!isNumericText(qty.value)) {
				alert("Enter Numerals Only");
				qty.select();
				return false;
			}
			if (toNumber(qty.value) > toNumber(stock && stock.value)) {
				alert("Return Quantity should be equal to or less than Quantity Remaining");
				qty.select();
				return false;
			}
			qtyTotal += toNumber(qty.value);
		}

		lotNo = normalizeLot(lotNo);
		lotNode = currentLotNode();
		if (lotNode) {
			removeLotSerialDetails(lotNode);
			setAttr(lotNode, "QTY", qtyTotal);
			for (i = 1; i <= rows; i += 1) {
				detail = createElement("LotSerialDetails");
				setAttr(detail, "LOTSERIAL", fieldValue("hSerial" + i));
				setAttr(detail, "QTYREC", fieldValue("txtQty" + i));
				setAttr(detail, "TAREREC", "0");
				setAttr(detail, "SELLINGTYPE", "0");
				setAttr(detail, "WEIGHTSTYPE", "0");
				setAttr(detail, "PACKINGTYPE", fieldValue("hPackCode" + i));
				setAttr(detail, "LOT", lotNo);
				setAttr(detail, "SELLINGFORM", "0");
				setAttr(detail, "PACKNUMBER", fieldValue("hPackNumber" + i));
				setAttr(detail, "IVALUE", "0");
				setAttr(detail, "ATTRIBUTELIST", fieldValue("hAttributeList" + i));
				lotNode.appendChild(detail);
			}
		}

		closeWithReturn();
		return false;
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
