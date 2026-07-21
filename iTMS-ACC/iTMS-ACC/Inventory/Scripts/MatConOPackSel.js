(function (window, document) {
	"use strict";

	var objTemp = null;
	var root = null;
	var rootDoc = null;
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

	function elementChildren(node) {
		var result = [];
		var children = node && node.childNodes || [];
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1) {
				result.push(children[i]);
			}
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
		return toNumber(field("hCtr") && field("hCtr").value);
	}

	function receiptNode(criteria) {
		var nodes = elementChildren(root);
		for (var i = 0; i < nodes.length; i += 1) {
			if (getAttr(nodes[i], "No") === criteria.receiptNo &&
					getAttr(nodes[i], "Item") === criteria.itemCode &&
					getAttr(nodes[i], "AttID") === criteria.attributeList &&
					getAttr(nodes[i], "RNumbering") === criteria.receiptNumber) {
				return nodes[i];
			}
		}
		return null;
	}

	function parseRequest(value) {
		var parts = String(value || "").split(":");
		return {
			receiptNo: parts[0] || "",
			itemCode: parts[1] || "",
			attributeList: parts[2] || "",
			quantity: parts[3] || "",
			receiptNumber: parts[4] || ""
		};
	}

	function setReturnValue() {
		if (!root) {
			root = xmlRoot(objTemp);
		}
		if (root) {
			setAttr(root, "SerQtyRet", qtyTotal);
		}
		window.returnValue = root;
		window.returnvalue = root;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		}
	}

	function closeWithReturn() {
		setReturnValue();
		window.close();
	}

	window.fnInit = function () {
		objTemp = window.dialogArguments;
		root = xmlRoot(objTemp);
		rootDoc = xmlDocument(objTemp);
		if (!rootDoc && root) {
			rootDoc = root.ownerDocument;
		}
	};

	window.CheckSubmit = function (requestValue) {
		var rows = rowCount();
		var criteria = parseRequest(requestValue);
		var target;
		var qty;
		var stock;
		var detail;
		var receiptQty = 0;

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

		target = receiptNode(criteria);
		if (target) {
			for (i = 1; i <= rows; i += 1) {
				qty = field("txtQty" + i);
				if (toNumber(qty && qty.value) > 0) {
					detail = createElement("SerialDetails");
					setAttr(detail, "SerNo", fieldValue("hSerial" + i));
					setAttr(detail, "Qty", qty.value);
					receiptQty += toNumber(qty.value);
					target.appendChild(detail);
				}
			}
			setAttr(target, "Qty", receiptQty);
		}

		closeWithReturn();
		return false;
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
