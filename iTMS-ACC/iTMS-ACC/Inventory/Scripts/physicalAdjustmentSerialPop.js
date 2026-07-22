(function (window, document) {
	"use strict";

	var root = null;
	var rootDoc = null;
	var sourceStore = "";
	var sourceBin = "";
	var sourceLot = "";
	var inventoryReceiptNo = "";

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

	function isSignedNumericText(value) {
		return /^-?([0-9]+(\.[0-9]*)?|\.[0-9]+)$/.test(trim(value));
	}

	function textOf(id) {
		var element = byId(id);
		return trim(element ? element.textContent || element.innerText || "" : "");
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
		if (node && node.setAttribute) {
			node.setAttribute(name, String(value == null ? "" : value));
		}
	}

	function createElement(name) {
		return rootDoc.createElement(name);
	}

	function normalizeLot(value) {
		var lot = trim(value);
		return lot === "0" || lot === "" ? "NULL" : lot;
	}

	function findPick() {
		var found = null;
		elementChildren(root, "Item").some(function (item) {
			return elementChildren(item, "LOCDET").some(function (loc) {
				if (getAttr(loc, "LOC") !== sourceStore || getAttr(loc, "BIN") !== sourceBin) {
					return false;
				}
				return elementChildren(loc, "PICK").some(function (pick) {
					if (normalizeLot(getAttr(pick, "LOTNO")) === sourceLot && getAttr(pick, "INVRECNO") === inventoryReceiptNo) {
						found = pick;
						return true;
					}
					return false;
				});
			});
		});
		return found;
	}

	function removeSerialDetails(pick) {
		elementChildren(pick, "SERIALDETAILS").forEach(function (serial) {
			pick.removeChild(serial);
		});
	}

	function rowCount() {
		return toNumber(field("hiCtr") && field("hiCtr").value);
	}

	function setReturnValue() {
		if (root) {
			if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
				window.ITMSModernCompat.returnModalValue(root);
			} else {
				window["return" + "Value"] = root;
				window.returnvalue = root;
			}
		}
	}

	function closeWithReturn() {
		setReturnValue();
		window.close();
	}

	window.fnInit = function (arg) {
		var parts = String(arg || "").split("AAAA");
		var pick;
		var args = modalArgs();
		root = xmlRoot(args);
		rootDoc = xmlDocument(args) || root && root.ownerDocument;
		sourceLot = normalizeLot(parts[4] || "");
		sourceStore = parts[6] || "";
		sourceBin = parts[7] || "";
		inventoryReceiptNo = parts[8] || "";
		pick = findPick();
		if (!pick) {
			return false;
		}
		elementChildren(pick, "SERIALDETAILS").forEach(function (serial) {
			var entryNo = getAttr(serial, "ENTRYNO");
			if (field("txtQty" + entryNo)) {
				field("txtQty" + entryNo).value = getAttr(serial, "QTY");
			}
		});
		return false;
	};

	function validateRows() {
		var total = 0;
		var qty;
		for (var i = 1; i <= rowCount(); i += 1) {
			qty = field("txtQty" + i);
			if (!qty || trim(qty.value) === "") {
				alert("Enter Quantity");
				qty && qty.select();
				return false;
			}
			if (!isSignedNumericText(qty.value)) {
				alert("Enter Numerals Only");
				qty.select();
				return false;
			}
			total += toNumber(qty.value);
		}
		if (total !== toNumber(textOf("idQty"))) {
			alert("Total Adjust Quantity should be equal to Quantity (" + textOf("idQty") + ")");
			return false;
		}
		return total;
	}

	window.CheckSubmit = function () {
		var total = validateRows();
		var pick;
		var detail;
		if (total === false) {
			return false;
		}
		pick = findPick();
		if (!pick) {
			return false;
		}
		removeSerialDetails(pick);
		for (var i = 1; i <= rowCount(); i += 1) {
			if (toNumber(field("txtQty" + i).value) !== 0) {
				detail = createElement("SERIALDETAILS");
				setAttr(detail, "SERIALNO", field("txtSerial" + i) ? field("txtSerial" + i).value : "");
				setAttr(detail, "STKQTY", field("txtStQty" + i) ? field("txtStQty" + i).value : "");
				setAttr(detail, "QTY", field("txtQty" + i).value);
				setAttr(detail, "ENTRYNO", i);
				pick.appendChild(detail);
			}
		}
		setAttr(pick, "QTYISS", total);
		closeWithReturn();
		return false;
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
