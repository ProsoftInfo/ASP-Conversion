(function (window, document) {
	"use strict";

	var objTemp = null;
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
		var value = frm && frm.elements ? frm.elements[name] : null;
		if (value && value.length && !value.tagName) {
			return value[0] || null;
		}
		return value || null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function normalizeLot(value) {
		var lot = trim(value);
		return lot === "" || lot === "-" || lot === "0" || lot.toUpperCase() === "N/A" || lot.toUpperCase() === "NULL" ? "NULL" : lot;
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

	function rowCount() {
		return toNumber(field("hiCtr") && field("hiCtr").value);
	}

	function dataRows() {
		var island = window.Data || document.getElementById("Data");
		var dataRoot = xmlRoot(island);
		return elementChildren(dataRoot, "SERIAL");
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

	function createElement(name) {
		return (rootDoc || document).createElement(name);
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

	function findRowBySerial(serialNo) {
		for (var i = 1; i <= rowCount(); i += 1) {
			if (field("txtSerial" + i) && trim(field("txtSerial" + i).value) === trim(serialNo)) {
				return i;
			}
		}
		return 0;
	}

	function fillExistingSerials(pick) {
		var nextRow = 1;
		elementChildren(pick, "SERIALDETAILS").forEach(function (serial) {
			var row = findRowBySerial(getAttr(serial, "SERIALNO")) || nextRow;
			nextRow += 1;
			if (field("txtRejQty" + row)) {
				field("txtRejQty" + row).value = getAttr(serial, "QTYREJ");
			}
			if (field("txtOhOQty" + row)) {
				field("txtOhOQty" + row).value = getAttr(serial, "QTYOHO");
			}
			if (field("txtResQty" + row)) {
				field("txtResQty" + row).value = getAttr(serial, "QTYRES");
			}
			if (field("txtRem" + row)) {
				field("txtRem" + row).value = getAttr(serial, "REMARKS");
			}
		});
	}

	function validateRows() {
		var rowsXml = dataRows();
		var totals = {
			clean: 0,
			rejected: 0,
			onHold: 0,
			reserved: 0
		};
		for (var i = 1; i <= rowCount(); i += 1) {
			var xmlRow = rowsXml[i - 1];
			var rejected = field("txtRejQty" + i);
			var onHold = field("txtOhOQty" + i);
			var reserved = field("txtResQty" + i);
			var clean = field("txtQty" + i);
			var previousQty;
			var currentQty;
			var difference;
			if (!rejected || trim(rejected.value) === "") {
				alert("Enter Rejected Stock");
				rejected && rejected.select();
				return false;
			}
			if (!onHold || trim(onHold.value) === "") {
				alert("Enter On Hold Stock");
				onHold && onHold.select();
				return false;
			}
			if (!reserved || trim(reserved.value) === "") {
				alert("Enter On Reserved Stock");
				reserved && reserved.select();
				return false;
			}
			currentQty = toNumber(rejected.value) + toNumber(onHold.value) + toNumber(reserved.value);
			previousQty = toNumber(getAttr(xmlRow, "RESQTY")) + toNumber(getAttr(xmlRow, "OHOQTY")) + toNumber(getAttr(xmlRow, "REJQTY"));
			difference = currentQty - previousQty;
			if (clean) {
				clean.value = difference;
			}
			if (trim(field("txtRem" + i) && field("txtRem" + i).value) === "" && difference !== 0) {
				alert("Enter Remarks");
				field("txtRem" + i).select();
				return false;
			}
			totals.clean += difference;
			totals.rejected += toNumber(rejected.value);
			totals.onHold += toNumber(onHold.value);
			totals.reserved += toNumber(reserved.value);
		}
		return totals;
	}

	function appendSerialDetails(pick) {
		for (var i = 1; i <= rowCount(); i += 1) {
			var serial = createElement("SERIALDETAILS");
			setAttr(serial, "SERIALNO", field("txtSerial" + i) ? field("txtSerial" + i).value : "");
			setAttr(serial, "STKQTY", field("txtStQty" + i) ? field("txtStQty" + i).value : "");
			setAttr(serial, "QTY", field("txtQty" + i) ? field("txtQty" + i).value : "");
			setAttr(serial, "QTYREJ", field("txtRejQty" + i) ? field("txtRejQty" + i).value : "");
			setAttr(serial, "QTYOHO", field("txtOhOQty" + i) ? field("txtOhOQty" + i).value : "");
			setAttr(serial, "QTYRES", field("txtResQty" + i) ? field("txtResQty" + i).value : "");
			setAttr(serial, "REMARKS", field("txtRem" + i) ? field("txtRem" + i).value : "");
			pick.appendChild(serial);
		}
	}

	window.fnInit = function (arg) {
		var parts = String(arg || "").split("AAAA");
		var pick;
		objTemp = modalArgs();
		root = xmlRoot(objTemp);
		rootDoc = xmlDocument(objTemp);
		if (!rootDoc && root) {
			rootDoc = root.ownerDocument;
		}
		sourceLot = normalizeLot(parts[4] || "");
		sourceStore = parts[6] || "";
		sourceBin = parts[7] || "";
		inventoryReceiptNo = parts[8] || "";
		if (!rowCount()) {
			return false;
		}
		pick = findPick();
		if (pick) {
			fillExistingSerials(pick);
		}
		return false;
	};

	window.CheckSubmit = function () {
		var totals;
		var pick;
		if (!rowCount()) {
			return false;
		}
		totals = validateRows();
		if (totals === false) {
			return false;
		}
		pick = findPick();
		if (!pick) {
			return false;
		}
		removeSerialDetails(pick);
		appendSerialDetails(pick);
		setAttr(pick, "QTYCLE", totals.clean);
		setAttr(pick, "QTYREJ", totals.rejected);
		setAttr(pick, "QTYOHO", totals.onHold);
		setAttr(pick, "QTYRES", totals.reserved);
		closeWithReturn();
		return false;
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
