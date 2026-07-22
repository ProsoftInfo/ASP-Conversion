(function (window, document) {
	"use strict";

	var objTemp = null;
	var root = null;
	var rootDoc = null;
	var lotNo = "";
	var classCode = "";
	var itemCode = "";
	var orgCode = "";
	var locCode = "";
	var binCode = "";

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

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function normalizeLot(value) {
		var lot = trim(value);
		return lot === "" || lot === "-" || lot.toUpperCase() === "N/A" || lot.toUpperCase() === "NULL" ? "NULL" : lot;
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

	function getAttr(node, name) {
		return trim(node && node.getAttribute ? node.getAttribute(name) : "");
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, String(value == null ? "" : value));
		}
	}

	function createElement(name) {
		return (rootDoc || document).createElement(name);
	}

	function pickNodes() {
		var picks = [];
		var roots = root ? [root] : [];
		var children;
		for (var r = 0; r < roots.length; r += 1) {
			children = elementChildren(roots[r]);
			for (var i = 0; i < children.length; i += 1) {
				if (String(children[i].nodeName).toLowerCase() === "pick") {
					picks.push(children[i]);
				} else {
					roots.push(children[i]);
				}
			}
		}
		return picks;
	}

	function pickContainer() {
		var nodes = root ? [root] : [];
		for (var i = 0; i < nodes.length; i += 1) {
			if (String(nodes[i].nodeName).toLowerCase() === "pickdet") {
				return nodes[i];
			}
			var children = elementChildren(nodes[i]);
			for (var c = 0; c < children.length; c += 1) {
				nodes.push(children[c]);
			}
		}
		return root;
	}

	function currentPick(createIfMissing) {
		var picks = pickNodes();
		var lot = normalizeLot(lotNo);
		for (var i = 0; i < picks.length; i += 1) {
			if (trim(getAttr(picks[i], "LOC")) === trim(locCode) &&
					trim(getAttr(picks[i], "BIN")) === trim(binCode) &&
					normalizeLot(getAttr(picks[i], "LOTNO")) === lot) {
				return picks[i];
			}
		}
		if (!createIfMissing || !root) {
			return null;
		}
		var pick = createElement("PICK");
		setAttr(pick, "LOC", locCode);
		setAttr(pick, "BIN", binCode);
		setAttr(pick, "LOTNO", lot);
		setAttr(pick, "ISSQTY", "");
		pickContainer().appendChild(pick);
		return pick;
	}

	function removeSerialHeader(pick) {
		var headers = elementChildren(pick, "SERIALHEADER");
		for (var i = 0; i < headers.length; i += 1) {
			pick.removeChild(headers[i]);
		}
	}

	function selectedSerialDetails(pick) {
		var details = [];
		var headers = elementChildren(pick, "SERIALHEADER");
		for (var h = 0; h < headers.length; h += 1) {
			var serials = elementChildren(headers[h], "SERIALDETAILS");
			for (var s = 0; s < serials.length; s += 1) {
				details.push(serials[s]);
			}
		}
		return details;
	}

	function rowCount() {
		return toNumber(field("hiCtr") && field("hiCtr").value);
	}

	function setRowChecked(row, checked) {
		var checkbox = field("ChkItem" + row);
		var qty = field("txtQty" + row);
		var stock = field("txtStQty" + row);
		if (checkbox) {
			checkbox.checked = !!checked;
		}
		if (qty) {
			qty.value = checked ? (toNumber(qty.value) > 0 ? qty.value : stock && stock.value || qty.value) : "0";
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
		var pick;
		var details;
		objTemp = modalArgs();
		root = xmlRoot(objTemp);
		rootDoc = xmlDocument(objTemp);
		if (!rootDoc && root) {
			rootDoc = root.ownerDocument;
		}
		lotNo = parts[1] || "";
		classCode = parts[2] || "";
		itemCode = parts[3] || "";
		orgCode = parts[4] || "";
		locCode = parts[5] || "";
		binCode = parts[6] || "";

		pick = currentPick(false);
		details = selectedSerialDetails(pick);
		for (var d = 0; d < details.length; d += 1) {
			for (var row = 1; row <= rowCount(); row += 1) {
				if (field("hSerial" + row) && trim(field("hSerial" + row).value) === getAttr(details[d], "SERIALNO")) {
					setRowChecked(row, true);
					if (field("txtQty" + row)) {
						field("txtQty" + row).value = getAttr(details[d], "QTY");
					}
				}
			}
		}
	};

	window.SelectItem = function (row) {
		var checkbox = field("ChkItem" + row);
		setRowChecked(row, checkbox && checkbox.checked);
	};

	window.SelectAll = function () {
		var all = field("ChkAll");
		var checked = all && all.checked;
		for (var row = 1; row <= rowCount(); row += 1) {
			setRowChecked(row, checked);
		}
	};

	window.CheckSubmit = function () {
		var pick = currentPick(true);
		var serialHeader = createElement("SERIALHEADER");
		var total = 0;
		var selected = 0;
		var serial;
		var detail;
		var qty;
		var stock;

		for (var row = 1; row <= rowCount(); row += 1) {
			if (!field("ChkItem" + row) || !field("ChkItem" + row).checked) {
				continue;
			}
			serial = field("hSerial" + row);
			qty = toNumber(field("txtQty" + row) && field("txtQty" + row).value);
			stock = toNumber(field("txtStQty" + row) && field("txtStQty" + row).value);
			if (qty <= 0) {
				alert("Enter Quantity");
				if (field("txtQty" + row)) {
					field("txtQty" + row).focus();
				}
				return;
			}
			if (stock > 0 && qty > stock) {
				alert("Quantity should not be greater than Stock");
				if (field("txtQty" + row)) {
					field("txtQty" + row).focus();
				}
				return;
			}
			detail = createElement("SERIALDETAILS");
			setAttr(detail, "SERIALNO", serial ? serial.value : "NULL");
			setAttr(detail, "QTY", qty);
			setAttr(detail, "LOTNO", normalizeLot(lotNo));
			setAttr(detail, "CLASS", classCode);
			setAttr(detail, "ITEM", itemCode);
			setAttr(detail, "ORG", orgCode);
			serialHeader.appendChild(detail);
			total += qty;
			selected += 1;
		}

		if (!selected) {
			alert("Select the Serial Numbers");
			return;
		}

		removeSerialHeader(pick);
		pick.appendChild(serialHeader);
		setAttr(pick, "ISSQTY", total);
		closeWithReturn();
	};

	window.window_onunload = setReturnValue;
	window.addEventListener("beforeunload", setReturnValue);
}(window, document));
