(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		var item;
		var wanted;
		var i;
		if (!frm || !frm.elements) {
			return null;
		}
		item = frm.elements[name];
		if (item) {
			return item;
		}
		wanted = String(name).toLowerCase();
		for (i = 0; i < frm.elements.length; i += 1) {
			if (String(frm.elements[i].name || "").toLowerCase() === wanted) {
				return frm.elements[i];
			}
		}
		return null;
	}

	function valueOf(name) {
		var item = field(name);
		return item ? item.value : "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function checked(name) {
		var item = field(name);
		return !!(item && item.checked);
	}

	function setChecked(name, value) {
		var item = field(name);
		if (item) {
			item.checked = !!value;
		}
	}

	function radioItems(name) {
		var items = form() && form().elements ? form().elements[name] : null;
		if (!items) {
			return [];
		}
		if (typeof items.length === "number" && !items.tagName) {
			return Array.prototype.slice.call(items);
		}
		return [items];
	}

	function radioValue(name) {
		var items = radioItems(name);
		for (var i = 0; i < items.length; i += 1) {
			if (items[i].checked) {
				return items[i].value;
			}
		}
		return "";
	}

	function setRadioValue(name, value, fallback) {
		var matched = false;
		radioItems(name).forEach(function (item) {
			item.checked = item.value === value;
			matched = matched || item.checked;
		});
		if (!matched && fallback != null) {
			setRadioValue(name, fallback);
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
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function xmlDocument(name) {
		var object = xmlObject(name);
		var root = xmlRoot(object);
		return object && object.XMLDocument || object && object._doc || root && root.ownerDocument || null;
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

	function clearNamedChildren(root, name) {
		childElements(root, name).forEach(function (node) {
			root.removeChild(node);
		});
	}

	function serializeXml(rootOrDoc) {
		var doc = rootOrDoc && rootOrDoc.nodeType === 9 ? rootOrDoc : rootOrDoc && rootOrDoc.ownerDocument;
		return new XMLSerializer().serializeToString(doc || rootOrDoc);
	}

	function postXml(url, root) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.setRequestHeader("Content-Type", "text/xml");
		xhr.send(serializeXml(root));
		return xhr.responseText || "";
	}

	function text(id, value) {
		var item = document.getElementById(id);
		if (item) {
			item.textContent = value == null ? "" : String(value);
		}
	}

	function show(id, visible) {
		var item = document.getElementById(id);
		if (item) {
			item.style.display = visible ? "block" : "none";
		}
	}

	function ClearTable() {
		var table = document.getElementById("tabNoOfLevel");
		if (!table) {
			return;
		}
		while (table.rows.length > 1) {
			table.deleteRow(1);
		}
	}

	function labelForLevel(levelNo) {
		var found = "";
		childElements(xmlRoot("PackingData"), "SubLevelDetails").some(function (node) {
			if (trim(attr(node, "LevelNo")) === String(levelNo)) {
				found = attr(node, "LevelLabel");
				return true;
			}
			return false;
		});
		return found;
	}

	function AddSubLevel() {
		var count = Number(valueOf("txtNoOfSubLevel")) || 0;
		var table = document.getElementById("tabNoOfLevel");
		var row;
		var cell;
		var input;
		ClearTable();
		if (count < 0) {
			alert("Enter Valid Number");
			return false;
		}
		for (var i = 1; i <= count; i += 1) {
			row = table.insertRow(-1);
			cell = row.insertCell(-1);
			cell.textContent = i;
			cell.className = "ExcelHeaderCell";
			cell.align = "center";
			cell = row.insertCell(-1);
			cell.className = "ExcelDisplayCell";
			input = document.createElement("input");
			input.type = "text";
			input.name = "txtLevelLable" + i;
			input.value = labelForLevel(i);
			input.className = "FormElem";
			cell.appendChild(input);
		}
		return false;
	}

	function PopSample() {
		var count = Number(valueOf("txtNoOfSubLevel")) || 0;
		var lastLabel = "";
		show("divPreview", true);
		text("tdGross", valueOf("txtGrossPack"));
		if (checked("chkTare")) {
			text("tdTare", valueOf("txtTarePack"));
			show("divTare", true);
		} else {
			text("tdTare", "");
			show("divTare", false);
		}
		for (var i = 1; i <= count; i += 1) {
			lastLabel = valueOf("txtLevelLable" + i);
		}
		show("divCone", trim(lastLabel) !== "");
		text("spanSellingForm", lastLabel);
		return false;
	}

	function ViewInvoice() {
		var value = valueOf("hApplnTransNo");
		if (value && value !== "0" && window.ITMSModernCompat) {
			window.ITMSModernCompat.openModalDialog("../../Purchase/Transaction/RepPurInvoiceDetailspopup.asp?iInvNo=" + encodeURIComponent(value), "", "dialogHeight:470px;dialogWidth:870px;center:Yes;help:No;resizable:No;status:No", function () {});
		}
		return false;
	}

	function saveSubLevels(root) {
		var count = Number(valueOf("txtNoOfSubLevel")) || 0;
		var doc = xmlDocument("PackingData");
		var node;
		clearNamedChildren(root, "SubLevelDetails");
		for (var i = 1; i <= count; i += 1) {
			node = doc.createElement("SubLevelDetails");
			node.setAttribute("LevelNo", i);
			node.setAttribute("LevelLabel", valueOf("txtLevelLable" + i));
			root.appendChild(node);
		}
	}

	function Save() {
		var root = xmlRoot("PackingData");
		var response;
		if (!trim(valueOf("txtShortName"))) {
			alert("Enter Short Name");
			field("txtShortName").focus();
			return false;
		}
		if (!trim(valueOf("txtName"))) {
			alert("Enter Packing Type Name");
			field("txtName").focus();
			return false;
		}
		root.setAttribute("Type", valueOf("hType"));
		root.setAttribute("PackCode", valueOf("hPackCode"));
		root.setAttribute("ShortName", valueOf("txtShortName"));
		root.setAttribute("Name", valueOf("txtName"));
		root.setAttribute("AltLabel", valueOf("txtLabel"));
		root.setAttribute("ShowBothChk", checked("chkShowBoth") ? "Y" : "N");
		root.setAttribute("ReceiptNumbering", radioValue("RadRecNo"));
		root.setAttribute("LotNoSelection", radioValue("RadLotNo"));
		root.setAttribute("LotNoEnforceCheck", checked("chkLotNo") ? "Y" : "N");
		root.setAttribute("SerialNoSelection", radioValue("RadSerNo"));
		root.setAttribute("SerialNoWithinLotCheck", checked("chkSerNo") ? "Y" : "N");
		root.setAttribute("NoOfSubLevel", valueOf("txtNoOfSubLevel"));
		root.setAttribute("GrossLabel", valueOf("txtGrossPack"));
		root.setAttribute("Tare", checked("chkTare") ? "Y" : "N");
		root.setAttribute("TareLabel", valueOf("txtTarePack"));
		saveSubLevels(root);
		response = trim(postXml("PackingTypeInsert.asp", root));
		if (response) {
			alert(response);
		} else {
			alert("Packing Type Inserted Successfully");
			if (window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.returnAndClose) {
				window.ITMSModalReturnCompat.returnAndClose("Done");
			} else {
				window.close();
			}
		}
		return false;
	}

	function ShowData(type) {
		var root = xmlRoot("PackingData");
		if (type !== "E" || !root) {
			return false;
		}
		setValue("txtShortName", attr(root, "ShortName"));
		setValue("txtName", attr(root, "Name"));
		setValue("txtLabel", attr(root, "AltLabel"));
		setChecked("chkShowBoth", attr(root, "ShowBothChk") === "Y");
		setRadioValue("RadRecNo", attr(root, "ReceiptNumbering"), "N");
		setRadioValue("RadLotNo", attr(root, "LotNoSelection"), "A");
		setChecked("chkLotNo", attr(root, "LotNoEnforceCheck") === "Y");
		setRadioValue("RadSerNo", attr(root, "SerialNoSelection"), "A");
		setChecked("chkSerNo", attr(root, "SerialNoWithinLotCheck") === "Y");
		setValue("txtNoOfSubLevel", attr(root, "NoOfSubLevel"));
		if (valueOf("txtNoOfSubLevel") !== "0") {
			AddSubLevel();
		}
		setValue("txtGrossPack", attr(root, "GrossLabel"));
		setChecked("chkTare", attr(root, "Tare") === "Y");
		setValue("txtTarePack", attr(root, "TareLabel"));
		return false;
	}

	window.PopSample = PopSample;
	window.AddSubLevel = AddSubLevel;
	window.ClearTable = ClearTable;
	window.ViewInvoice = ViewInvoice;
	window.Save = Save;
	window.ShowData = ShowData;
}(window, document));
