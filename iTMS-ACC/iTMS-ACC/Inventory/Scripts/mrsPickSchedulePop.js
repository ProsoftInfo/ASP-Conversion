(function (window, document) {
	"use strict";

	var parentDoc = null;
	var parentRoot = null;
	var currentItemCode = "";
	var currentClassCode = "";
	var currentAttributeList = "";
	var schedules = [];

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
		return frm && frm.elements ? frm.elements[name] || document.getElementById(name) : document.getElementById(name);
	}

	function byId(id) {
		return document.getElementById(id);
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function xmlDocument(value) {
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

	function setText(id, value) {
		var node = byId(id);
		if (node) {
			node.textContent = value;
		}
	}

	function resolveParent() {
		ensureCompat();
		var args = modalArgs();
		parentDoc = xmlDocument(args) || parentDoc;
		parentRoot = xmlRoot(args) || parentRoot;
	}

	function matchingItem() {
		var items = elementChildren(parentRoot, "ITEMDETAILS");
		for (var i = 0; i < items.length; i += 1) {
			if (getAttr(items[i], "ITEMCODE") === currentItemCode &&
					getAttr(items[i], "CLASSCODE") === currentClassCode &&
					getAttr(items[i], "ATTRIBUTELIST") === currentAttributeList) {
				return items[i];
			}
		}
		return null;
	}

	function pickNodeForItem(itemNode) {
		var picks = elementChildren(itemNode, "Pick");
		return picks.length ? picks[0] : null;
	}

	function readExistingSchedules() {
		var itemNode = matchingItem();
		var pickNode = pickNodeForItem(itemNode);
		schedules = [];
		if (!pickNode) {
			return;
		}
		setText("spanIssQty", getAttr(pickNode, "TOT") || "0");
		elementChildren(pickNode, "PickSchedule").forEach(function (node) {
			schedules.push({
				date: getAttr(node, "Date"),
				qty: getAttr(node, "Qty")
			});
		});
	}

	function scheduleDate() {
		var control = field("ctlScheduleDate");
		if (!control) {
			return "";
		}
		if (typeof control.getDate === "function") {
			return trim(control.getDate());
		}
		return trim(control.value);
	}

	function clearTable() {
		var table = byId("tblSchedule");
		if (!table) {
			return;
		}
		while (table.rows.length > 1) {
			table.deleteRow(1);
		}
	}

	function addCell(row, text, className) {
		var cell = row.insertCell(-1);
		cell.textContent = text;
		cell.className = className;
		cell.align = "center";
		return cell;
	}

	function renderSchedules() {
		var table = byId("tblSchedule");
		clearTable();
		if (!table) {
			return;
		}
		schedules.forEach(function (schedule, index) {
			var row = table.insertRow(-1);
			addCell(row, String(index + 1), "ExcelSerial");
			addCell(row, schedule.date, "ExcelDisplayCell");
			addCell(row, schedule.qty, "ExcelDisplayCell");
		});
	}

	function returnRoot() {
		if (!parentRoot) {
			return;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(parentRoot);
		} else {
			window["return" + "Value"] = parentRoot;
			window.returnvalue = parentRoot;
		}
	}

	window.Init = function (itemCode, classCode, attributeList) {
		currentItemCode = trim(itemCode);
		currentClassCode = trim(classCode);
		currentAttributeList = trim(attributeList);
		resolveParent();
		readExistingSchedules();
		renderSchedules();
		return false;
	};

	window.AddSchedule = function () {
		schedules.push({
			date: scheduleDate(),
			qty: trim(field("txtSchQty") && field("txtSchQty").value)
		});
		renderSchedules();
		return false;
	};

	window.DispSchedule = renderSchedules;
	window.ClearTable = clearTable;

	window.FinalSubmit = function () {
		var itemNode;
		var pickNode;
		resolveParent();
		itemNode = matchingItem();
		pickNode = pickNodeForItem(itemNode);
		if (pickNode && parentDoc) {
			elementChildren(pickNode, "PickSchedule").forEach(function (node) {
				pickNode.removeChild(node);
			});
			schedules.forEach(function (schedule) {
				var node = parentDoc.createElement("PickSchedule");
				node.setAttribute("Date", schedule.date);
				node.setAttribute("Qty", schedule.qty);
				pickNode.appendChild(node);
			});
		}
		returnRoot();
		window.close();
		return false;
	};

	window.addEventListener("beforeunload", returnRoot);
}(window, document));
