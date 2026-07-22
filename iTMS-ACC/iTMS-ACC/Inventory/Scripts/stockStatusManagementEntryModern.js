(function (window, document) {
	"use strict";

	var rows = [];

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

	function normalizeLot(value) {
		var lot = trim(value);
		return lot === "" || lot === "-" || lot === "0" || lot.toUpperCase() === "N/A" || lot.toUpperCase() === "NULL" ? "NULL" : lot;
	}

	function xmlIsland(name) {
		ensureCompat();
		return window[name] || document[name] || byId(name);
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

	function serializeXml(value) {
		var doc = xmlDocument(value);
		var root = xmlRoot(value);
		if (value && typeof value.xml === "string") {
			return value.xml;
		}
		if (doc && typeof doc.xml === "string") {
			return doc.xml;
		}
		if (!doc && !root) {
			return "";
		}
		return new XMLSerializer().serializeToString(doc || root);
	}

	function loadXmlIsland(name, text) {
		var island = xmlIsland(name);
		if (island && typeof island.loadXML === "function") {
			island.loadXML(text || "<Root/>");
			return xmlRoot(island);
		}
		return new DOMParser().parseFromString(text || "<Root/>", "text/xml").documentElement;
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

	function firstStatus(pick) {
		return elementChildren(pick, "STATUS")[0] || null;
	}

	function statusAttr(pick, name) {
		var status = firstStatus(pick);
		return status ? getAttr(status, name) : "";
	}

	function allItems(root) {
		return elementChildren(root, "Item");
	}

	function allPicks(root) {
		var result = [];
		allItems(root).forEach(function (item) {
			elementChildren(item, "LOCDET").forEach(function (loc) {
				elementChildren(loc, "PICK").forEach(function (pick) {
					result.push({ item: item, loc: loc, pick: pick });
				});
			});
		});
		return result;
	}

	function uomNode(item) {
		return elementChildren(item, "UOM")[0] || null;
	}

	function uomDecimal(item) {
		var node = uomNode(item);
		return node ? getAttr(node, "UoMDecimal") : "3";
	}

	function uomName(item) {
		var node = uomNode(item);
		return node ? getAttr(node, "UoMName") : "";
	}

	function samePick(pick, token) {
		var parts = String(token || "").split("AAAA");
		return normalizeLot(getAttr(pick, "LOTNO")) === normalizeLot(parts[4]) &&
			trim(getAttr(pick, "INVRECNO")) === trim(parts[8] || "") &&
			trim(getAttr(pick.parentNode, "LOC")) === trim(parts[6] || "") &&
			trim(getAttr(pick.parentNode, "BIN")) === trim(parts[7] || "");
	}

	function findPickByToken(token) {
		var root = xmlRoot(xmlIsland("OutData"));
		var found = null;
		allPicks(root).some(function (entry) {
			if (samePick(entry.pick, token)) {
				found = entry;
				return true;
			}
			return false;
		});
		return found;
	}

	function findMgmtPick(root, row) {
		var found = null;
		var lot = normalizeLot(field("txtLot" + row.index) && field("txtLot" + row.index).value);
		var loc = getAttr(row.loc, "LOC");
		var bin = getAttr(row.loc, "BIN");
		elementChildren(root).some(function walk(node) {
			if (String(node.nodeName).toLowerCase() === "pick" &&
					getAttr(node, "LOC") === loc &&
					getAttr(node, "BIN") === bin &&
					normalizeLot(getAttr(node, "LOTNO")) === lot) {
				found = node;
				return true;
			}
			return elementChildren(node).some(walk);
		});
		return found;
	}

	function clearXml(name) {
		var root = xmlRoot(xmlIsland(name));
		while (root && root.firstChild) {
			root.removeChild(root.firstChild);
		}
	}

	function clearTable() {
		var table = byId("tblData");
		rows = [];
		if (!table) {
			return;
		}
		while (table.rows.length > 2) {
			table.deleteRow(2);
		}
	}

	function setText(id, value) {
		var element = byId(id);
		if (element) {
			element.textContent = value || "";
			element.innerText = value || "";
		}
	}

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function syncPost(url, xmlText) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.setRequestHeader("Content-Type", "text/xml");
		xhr.send(xmlText || "");
		return xhr;
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return null;
		}
		return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function appendCell(row, className, align, content) {
		var cell = row.insertCell(row.cells.length);
		cell.className = className || "";
		if (align) {
			cell.align = align;
		}
		if (content == null) {
			return cell;
		}
		if (typeof content === "string" || typeof content === "number") {
			cell.textContent = String(content);
		} else {
			cell.appendChild(content);
		}
		return cell;
	}

	function makeInput(name, value, className, size, readOnly) {
		var input = document.createElement("input");
		input.type = "text";
		input.name = name;
		input.value = value == null ? "" : String(value);
		input.className = className || "FormElem";
		if (size) {
			input.size = size;
		}
		if (readOnly) {
			input.readOnly = true;
		}
		input.style.textAlign = "right";
		return input;
	}

	function ensureHidden(name, value) {
		var input = field(name);
		if (!input) {
			input = document.createElement("input");
			input.type = "hidden";
			input.name = name;
			form().appendChild(input);
		}
		input.value = value == null ? "" : String(value);
		return input;
	}

	function tokenFor(row) {
		return [
			"btn",
			getAttr(row.item, "Unit") || (field("hOrgID") && field("hOrgID").value),
			getAttr(row.item, "ICode"),
			getAttr(row.item, "CCode"),
			getAttr(row.pick, "LOTNO"),
			getAttr(row.pick, "QTYSTK"),
			getAttr(row.loc, "LOC"),
			getAttr(row.loc, "BIN") || "0",
			getAttr(row.pick, "INVRECNO"),
			statusAttr(row.pick, "STOCKNO") || "0",
			row.uomDecimal
		].join("AAAA");
	}

	function resetPickTotals(root) {
		allPicks(root).forEach(function (entry) {
			setAttr(entry.pick, "QTYCLE", "0");
			setAttr(entry.pick, "QTYREJ", "0");
			setAttr(entry.pick, "QTYOHO", "0");
			setAttr(entry.pick, "QTYRES", "0");
			setAttr(entry.pick, "QTY", "0");
		});
	}

	function existingReservedQty(pick) {
		var total = 0;
		elementChildren(pick).forEach(function (child) {
			elementChildren(child, "RESERVEDETAILS").forEach(function (reserve) {
				total += toNumber(getAttr(reserve, "QTY")) || toNumber(getAttr(reserve, "QTYRES")) || toNumber(getAttr(reserve, "RESQTY"));
				total -= toNumber(getAttr(reserve, "QTYISS")) || toNumber(getAttr(reserve, "ISSQTY"));
			});
		});
		elementChildren(pick, "RESERVEDETAILS").forEach(function (reserve) {
			total += toNumber(getAttr(reserve, "QTY")) || toNumber(getAttr(reserve, "QTYRES")) || toNumber(getAttr(reserve, "RESQTY"));
			total -= toNumber(getAttr(reserve, "QTYISS")) || toNumber(getAttr(reserve, "ISSQTY"));
		});
		return total;
	}

	function renderPick(table, entry) {
		var rowInfo;
		var row;
		var index;
		var lotNo = getAttr(entry.pick, "LOTNO");
		var stockNo = statusAttr(entry.pick, "STOCKNO") || "0";
		var recNum = getAttr(entry.loc, "RECNUM");
		var serialIcon;
		var link;
		var rejected;
		var reserved;
		var token;

		rowInfo = {
			index: rows.length + 1,
			item: entry.item,
			loc: entry.loc,
			pick: entry.pick,
			uomDecimal: uomDecimal(entry.item),
			stockInput: null,
			rejectedInput: null,
			onHoldInput: null,
			reservedInput: null
		};
		rows.push(rowInfo);
		index = rowInfo.index;
		row = table.insertRow(table.rows.length);

		appendCell(row, "ExcelSerial", "center", index);
		appendCell(row, "ExcelDisplayCell", null, getAttr(entry.item, "IName"));
		appendCell(row, "ExcelDisplayCell", null, getAttr(entry.loc, "LOCNAME"));
		appendCell(row, "ExcelDisplayCell", "left", makeInput("txtLot" + index, normalizeLot(lotNo) === "NULL" ? "-" : lotNo, "FormelemRead", 30, true)).width = "150";

		rowInfo.stockInput = makeInput("txtStkA" + index, getAttr(entry.pick, "QTYSTK"), "FormelemRead", 12, true);
		appendCell(row, "ExcelDisplayCell", null, rowInfo.stockInput).width = "10";

		token = tokenFor(rowInfo);
		ensureHidden("hA" + index, token);

		link = document.createElement("a");
		link.href = "#";
		if (recNum === "S" || recNum === "LS") {
			serialIcon = document.createElement("img");
			serialIcon.name = "btnA" + index;
			serialIcon.value = String(index);
			serialIcon.border = "0";
			serialIcon.src = "../../assets/images/iTMS%20Icons/Entry.gif";
			serialIcon.width = 15;
			serialIcon.height = 15;
			serialIcon.alt = "Serial Details";
			serialIcon.onclick = function () {
				return window.SelPopUp(index);
			};
			link.appendChild(serialIcon);
		}
		rejected = makeInput("txtRejA" + index, "0", "Formelem", 12, normalizeLot(lotNo) !== "NULL");
		rejected.onkeypress = function () {
			if (typeof window.DoKeyPress === "function") {
				window.DoKeyPress(rowInfo.uomDecimal, 7, 3);
			}
		};
		rowInfo.rejectedInput = rejected;
		appendCell(row, "ExcelDisplayCell", "center", link).appendChild(rejected);

		rowInfo.onHoldInput = makeInput("txtOnHA" + index, statusAttr(entry.pick, "QTYOHO") || "0", "Formelem", 12, false);
		rowInfo.onHoldInput.onkeypress = function () {
			if (typeof window.DoKeyPress === "function") {
				window.DoKeyPress(rowInfo.uomDecimal, 7, 3);
			}
		};
		appendCell(row, "ExcelInputCell", null, rowInfo.onHoldInput).width = "10";

		reserved = makeInput("txtResA" + index, statusAttr(entry.pick, "QTYRES") || "0", "Formelem", 12, false);
		reserved.onkeypress = function () {
			if (typeof window.DoKeyPress === "function") {
				window.DoKeyPress(rowInfo.uomDecimal, 7, 3);
			}
		};
		if (toNumber(reserved.value) === 0 || normalizeLot(lotNo) !== "NULL") {
			reserved.readOnly = true;
		}
		if (toNumber(reserved.value) !== 0 && normalizeLot(lotNo) !== "NULL") {
			reserved.className = "FormelemRead";
			reserved.style.cursor = "pointer";
			reserved.style.fontWeight = "bold";
			reserved.onclick = function () {
				return window.CheckUn(this);
			};
		}
		rowInfo.reservedInput = reserved;
		appendCell(row, reserved.readOnly ? "ExcelDisplayCell" : "ExcelInputCell", null, reserved).width = "10";
		appendCell(row, "ExcelDisplayCell", null, uomName(entry.item));

		if (field("hRcptNumbering")) {
			field("hRcptNumbering").value = recNum;
		}
		if (field("hLoc")) {
			field("hLoc").value = getAttr(entry.loc, "LOC");
		}
		if (field("hBin")) {
			field("hBin").value = getAttr(entry.loc, "BIN");
		}
		if (field("hItem")) {
			field("hItem").value = getAttr(entry.item, "ICode");
		}
		stockNo = stockNo;
	}

	function renderRows(storeValue) {
		var root = xmlRoot(xmlIsland("OutData"));
		var table = byId("tblData");
		var selected = String(storeValue || "");
		var loc = "";
		var bin = "";
		var parts;
		if (!table || !root) {
			return false;
		}
		if (selected && selected !== "select") {
			parts = selected.split(":");
			loc = parts[0] || "";
			bin = parts.length > 1 ? parts[1] : "0";
		}
		clearTable();
		resetPickTotals(root);
		allPicks(root).forEach(function (entry) {
			if (loc && (getAttr(entry.loc, "LOC") !== loc || getAttr(entry.loc, "BIN") !== bin)) {
				return;
			}
			renderPick(table, entry);
		});
		return false;
	}

	function loadSerialStockXml(row) {
		var url = "itmMgmtXMLSelect.asp?orgID=" + encodeURIComponent(getAttr(row.item, "Unit") || (field("hOrgID") && field("hOrgID").value) || "") +
			"&iClass=" + encodeURIComponent(getAttr(row.item, "CCode")) +
			"&iItem=" + encodeURIComponent(getAttr(row.item, "ICode"));
		var xhr = syncGet(url);
		if (trim(xhr.responseText)) {
			loadXmlIsland("OutData1", xhr.responseText);
		}
	}

	function updateFromStatusPopup(token) {
		var found = findPickByToken(token);
		var index;
		if (!found) {
			return;
		}
		index = rows.findIndex(function (row) {
			return row.pick === found.pick;
		}) + 1;
		if (!index) {
			return;
		}
		if (field("txtRejA" + index)) {
			field("txtRejA" + index).value = getAttr(found.pick, "QTYREJ") || statusAttr(found.pick, "QTYREJ") || "0";
		}
		if (field("txtOnHA" + index)) {
			field("txtOnHA" + index).value = getAttr(found.pick, "QTYOHO") || statusAttr(found.pick, "QTYOHO") || "0";
		}
		if (field("txtResA" + index)) {
			field("txtResA" + index).value = getAttr(found.pick, "QTYRES") || statusAttr(found.pick, "QTYRES") || "0";
		}
	}

	function validateRows() {
		var valid = true;
		rows.some(function (row) {
			var statusQty;
			var reservedBase;
			if (trim(row.rejectedInput.value) === "") {
				alert("Enter Rejected Quantity");
				row.rejectedInput.select();
				valid = false;
				return true;
			}
			if (trim(row.onHoldInput.value) === "") {
				alert("Enter On Hold Quantity");
				row.onHoldInput.select();
				valid = false;
				return true;
			}
			if (trim(row.reservedInput.value) === "") {
				alert("Enter Reserved Quantity");
				row.reservedInput.select();
				valid = false;
				return true;
			}
			statusQty = toNumber(row.rejectedInput.value) + toNumber(row.onHoldInput.value) + toNumber(row.reservedInput.value);
			reservedBase = existingReservedQty(row.pick);
			if (statusQty > toNumber(row.stockInput.value)) {
				alert("Total Quantity should be equal to or less than Stock Quantity (" + row.stockInput.value + ")");
				valid = false;
				return true;
			}
			row.cleanQty = statusQty - reservedBase;
			return false;
		});
		return valid;
	}

	function applyRowsToXml() {
		rows.forEach(function (row) {
			setAttr(row.pick, "QTYCLE", row.cleanQty || 0);
			setAttr(row.pick, "QTYREJ", toNumber(row.rejectedInput.value));
			setAttr(row.pick, "QTYOHO", toNumber(row.onHoldInput.value));
			setAttr(row.pick, "QTYRES", toNumber(row.reservedInput.value));
		});
	}

	window.FnInit = function () {
		window.GetXML();
		window.DisplayDetails();
		return false;
	};

	window.GetXML = function () {
		var xhr;
		clearXml("OutData");
		xhr = syncGet("itmStatusXMLSelect.asp");
		if (trim(xhr.responseText)) {
			loadXmlIsland("OutData", xhr.responseText);
		}
		return false;
	};

	window.CheckLot = function (obj, value) {
		var token = obj && obj.value || obj && obj.name || "";
		var label = value == null ? "" : String(value);
		openDialog("stkMgmtSMPoP.asp?sTemp=" + encodeURIComponent(token) + "&sValue=" + encodeURIComponent(label), xmlIsland("OutData"), "dialogHeight:310px;dialogWidth:600px;center:Yes;help:No;resizable:No;status:No", function () {
			updateFromStatusPopup(token);
		});
		return false;
	};

	window.CheckLotUn = function (obj, value) {
		var token = obj && obj.value || obj && obj.name || "";
		var label = value == null ? "" : String(value);
		openDialog("stkMgmtSMUnPoP.asp?sTemp=" + encodeURIComponent(token) + "&sValue=" + encodeURIComponent(label), xmlIsland("OutData"), "dialogHeight:310px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No", function () {
			var found = findPickByToken(token);
			var total = found ? existingReservedQty(found.pick) : 0;
			var index = rows.findIndex(function (row) {
				return found && row.pick === found.pick;
			}) + 1;
			if (index && field("txtResA" + index)) {
				field("txtResA" + index).value = total;
			}
		});
		return false;
	};

	window.Check = function (obj) {
		var hidden;
		var value;
		if (!obj) {
			return false;
		}
		if (trim(obj.value) === "") {
			obj.value = obj.defaultValue || "0";
		}
		obj.defaultValue = obj.value;
		value = toNumber(obj.value);
		hidden = field(String(obj.name || "").replace("txtStk", "h"));
		if (value !== 0 && hidden) {
			window.CheckLot(hidden, value);
		}
		return false;
	};

	window.SelPop = function () {
		var current = rows.length ? rows[rows.length - 1] : null;
		if (current) {
			loadSerialStockXml(current);
		}
		return false;
	};

	window.SelPopUp = function (index) {
		var row = rows[toNumber(index) - 1];
		var lotField = field("txtLot" + index);
		var lot = lotField ? lotField.value : getAttr(row && row.pick, "LOTNO");
		var values;
		if (!row) {
			return false;
		}
		loadSerialStockXml(row);
		values = [
			lot,
			lot,
			getAttr(row.item, "CCode"),
			getAttr(row.item, "ICode"),
			getAttr(row.item, "Unit") || (field("hOrgID") && field("hOrgID").value) || "",
			getAttr(row.loc, "LOC"),
			getAttr(row.loc, "BIN") || "0"
		].join("`");
		openDialog("SerialWiseStockPoP.asp?sTemp=" + encodeURIComponent(values), xmlIsland("OutData1"), "dialogHeight:370px;dialogWidth:380px;center:Yes;help:No;resizable:No;status:No", function () {
			var pick = findMgmtPick(xmlRoot(xmlIsland("OutData1")), row);
			if (pick && field("txtRejA" + index)) {
				field("txtRejA" + index).value = getAttr(pick, "ISSQTY") || "0";
			}
		});
		return false;
	};
	window.selPopUp = window.SelPopUp;
	window.SelPopup = window.SelPopUp;

	window.CheckUn = function (obj) {
		var hidden;
		var value;
		if (!obj) {
			return false;
		}
		if (trim(obj.value) === "") {
			obj.value = obj.defaultValue || "0";
		}
		obj.defaultValue = obj.value;
		value = toNumber(obj.value);
		hidden = field(String(obj.name || "").replace("txtRes", "h"));
		if (value !== 0 && hidden) {
			window.CheckLotUn(hidden, value);
		}
		return false;
	};

	window.DisplayDetails = function (storeValue) {
		return renderRows(storeValue);
	};

	window.popClaDisplay = function () {
		var root = xmlRoot(xmlIsland("OutData"));
		var select = field("selStore");
		if (select) {
			select.options.length = 1;
		}
		allItems(root).forEach(function (item) {
			var uom = uomNode(item);
			elementChildren(item, "LOCDET").forEach(function (loc) {
				var text;
				if (field("hRcptNumbering")) {
					field("hRcptNumbering").value = getAttr(loc, "RECNUM");
				}
				if (select) {
					text = getAttr(loc, "BIN") !== "0" ? getAttr(loc, "LOCNAME") + " -- " + getAttr(loc, "BINNAME") : getAttr(loc, "LOCNAME");
					select.options[select.options.length] = new Option(text, getAttr(loc, "LOC") + (getAttr(loc, "BIN") !== "0" ? ":" + getAttr(loc, "BIN") : ""));
				}
			});
			if (uom) {
				setText("idUoM", getAttr(uom, "UoMName") + " ");
			}
		});
		return false;
	};

	window.clearXML = function () {
		clearXml("OutData");
		return false;
	};

	window.ClearTable = function () {
		clearTable();
		return false;
	};

	window.CheckSubmit = function () {
		var xhr;
		if (!rows.length) {
			return false;
		}
		if (!validateRows()) {
			return false;
		}
		applyRowsToXml();
		syncPost("SerialWiseStockXML.asp", serializeXml(xmlIsland("OutData1")));
		xhr = syncPost("stkMgmtSMInsert.asp", serializeXml(xmlIsland("OutData")));
		if (trim(xhr.responseText) === "") {
			alert("Status Management has done");
			window.location.href = "../master/ITEMLISTENTRY.ASP";
		} else {
			alert(xhr.responseText);
			window.location.href = "stkMgmtSMEntry.asp";
		}
		return false;
	};
}(window, document));
