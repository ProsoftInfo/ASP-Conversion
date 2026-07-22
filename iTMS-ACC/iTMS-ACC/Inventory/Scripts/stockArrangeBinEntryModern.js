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

	function removeChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function setText(id, value) {
		var element = byId(id);
		if (element) {
			element.textContent = value || "";
			element.innerText = value || "";
		}
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

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function clearTable() {
		var table = byId("tblData");
		rows = [];
		if (!table) {
			return;
		}
		while (table.rows.length > 1) {
			table.deleteRow(1);
		}
	}

	function clearXml(name) {
		var root = xmlRoot(xmlIsland(name));
		removeChildren(root);
	}

	function sanitizeDialogPart(value) {
		return trim(value).replace(/&/g, "AND").replace(/:/g, " - ").replace(/'/g, "~~").replace(/"/g, "``");
	}

	function dialogPayload(item, loc) {
		return [
			getAttr(item, "ICode"),
			getAttr(item, "CCode"),
			sanitizeDialogPart(getAttr(item, "IName")),
			sanitizeDialogPart(getAttr(loc, "LOCNAME")),
			getAttr(loc, "STOCK"),
			getAttr(loc, "LOC"),
			getAttr(loc, "BINSTATUS")
		].join(":");
	}

	function findItem(itemCode, classCode) {
		var root = xmlRoot(xmlIsland("OutData"));
		var items = elementChildren(root, "Item");
		for (var i = 0; i < items.length; i += 1) {
			if (getAttr(items[i], "ICode") === itemCode && getAttr(items[i], "CCode") === classCode) {
				return items[i];
			}
		}
		return null;
	}

	function prepareDialogData(itemCode, classCode) {
		var newRoot = loadXmlIsland("NewData", "<ROOT/>");
		var newDoc = xmlDocument(xmlIsland("NewData")) || newRoot.ownerDocument;
		var item = findItem(itemCode, classCode);
		var cloned;
		if (!newRoot || !item) {
			return xmlIsland("NewData");
		}
		cloned = newDoc.importNode ? newDoc.importNode(item, true) : item.cloneNode(true);
		newRoot.appendChild(cloned);
		return xmlIsland("NewData");
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return null;
		}
		return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function renderRow(table, item, loc) {
		var row = table.insertRow(table.rows.length);
		var stockInput;
		var link;
		var icon;
		var payload = dialogPayload(item, loc);

		rows.push({ item: item, loc: loc, payload: payload });
		appendCell(row, "ExcelSerial", "center", rows.length);
		appendCell(row, "ExcelDisplayCell", "left", getAttr(item, "IName"));
		appendCell(row, "ExcelDisplayCell", "left", getAttr(loc, "LOCNAME"));

		stockInput = makeInput("txtStk" + rows.length, getAttr(loc, "STOCK"), "FormelemRead", 12, true);
		appendCell(row, "ExcelDisplayCell", null, stockInput);

		link = document.createElement("a");
		link.href = "#";
		link.onclick = function () {
			return false;
		};

		icon = document.createElement("img");
		icon.name = "btn";
		icon.border = "0";
		icon.src = "../../assets/images/iTMS%20Icons/Entry.gif";
		icon.width = 15;
		icon.height = 15;
		icon.alt = "Arrange Bin";
		icon.value = payload;
		icon.setAttribute("value", payload);
		icon.onclick = function () {
			return window.BinStockDet(this);
		};
		link.appendChild(icon);
		appendCell(row, "ExcelDisplayCell", "center", link);
	}

	function updateUomText() {
		var root = xmlRoot(xmlIsland("OutData"));
		var item = elementChildren(root, "Item")[0];
		var uom = item && elementChildren(item, "UOM")[0];
		setText("idUoM", uom ? getAttr(uom, "UoMName") + " " : "");
	}

	window.FnInit = function () {
		window.GetXML();
		return false;
	};

	window.GetXML = function () {
		var xhr;
		clearTable();
		clearXml("OutData");
		xhr = syncGet("ArrangeBinDetXML.asp");
		if (trim(xhr.responseText)) {
			loadXmlIsland("OutData", xhr.responseText);
			window.DisplayDetails();
			window.popClaDisplay();
		}
		return false;
	};

	window.DisplayDetails = function () {
		var root = xmlRoot(xmlIsland("OutData"));
		var table = byId("tblData");
		clearTable();
		if (!root || !table) {
			return false;
		}
		elementChildren(root, "Item").forEach(function (item) {
			elementChildren(item, "LOCDET").forEach(function (loc) {
				renderRow(table, item, loc);
			});
		});
		return false;
	};

	window.BinStockDet = function (source) {
		var payload = trim(source && (source.value || source.getAttribute("value")));
		var parts = payload.split(":");
		var itemCode = trim(parts[0]);
		var classCode = trim(parts[1]);
		var binStatus = trim(parts[6]);
		if (binStatus === "N") {
			alert("Create the Bin No for Selected Item");
			return false;
		}
		openDialog("ArrangeBinDetailsEntry.asp?Data=" + encodeURIComponent(payload), prepareDialogData(itemCode, classCode), "dialogHeight:350px;dialogWidth:400px;center:Yes;status:No", function (value) {
			if (value === "Done") {
				window.GetXML();
			}
		});
		return false;
	};

	window.popClaDisplay = function () {
		updateUomText();
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

	window.CheckLot = function () {
		return false;
	};

	window.CheckQty = function (stockQty, rowNo) {
		var qty = field("txtQtyA" + rowNo);
		if (qty && toNumber(qty.value) > toNumber(stockQty)) {
			alert("Enter Issue Qty which is less than Stock Qty");
			qty.value = "0";
		}
		return false;
	};

	window.CheckSubmit = function () {
		return false;
	};

	window.Submit = function () {
		var frm = form();
		if (frm) {
			frm.action = "../MASTER/ITEMLISTENTRY.ASP?ACTN=M";
			frm.submit();
		}
		return false;
	};
}(window, document));
