(function (window, document) {
	"use strict";

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

	function clearTable() {
		var table = byId("tblData");
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

	function currentItemType(root) {
		var items = elementChildren(root, "Item");
		return items.length ? getAttr(items[0], "ItemTypeID") : "";
	}

	function updateHeaderText(root) {
		var itemNames = [];
		var uomName = "";
		elementChildren(root, "Item").forEach(function (item) {
			var name = getAttr(item, "IName");
			var uom = elementChildren(item, "UOM")[0];
			if (name) {
				itemNames.push(name);
			}
			if (!uomName && uom) {
				uomName = getAttr(uom, "UoMName");
			}
		});
		setText("idItem", itemNames.join(","));
		setText("idUoM", uomName ? uomName + " " : "");
	}

	function renderAttribute(table, attribute, rowNo, itemTypeId) {
		var row = table.insertRow(table.rows.length);
		var link;
		var icon;
		var optValue = getAttr(attribute, "OptValue");

		appendCell(row, "ExcelSerial", "center", rowNo).width = "50";
		appendCell(row, "ExcelDisplayCell", "left", getAttr(attribute, "OptName"));

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
		icon.alt = "Lot Details";
		icon.value = optValue;
		icon.setAttribute("value", optValue);
		icon.onclick = function () {
			return window.LotDetails(this, itemTypeId);
		};
		link.appendChild(icon);
		appendCell(row, "ExcelDisplayCell", "center", link).width = "100";
	}

	window.FnInit = function () {
		window.GetXML();
		return false;
	};

	window.GetXML = function () {
		var xhr;
		clearTable();
		clearXml("OutData");
		xhr = syncGet("AttributesDetXML.asp");
		if (trim(xhr.responseText)) {
			loadXmlIsland("OutData", xhr.responseText);
			window.DisplayDetails();
		}
		return false;
	};

	window.DisplayDetails = function () {
		var root = xmlRoot(xmlIsland("OutData"));
		var table = byId("tblData");
		var attrRoot;
		var attrs;
		var itemTypeId;
		clearTable();
		if (!root || !table) {
			return false;
		}
		updateHeaderText(root);
		itemTypeId = currentItemType(root);
		attrRoot = elementChildren(root, "AttributeDet")[0];
		attrs = attrRoot ? elementChildren(attrRoot, "Attribute") : [];
		for (var i = 0; i < attrs.length; i += 1) {
			renderAttribute(table, attrs[i], i + 1, itemTypeId);
		}
		return false;
	};

	window.LotDetails = function (source, itemTypeId) {
		var optValue = trim(source && (source.value || source.getAttribute("value")));
		openDialog("AttributeWiseLotDetailSelection.asp?Data=" + encodeURIComponent(optValue) + "&sType=" + encodeURIComponent(itemTypeId || ""), xmlIsland("OutData"), "dialogHeight:310px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No", function () {});
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
		var button = field("B7");
		var xhr;
		if (button) {
			button.disabled = true;
		}
		xhr = syncPost("stkMgmtAttWiseLotDetInsert.asp", serializeXml(xmlIsland("OutData")));
		if (trim(xhr.responseText)) {
			alert(xhr.responseText);
			if (button) {
				button.disabled = false;
			}
		} else {
			alert("Attribute Wise Lot Details Allocated");
			window.location.href = "../MASTER/ITEMLISTENTRY.ASP?ACTN=M";
		}
		return false;
	};

	window.Back = function () {
		window.location.href = "../MASTER/ITEMLISTENTRY.ASP?ACTN=M";
		return false;
	};
}(window, document));
