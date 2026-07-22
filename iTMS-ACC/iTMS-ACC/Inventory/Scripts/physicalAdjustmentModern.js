(function (window, document) {
	"use strict";

	var rows = [];
	var attrNameCache = {};

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

	function today() {
		var date = new Date();
		var day = date.getDate() < 10 ? "0" + date.getDate() : String(date.getDate());
		var month = date.getMonth() + 1 < 10 ? "0" + (date.getMonth() + 1) : String(date.getMonth() + 1);
		return day + "/" + month + "/" + date.getFullYear();
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

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, String(value == null ? "" : value));
		}
	}

	function createElement(doc, name) {
		return doc.createElement(name);
	}

	function removeChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
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
		input.defaultValue = input.value;
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

	function makeHidden(name, value) {
		var input = document.createElement("input");
		input.type = "hidden";
		input.name = name;
		input.value = value == null ? "" : String(value);
		return input;
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

	function clearTable() {
		var table = byId("tblData");
		rows = [];
		if (!table) {
			return;
		}
		while (table.rows.length > 2) {
			table.deleteRow(2);
		}
		Array.prototype.slice.call(document.querySelectorAll("input[data-pa-row-hidden='1']")).forEach(function (node) {
			node.parentNode.removeChild(node);
		});
	}

	function clearXml(name) {
		var root = xmlRoot(xmlIsland(name));
		removeChildren(root);
	}

	function normalizeLot(value) {
		var lot = trim(value);
		return lot === "0" || lot === "" ? "NULL" : lot;
	}

	function getAttrName(attList) {
		var key = trim(attList);
		var xhr;
		if (!key) {
			return "";
		}
		if (Object.prototype.hasOwnProperty.call(attrNameCache, key)) {
			return attrNameCache[key];
		}
		xhr = syncGet("../../Include/GetAttrName.asp?AttID=" + encodeURIComponent(key.replace(/,/g, ":")));
		attrNameCache[key] = trim(xhr.responseText);
		return attrNameCache[key];
	}

	function rowToken(item, loc, pick, uomDecimal) {
		return [
			"btn",
			getAttr(item, "Unit"),
			getAttr(item, "ICode"),
			getAttr(item, "CCode"),
			getAttr(pick, "LOTNO"),
			getAttr(pick, "QTYSTK"),
			getAttr(loc, "LOC"),
			getAttr(loc, "BIN"),
			getAttr(pick, "INVRECNO"),
			uomDecimal || "3",
			getAttr(pick, "ATTLIST")
		].join("AAAA");
	}

	function findPickFromToken(token) {
		var parts = String(token || "").split("AAAA");
		var itemCode = parts[2] || "";
		var classCode = parts[3] || "";
		var lotNo = normalizeLot(parts[4] || "");
		var locNo = parts[6] || "";
		var binNo = parts[7] || "";
		var invRecNo = parts[8] || "";
		var found = null;
		elementChildren(xmlRoot(xmlIsland("OutData")), "Item").some(function (item) {
			if (getAttr(item, "ICode") !== itemCode || getAttr(item, "CCode") !== classCode) {
				return false;
			}
			return elementChildren(item, "LOCDET").some(function (loc) {
				if (getAttr(loc, "LOC") !== locNo || getAttr(loc, "BIN") !== binNo) {
					return false;
				}
				return elementChildren(loc, "PICK").some(function (pick) {
					if (normalizeLot(getAttr(pick, "LOTNO")) === lotNo && getAttr(pick, "INVRECNO") === invRecNo) {
						found = { item: item, loc: loc, pick: pick };
						return true;
					}
					return false;
				});
			});
		});
		return found;
	}

	function rowByQuantity(input) {
		var rowNo = Number(String(input && input.name || "").replace(/^txtQtyA/, ""));
		return rows[rowNo - 1] || null;
	}

	function updateAdjusted(row) {
		var qty = row.quantity;
		var stock = row.stock;
		var adjusted = toNumber(qty.value) - toNumber(stock.value);
		row.adjusted.value = adjusted;
		return adjusted;
	}

	function serialButton(row) {
		var link = document.createElement("a");
		var icon = document.createElement("img");
		link.href = "#";
		link.onclick = function () {
			return false;
		};
		icon.name = row.token;
		icon.border = "0";
		icon.src = "../../assets/images/iTMS%20Icons/Entry.gif";
		icon.width = 15;
		icon.height = 15;
		icon.alt = "Serial Details";
		icon.onclick = function () {
			return window.CheckSerial(this, row.index);
		};
		link.appendChild(icon);
		return link;
	}

	function renderRow(table, item, loc, pick, uomName, uomDecimal) {
		var row = table.insertRow(table.rows.length);
		var rowInfo;
		var lotText = getAttr(pick, "LOTNO") === "NULL" ? "-" : getAttr(pick, "LOTNO");
		var attName = getAttrName(getAttr(pick, "ATTLIST"));
		var hidden;
		rows.push(rowInfo = {
			index: rows.length + 1,
			item: item,
			loc: loc,
			pick: pick,
			token: rowToken(item, loc, pick, uomDecimal)
		});
		if (attName && lotText !== "-") {
			lotText += "[" + attName + "]";
		}
		appendCell(row, "ExcelSerial", "center", rowInfo.index);
		appendCell(row, "ExcelDisplayCell", null, getAttr(item, "IName")).appendChild(makeHidden("txtItemName" + rowInfo.index, getAttr(item, "IName")));
		appendCell(row, "ExcelDisplayCell", null, getAttr(loc, "LOCNAME")).appendChild(makeHidden("txtClassName" + rowInfo.index, getAttr(loc, "LOCNAME")));
		appendCell(row, "ExcelDisplayCell", "left", makeInput("txtLot" + rowInfo.index, lotText, "FormelemRead", 30, true)).width = "50";
		rowInfo.stock = makeInput("txtStk" + rowInfo.index, getAttr(pick, "QTYSTK"), "FormelemRead", 12, true);
		appendCell(row, "ExcelDisplayCell", null, rowInfo.stock).width = "10";
		rowInfo.quantity = makeInput("txtQtyA" + rowInfo.index, getAttr(pick, "QTYSTK"), "Formelem", 12, false);
		rowInfo.quantity.onkeypress = function () {
			if (typeof window.DoKeyPress === "function") {
				window.DoKeyPress(uomDecimal || "3", 7, 3);
			}
		};
		rowInfo.quantity.onblur = function () {
			return window.Check(this);
		};
		appendCell(row, "ExcelInputCell", null, rowInfo.quantity).width = "10";
		hidden = makeHidden("hA" + rowInfo.index, rowInfo.token);
		hidden.setAttribute("data-pa-row-hidden", "1");
		form().appendChild(hidden);
		rowInfo.hidden = hidden;
		appendCell(row, "ExcelDisplayCell", "center", serialButton(rowInfo)).width = "10";
		rowInfo.adjusted = makeInput("txtJusA" + rowInfo.index, "0", "Formelem", 12, true);
		appendCell(row, "ExcelInputCell", "center", rowInfo.adjusted).width = "10";
		rowInfo.reason = makeInput("txtReasonA" + rowInfo.index, "", "Formelem", 12, false);
		rowInfo.reason.maxLength = 100;
		appendCell(row, "ExcelInputCell", "center", rowInfo.reason).width = "10";
		appendCell(row, "ExcelDisplayCell", null, uomName || "");
	}

	function uomFor(item) {
		var uom = elementChildren(item, "UOM")[0];
		return {
			name: uom ? getAttr(uom, "UoMName") : "",
			decimal: uom ? getAttr(uom, "UoMDecimal") : "3",
			code: uom ? getAttr(uom, "UoMCode") : ""
		};
	}

	window.FnInit = function () {
		window.GetXML();
		window.DisplayDetails();
		return false;
	};

	window.GetXML = function () {
		var item = trim(field("hItem") && field("hItem").value);
		var classCode = trim(field("hClass") && field("hClass").value);
		var orgId = trim(field("hOrgID") && field("hOrgID").value);
		var xhr;
		clearXml("OutData");
		xhr = syncGet("itmStoreXMLSelect.asp?Data=" + encodeURIComponent(orgId + ":" + classCode + ":" + item));
		if (trim(xhr.responseText)) {
			loadXmlIsland("OutData", xhr.responseText);
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
			var uom = uomFor(item);
			elementChildren(item, "LOCDET").forEach(function (loc) {
				elementChildren(loc, "PICK").forEach(function (pick) {
					setAttr(pick, "QTYISS", "");
					setAttr(pick, "ADJUSTED", "");
					renderRow(table, item, loc, pick, uom.name, uom.decimal);
				});
			});
		});
		return false;
	};

	window.CheckLot = function (hidden, adjustedValue) {
		var token = hidden && hidden.value || "";
		var found = findPickFromToken(token);
		var parts = token.split("AAAA");
		var rowNo = Number(String(hidden && hidden.name || "").replace(/^hA/, ""));
		var row = rows[rowNo - 1] || null;
		var storeName = found ? getAttr(found.loc, "LOCNAME") : "";
		var receiptType = trim(field("hRcptNum") && field("hRcptNum").value);
		var uom = found ? uomFor(found.item) : { code: "", decimal: "" };
		if (receiptType === "N" || !found) {
			return false;
		}
		if (toNumber(adjustedValue) < 0) {
			openDialog("stkMgmtPAPoP.asp?sTemp=" + encodeURIComponent(token) + "&sValue=" + encodeURIComponent(" -- " + storeName.replace(/&/g, "and") + "`" + adjustedValue), xmlIsland("OutData"), "dialogHeight:400px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No", function () {
				if (row) {
					row.adjusted.value = updateAdjusted(row);
				}
			});
		} else {
			var tempValues = [
				receiptType,
				getAttr(found.item, "ICode"),
				getAttr(found.item, "CCode"),
				getAttr(found.item, "Unit"),
				getAttr(found.loc, "LOC"),
				getAttr(found.loc, "BIN"),
				adjustedValue,
				storeName.replace(/&/g, "and"),
				uom.code,
				"0",
				today(),
				"",
				"",
				getAttr(found.pick, "LOTNO"),
				getAttr(found.pick, "ATTLIST")
			].join("``");
			openDialog("../../Common/PackingLotSerForPA.asp?sTemp=" + encodeURIComponent(tempValues), xmlIsland("OutData"), "dialogHeight:580px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No", function () {});
		}
		return false;
	};

	window.CheckSerial = function (icon, rowNo) {
		var row = rows[Number(rowNo) - 1];
		if (row) {
			window.Check(row.quantity);
		}
		return false;
	};

	window.CheckStock = function (input) {
		return window.Check(input);
	};

	window.Check = function (input) {
		var row = rowByQuantity(input);
		var adjusted;
		if (!row) {
			return false;
		}
		if (trim(input.value) === "") {
			input.value = input.defaultValue;
		}
		input.defaultValue = input.value;
		adjusted = updateAdjusted(row);
		if (toNumber(adjusted) !== 0) {
			window.CheckLot(row.hidden, adjusted);
		}
		return false;
	};

	function validateRows() {
		for (var i = 0; i < rows.length; i += 1) {
			if (trim(rows[i].quantity.value) === "") {
				alert("Enter Quantity");
				rows[i].quantity.select();
				return false;
			}
			if (toNumber(rows[i].adjusted.value) > 0 && trim(rows[i].reason.value) === "") {
				alert("Enter the Reason");
				rows[i].reason.focus();
				return false;
			}
		}
		return rows.length > 0;
	}

	function applyRowsToOutData() {
		rows.forEach(function (row) {
			var adjusted = toNumber(row.stock.value) - toNumber(row.quantity.value);
			setAttr(row.item, "Reason", row.reason.value);
			setAttr(row.pick, "QTYISS", toNumber(row.quantity.value));
			setAttr(row.pick, "ADJUSTED", adjusted);
			row.adjusted.value = toNumber(row.quantity.value) - toNumber(row.stock.value);
		});
	}

	function resetIssueRoot(root) {
		removeChildren(root);
		setAttr(root, "ISSTYPE", "F");
		setAttr(root, "ISSTOCODE", "INV");
		setAttr(root, "ISSTOTYPE", "DEPT");
		setAttr(root, "ISSTOSUBCODE", "");
		setAttr(root, "POConfirm", "N");
		setAttr(root, "SInvConfirm", "N");
		setAttr(root, "Invoice", "A");
		setAttr(root, "GPConfirm", "N");
		setAttr(root, "ProConfirm", "N");
		setAttr(root, "MCallFrom", "N");
		setAttr(root, "RedirectTo", "");
		setAttr(root, "AppRefType", "");
		setAttr(root, "AppRefNo", "");
		setAttr(root, "AppRefDate", "");
		setAttr(root, "ConsumptionAccHead", "");
		setAttr(root, "IssueToCode", "");
		setAttr(root, "PickPackFlag", "");
		setAttr(root, "IssFrom", "IN");
		setAttr(root, "Returnable", "N");
		setAttr(root, "ReturnItem", "S");
		setAttr(root, "TYPE", "GEN");
	}

	function resetReceiptRoot(root) {
		removeChildren(root);
		setAttr(root, "DEPT", "OTH");
		setAttr(root, "SOURCE", "N");
		setAttr(root, "ORGCODE", field("hOrgID") ? field("hOrgID").value : "");
		setAttr(root, "STYPE", "N");
		setAttr(root, "ITEMTYPE", "");
		setAttr(root, "PACKNUM", "");
		setAttr(root, "SRCREFTYPE", "N");
		setAttr(root, "SRCREFNO", "");
		setAttr(root, "RCPTNUMBERINV", "");
		setAttr(root, "sTypeRcpt", "");
		setAttr(root, "APPREFTYPE", "");
		setAttr(root, "APPREFNO", "");
		setAttr(root, "APPREFDATE", "");
		setAttr(root, "RCVDON", today());
	}

	function buildIssueData() {
		var issueIsland = xmlIsland("IssueData");
		var issueDoc = xmlDocument(issueIsland);
		var issueRoot = xmlRoot(issueIsland);
		var entryNo = 0;
		resetIssueRoot(issueRoot);
		elementChildren(xmlRoot(xmlIsland("OutData")), "Item").forEach(function (item) {
			var issueItem = createElement(issueDoc, "ITEM");
			var issueTotal = 0;
			entryNo += 1;
			setAttr(issueItem, "ENTRYNO", entryNo);
			setAttr(issueItem, "ITMCODE", getAttr(item, "ICode"));
			setAttr(issueItem, "CLACODE", getAttr(item, "CCode"));
			setAttr(issueItem, "ITMNAME", getAttr(item, "IName"));
			setAttr(issueItem, "SSTORE", "");
			setAttr(issueItem, "REQQTY", "0");
			setAttr(issueItem, "REQBY", "");
			setAttr(issueItem, "REMARKS", "");
			setAttr(issueItem, "ITEMTYPE", "");
			setAttr(issueItem, "ISSUEDATE", today());
			setAttr(issueItem, "TRAQTY", "0");
			setAttr(issueItem, "PRQTY", "0");
			setAttr(issueItem, "IVALUE", "0");
			setAttr(issueItem, "ORGCODE", field("hOrgID") ? field("hOrgID").value : "");
			setAttr(issueItem, "MRSNO", "");
			setAttr(issueItem, "MRSDATE", "");
			setAttr(issueItem, "ATTRIBUTELIST", "");
			setAttr(issueItem, "CREATEDBY", field("hUserID") ? field("hUserID").value : "");
			setAttr(issueItem, "CREATEDON", today());
			setAttr(issueItem, "RETURNABLE", "1");
			setAttr(issueItem, "RefNo", "");
			setAttr(issueItem, "ONLYLOT", "");
			setAttr(issueItem, "RETURNITEM", "S");
			setAttr(issueItem, "MatType", "");
			issueRoot.appendChild(issueItem);
			elementChildren(item, "LOCDET").forEach(function (loc) {
				elementChildren(loc, "PICK").forEach(function (pick) {
					var adjusted = toNumber(getAttr(pick, "ADJUSTED"));
					var issuePick;
					var issueStore;
					var serialHeader;
					if (adjusted <= 0) {
						return;
					}
					issueTotal += adjusted;
					issuePick = createElement(issueDoc, "Pick");
					setAttr(issuePick, "TOT", adjusted);
					setAttr(issuePick, "NoofPack", "");
					issueItem.appendChild(issuePick);
					issueStore = createElement(issueDoc, getAttr(item, "RecNumStatus") === "N" ? "STORE" : "PICK");
					setAttr(issueStore, "LOC", getAttr(loc, "LOC"));
					setAttr(issueStore, "BIN", getAttr(loc, "BIN"));
					setAttr(issueStore, "LOTNO", getAttr(pick, "LOTNO") === "NULL" ? "N/A" : getAttr(pick, "LOTNO"));
					setAttr(issueStore, "INVRECNO", getAttr(pick, "INVRECNO"));
					setAttr(issueStore, "QTYISS", adjusted);
					setAttr(issueStore, "NoofPack", "");
					issuePick.appendChild(issueStore);
					if (getAttr(item, "RecNumStatus") !== "N") {
						serialHeader = createElement(issueDoc, "SERIALHEADER");
						elementChildren(pick, "SERIALDETAILS").forEach(function (serial) {
							var serialNode = createElement(issueDoc, "SERIALDETAILS");
							setAttr(serialNode, "SERIALNO", getAttr(serial, "SERIALNO"));
							setAttr(serialNode, "QTY", toNumber(getAttr(serial, "QTY")) * -1);
							serialHeader.appendChild(serialNode);
						});
						if (serialHeader.childNodes.length) {
							issueStore.appendChild(serialHeader);
						}
					}
				});
			});
			setAttr(issueItem, "ISSQTY", issueTotal);
			setAttr(issueItem, "IVALUE", issueTotal);
		});
	}

	function buildReceiptData() {
		var receiptIsland = xmlIsland("IntReceipt");
		var receiptDoc = xmlDocument(receiptIsland);
		var receiptRoot = xmlRoot(receiptIsland);
		var details;
		var entryNo = 0;
		var receiptTotal = 0;
		resetReceiptRoot(receiptRoot);
		details = createElement(receiptDoc, "Details");
		receiptRoot.appendChild(details);
		elementChildren(xmlRoot(xmlIsland("OutData")), "Item").forEach(function (item) {
			var itemDetail = createElement(receiptDoc, "ItemDetail");
			var itemQty = 0;
			entryNo += 1;
			setAttr(itemDetail, "ItemCode", getAttr(item, "ICode"));
			setAttr(itemDetail, "CLACODE", getAttr(item, "CCode"));
			setAttr(itemDetail, "MRSNO", "N");
			setAttr(itemDetail, "ISSNO", "N");
			setAttr(itemDetail, "ENTRYNO", entryNo);
			setAttr(itemDetail, "UNIT", field("hOrgID") ? field("hOrgID").value : "");
			setAttr(itemDetail, "ITEMNAME", getAttr(item, "IName"));
			setAttr(itemDetail, "UOM", "");
			setAttr(itemDetail, "ATTRIBUTELIST", "");
			setAttr(itemDetail, "RefNo", "");
			setAttr(itemDetail, "RefQty", "");
			setAttr(itemDetail, "RECEIPTNUM", getAttr(item, "RecNumStatus"));
			setAttr(itemDetail, "BYPRODUCT", "P");
			details.appendChild(itemDetail);
			elementChildren(item, "LOCDET").forEach(function (loc) {
				elementChildren(loc, "LotSerial").forEach(function (lotSerial) {
					var qty = toNumber(getAttr(lotSerial, "QTY")) - toNumber(getAttr(lotSerial, "TARE"));
					var storage = createElement(receiptDoc, "STORAGE");
					itemQty += qty;
					receiptTotal += qty;
					setAttr(storage, "STORE", getAttr(loc, "LOC"));
					setAttr(storage, "BIN", getAttr(loc, "BIN") === "0" ? "NULL" : getAttr(loc, "BIN"));
					setAttr(storage, "APPLICABLE", "IN");
					setAttr(storage, "MONTHYEAR", today());
					setAttr(storage, "QTY", qty);
					setAttr(storage, "STORAGEVALUE", getAttr(lotSerial, "IVALUE"));
					storage.appendChild(lotSerial.cloneNode(true));
					itemDetail.appendChild(storage);
				});
			});
			setAttr(itemDetail, "QTY", itemQty);
		});
		return receiptTotal;
	}

	window.CheckSubmit = function () {
		var receiptTotal;
		if (!validateRows()) {
			return false;
		}
		applyRowsToOutData();
		buildIssueData();
		receiptTotal = buildReceiptData();
		syncPost("XMLSave.asp?SessionFlag=true&Name=PhyAdjData", serializeXml(xmlIsland("OutData")));
		syncPost("XMLSave.asp?SessionFlag=true&Name=mrsIssueData", serializeXml(xmlIsland("IssueData")));
		if (receiptTotal > 0) {
			syncPost("XMLSave.asp?SessionFlag=true&Name=ReceiptLotData", serializeXml(xmlIsland("IntReceipt")));
		}
		if (field("B7")) {
			field("B7").disabled = true;
		}
		form().action = "stkMgmtPAInsert.asp";
		form().submit();
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
}(window, document));
