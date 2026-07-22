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

	function setText(id, value) {
		var element = byId(id);
		if (element) {
			element.textContent = value || "";
			element.innerText = value || "";
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

	function clearXml(name) {
		var root = xmlRoot(xmlIsland(name));
		removeChildren(root);
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

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return null;
		}
		return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
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

	function locDetNodes(root) {
		var result = [];
		elementChildren(root, "Item").forEach(function (item) {
			elementChildren(item, "LOCDET").forEach(function (loc) {
				result.push({ item: item, loc: loc });
			});
		});
		return result;
	}

	function pickNodes(root) {
		var result = [];
		locDetNodes(root).forEach(function (entry) {
			elementChildren(entry.loc, "PICK").forEach(function (pick) {
				result.push({ item: entry.item, loc: entry.loc, pick: pick });
			});
		});
		return result;
	}

	function storeText(loc) {
		var locName = getAttr(loc, "LOCNAME");
		var bin = getAttr(loc, "BIN");
		var binName = getAttr(loc, "BINNAME");
		if (bin && bin !== "0") {
			return locName + " -- " + bin + " [" + binName + "]";
		}
		return locName + " [" + getAttr(loc, "STOCK") + "]";
	}

	function storeValue(loc) {
		return getAttr(loc, "LOC") + ":" + (getAttr(loc, "BIN") || "0");
	}

	function addDestinationOptions(select, sourceLoc) {
		locDetNodes(xmlRoot(xmlIsland("OutData"))).forEach(function (entry) {
			if (getAttr(entry.loc, "LOC") === getAttr(sourceLoc, "LOC") && getAttr(entry.loc, "BIN") === getAttr(sourceLoc, "BIN")) {
				return;
			}
			select.options[select.options.length] = new Option(storeText(entry.loc), storeValue(entry.loc));
		});
	}

	function rowByElement(element) {
		var index = Number(element && element.getAttribute("data-row-index") || 0);
		return rows[index - 1] || null;
	}

	function findPickFromToken(token) {
		var parts = String(token || "").split("AAAA");
		var itemCode = parts[2] || "";
		var classCode = parts[3] || "";
		var lotNo = parts[4] || "";
		var locNo = parts[6] || "";
		var binNo = parts[7] || "";
		var invRecNo = parts[8] || "";
		var found = null;
		pickNodes(xmlRoot(xmlIsland("OutData"))).some(function (entry) {
			if (getAttr(entry.item, "ICode") === itemCode &&
					getAttr(entry.item, "CCode") === classCode &&
					getAttr(entry.loc, "LOC") === locNo &&
					getAttr(entry.loc, "BIN") === binNo &&
					(getAttr(entry.pick, "LOTNO") === lotNo || !lotNo && getAttr(entry.pick, "LOTNO") === "NULL") &&
					getAttr(entry.pick, "INVRECNO") === invRecNo) {
				found = entry.pick;
				return true;
			}
			return false;
		});
		return found;
	}

	function renderPickRow(table, item, loc, pick, uomDecimal) {
		var recNumStatus = getAttr(item, "RecNumStatus");
		var row = table.insertRow(table.rows.length);
		var rowInfo;
		var qtyInput;
		var serialIcon;
		var serialLink;
		var select;
		var token;
		var lotNo = getAttr(pick, "LOTNO");

		rowInfo = {
			item: item,
			loc: loc,
			pick: pick,
			qtyInput: null,
			stockInput: null,
			targetSelect: null
		};
		rows.push(rowInfo);

		appendCell(row, "ExcelSerial", "center", rows.length);
		appendCell(row, "ExcelDisplayCell", "center", getAttr(item, "IName"));
		appendCell(row, "ExcelDisplayCell", "center", getAttr(loc, "BIN") !== "0" && getAttr(loc, "BIN") !== "" ? getAttr(loc, "LOCNAME") + "-" + getAttr(loc, "BIN") : getAttr(loc, "LOCNAME"));
		appendCell(row, "ExcelDisplayCell", "left", makeInput("txtLot" + rows.length, lotNo === "NULL" ? "-" : lotNo, "FormelemRead", 30, true)).width = "150";
		rowInfo.stockInput = makeInput("txtStk" + rows.length, getAttr(pick, "QTYSTK"), "FormelemRead", 12, true);
		appendCell(row, "ExcelDisplayCell", null, rowInfo.stockInput).width = "10";

		qtyInput = makeInput("txtQtyA" + getAttr(loc, "LOC") + "A" + getAttr(loc, "BIN"), getAttr(pick, "QTYISS") || "0", recNumStatus === "N" ? "Formelem" : "Formelem", 12, recNumStatus !== "N");
		qtyInput.setAttribute("data-row-index", rows.length);
		qtyInput.maxLength = 10;
		qtyInput.onchange = function () {
			window.CheckQty(getAttr(pick, "QTYSTK"), this);
		};
		qtyInput.onkeypress = function () {
			if (typeof window.DoKeyPress === "function") {
				window.DoKeyPress(uomDecimal || "3", 7, 3);
			}
		};
		rowInfo.qtyInput = qtyInput;
		appendCell(row, "ExcelInputCell", null, qtyInput).width = "10";

		serialLink = document.createElement("a");
		serialLink.href = "#";
		serialIcon = document.createElement("img");
		serialIcon.border = "0";
		serialIcon.src = "../../assets/images/iTMS%20Icons/Entry.gif";
		serialIcon.width = 15;
		serialIcon.height = 15;
		serialIcon.setAttribute("data-row-index", rows.length);
		if (recNumStatus !== "N") {
			token = [
				"btn",
				getAttr(item, "Unit"),
				getAttr(item, "ICode"),
				getAttr(item, "CCode"),
				lotNo,
				getAttr(pick, "QTYSTK"),
				getAttr(loc, "LOC"),
				getAttr(loc, "BIN"),
				getAttr(pick, "INVRECNO"),
				uomDecimal || "3"
			].join("AAAA");
			serialIcon.name = token;
			serialIcon.alt = "Serial Details";
			serialIcon.onclick = function () {
				return window.CheckLot(this);
			};
		} else {
			serialIcon.name = "btnAAAA";
			serialIcon.alt = "";
		}
		serialLink.appendChild(serialIcon);
		appendCell(row, "ExcelDisplayCell", "center", serialLink).width = "10";

		select = document.createElement("select");
		select.name = "selSTA" + rows.length;
		select.className = "FormElem";
		addDestinationOptions(select, loc);
		rowInfo.targetSelect = select;
		appendCell(row, "ExcelFieldCell", "center", select).width = "10";
	}

	function uomDecimalFor(item) {
		var uom = elementChildren(item, "UOM")[0];
		return uom ? getAttr(uom, "UoMDecimal") : "3";
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
		xhr = syncGet("itmStoreXMLSelectNew.asp");
		if (trim(xhr.responseText)) {
			loadXmlIsland("OutData", xhr.responseText);
			window.DisplayDetails();
			window.popClaDisplay();
		}
		return false;
	};

	window.CheckLot = function (obj) {
		var row = rowByElement(obj);
		openDialog("stkMgmtSTPoP.asp?sTemp=" + encodeURIComponent(obj.name || "") + "&sValue=", xmlIsland("OutData"), "dialogHeight:310px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No", function () {
			var pick = findPickFromToken(obj.name);
			if (row && row.qtyInput && pick) {
				row.qtyInput.value = getAttr(pick, "QTYISS") || "0";
			}
		});
		return false;
	};

	window.DisplayDetails = function () {
		var root = xmlRoot(xmlIsland("OutData"));
		var table = byId("tblData");
		clearTable();
		elementChildren(root, "Item").forEach(function (item) {
			var uomDecimal = uomDecimalFor(item);
			elementChildren(item, "LOCDET").forEach(function (loc) {
				elementChildren(loc, "PICK").forEach(function (pick) {
					setAttr(pick, "QTYISS", getAttr(pick, "QTYISS") || "");
					setAttr(pick, "STORE", getAttr(pick, "STORE") || "");
					renderPickRow(table, item, loc, pick, uomDecimal);
				});
			});
		});
		return false;
	};

	window.CheckQty = function (stockQty, objOrLoc, bin) {
		var row = typeof objOrLoc === "object" ? rowByElement(objOrLoc) : null;
		var qtyInput = row ? row.qtyInput : field("txtQtyA" + objOrLoc + "A" + bin);
		if (qtyInput && toNumber(qtyInput.value) > toNumber(stockQty)) {
			alert("Enter Issue Qty which is less than Stock Qty");
			qtyInput.value = "0";
		}
		return false;
	};

	window.popClaDisplay = function () {
		updateUomText();
		return false;
	};

	function validateRows() {
		if (!rows.length) {
			return false;
		}
		if (!rows[0].targetSelect || rows[0].targetSelect.options.length < 1) {
			return false;
		}
		for (var i = 0; i < rows.length; i += 1) {
			if (trim(rows[i].qtyInput.value) === "") {
				alert("Enter Quantity");
				rows[i].qtyInput.select();
				return false;
			}
			if (toNumber(rows[i].qtyInput.value) > toNumber(rows[i].stockInput.value)) {
				alert("Transfer Quantity should be equal to or less than Stock Quantity");
				rows[i].qtyInput.select();
				return false;
			}
		}
		return true;
	}

	function applyRowsToOutData() {
		rows.forEach(function (row) {
			var target = String(row.targetSelect.value || "0:0").split(":");
			setAttr(row.pick, "QTYISS", toNumber(row.qtyInput.value));
			setAttr(row.pick, "STORE", target[0] || "");
			setAttr(row.pick, "BIN", target[1] || "0");
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
		setAttr(root, "AUTOACCOUNT", "Y");
	}

	function itemTotalIssued(item) {
		var total = 0;
		elementChildren(item, "LOCDET").forEach(function (loc) {
			elementChildren(loc, "PICK").forEach(function (pick) {
				total += toNumber(getAttr(pick, "QTYISS"));
			});
		});
		return total;
	}

	function serialDetailsFor(item, loc) {
		var details = [];
		elementChildren(loc, "PICK").forEach(function (pick) {
			elementChildren(pick, "SERIALDETAILS").forEach(function (serial) {
				details.push({ pick: pick, serial: serial });
			});
		});
		return details;
	}

	function appendIssuePick(doc, issuePickRoot, pick, loc, recNumStatus) {
		var qty = toNumber(getAttr(pick, "QTYISS"));
		var node;
		var serialHeader;
		if (!qty) {
			return;
		}
		node = createElement(doc, recNumStatus === "N" ? "STORE" : "PICK");
		setAttr(node, "LOC", getAttr(loc, "LOC"));
		setAttr(node, "BIN", getAttr(loc, "BIN"));
		setAttr(node, "LOTNO", getAttr(pick, "LOTNO") || "N/A");
		setAttr(node, "INVRECNO", getAttr(pick, "INVRECNO"));
		setAttr(node, "QTYISS", qty);
		setAttr(node, "NoofPack", "");
		issuePickRoot.appendChild(node);
		if (recNumStatus !== "N") {
			serialHeader = createElement(doc, "SERIALHEADER");
			elementChildren(pick, "SERIALDETAILS").forEach(function (serial) {
				var serialNode = createElement(doc, "SERIALDETAILS");
				setAttr(serialNode, "SERIALNO", getAttr(serial, "SERIALNO"));
				setAttr(serialNode, "QTY", getAttr(serial, "QTY"));
				serialHeader.appendChild(serialNode);
			});
			if (serialHeader.childNodes.length) {
				node.appendChild(serialHeader);
			}
		}
	}

	function buildIssueData() {
		var sourceRoot = xmlRoot(xmlIsland("OutData"));
		var issueIsland = xmlIsland("IssueData");
		var issueDoc = xmlDocument(issueIsland);
		var issueRoot = xmlRoot(issueIsland);
		var entryNo = 0;
		resetIssueRoot(issueRoot);
		elementChildren(sourceRoot, "Item").forEach(function (item) {
			var total = itemTotalIssued(item);
			var issueItem;
			var issuePickRoot;
			entryNo += 1;
			issueItem = createElement(issueDoc, "ITEM");
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
			setAttr(issueItem, "ISSQTY", total);
			setAttr(issueItem, "TRAQTY", "0");
			setAttr(issueItem, "PRQTY", "0");
			setAttr(issueItem, "IVALUE", "0");
			setAttr(issueItem, "ORGCODE", field("hOrgID") ? field("hOrgID").value : "");
			setAttr(issueItem, "MRSNO", "");
			setAttr(issueItem, "MRSDATE", "");
			setAttr(issueItem, "ATTRIBUTELIST", "");
			setAttr(issueItem, "CREATEDBY", field("hUserID") ? field("hUserID").value : "");
			setAttr(issueItem, "CREATEDON", today());
			setAttr(issueItem, "RETURNABLE", "N");
			setAttr(issueItem, "RefNo", "");
			setAttr(issueItem, "ONLYLOT", "");
			setAttr(issueItem, "RETURNITEM", "S");
			setAttr(issueItem, "MatType", "");
			issueRoot.appendChild(issueItem);
			issuePickRoot = createElement(issueDoc, "Pick");
			setAttr(issuePickRoot, "TOT", total);
			setAttr(issuePickRoot, "NoofPack", "");
			issueItem.appendChild(issuePickRoot);
			elementChildren(item, "LOCDET").forEach(function (loc) {
				elementChildren(loc, "PICK").forEach(function (pick) {
					appendIssuePick(issueDoc, issuePickRoot, pick, loc, getAttr(item, "RecNumStatus"));
				});
			});
		});
	}

	function appendReceiptLotSerial(receiptDoc, storage, pick) {
		var serials = elementChildren(pick, "SERIALDETAILS");
		var lotSerial;
		var totalValue = 0;
		if (!serials.length) {
			return;
		}
		lotSerial = createElement(receiptDoc, "LotSerial");
		setAttr(lotSerial, "QTYIN", "N");
		setAttr(lotSerial, "TARE", "0");
		setAttr(lotSerial, "LOT", getAttr(pick, "LOTNO") !== "NULL" ? getAttr(pick, "LOTNO") : "");
		setAttr(lotSerial, "SERIALFROM", "");
		setAttr(lotSerial, "SERIALTO", "");
		setAttr(lotSerial, "TAREWEIGHT", "U");
		setAttr(lotSerial, "QTY", getAttr(pick, "QTYISS"));
		setAttr(lotSerial, "COUNTER", "1");
		setAttr(lotSerial, "STAGE", "");
		setAttr(lotSerial, "ALTGROSS", "0");
		setAttr(lotSerial, "ALTNETT", "0");
		setAttr(lotSerial, "ALTUOM", "select");
		setAttr(lotSerial, "IVALUE", "0");
		setAttr(lotSerial, "AUTOGEN", "AUTO");
		storage.appendChild(lotSerial);
		serials.forEach(function (serial, index) {
			var other = getAttr(serial, "OTHDET").split(":");
			var qty = toNumber(getAttr(serial, "QTY"));
			var rate = toNumber(other[2]);
			var detail = createElement(receiptDoc, "LotSerialDetails");
			totalValue += qty * rate;
			setAttr(detail, "LOTSERIAL", index + 1);
			setAttr(detail, "QTYREC", qty);
			setAttr(detail, "TAREREC", "0");
			setAttr(detail, "SELLINGTYPE", "");
			setAttr(detail, "WEIGHTSTYPE", "0");
			setAttr(detail, "PACKINGTYPE", other[1] || "");
			setAttr(detail, "LOT", other[4] || "");
			setAttr(detail, "SELLINGFORM", "");
			setAttr(detail, "PACKNUMBER", other[0] || "");
			setAttr(detail, "IVALUE", qty * rate);
			setAttr(detail, "ATTRIBUTELIST", other[3] || "");
			setAttr(detail, "NOOFCONE", "0");
			setAttr(detail, "SUBLEVELID", "");
			lotSerial.appendChild(detail);
		});
		setAttr(lotSerial, "IVALUE", totalValue);
		setAttr(storage, "STORAGEVALUE", totalValue);
	}

	function buildReceiptData() {
		var sourceRoot = xmlRoot(xmlIsland("OutData"));
		var receiptIsland = xmlIsland("IntReceipt");
		var receiptDoc = xmlDocument(receiptIsland);
		var receiptRoot = xmlRoot(receiptIsland);
		var details;
		var entryNo = 0;
		resetReceiptRoot(receiptRoot);
		details = createElement(receiptDoc, "Details");
		receiptRoot.appendChild(details);
		elementChildren(sourceRoot, "Item").forEach(function (item) {
			var total = itemTotalIssued(item);
			var itemDetail;
			entryNo += 1;
			itemDetail = createElement(receiptDoc, "ItemDetail");
			setAttr(itemDetail, "ItemCode", getAttr(item, "ICode"));
			setAttr(itemDetail, "CLACODE", getAttr(item, "CCode"));
			setAttr(itemDetail, "QTY", total);
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
				elementChildren(loc, "PICK").forEach(function (pick) {
					var storage;
					var qty = toNumber(getAttr(pick, "QTYISS"));
					if (!qty) {
						return;
					}
					storage = createElement(receiptDoc, "STORAGE");
					setAttr(storage, "STORE", getAttr(pick, "STORE"));
					setAttr(storage, "BIN", getAttr(pick, "BIN") === "0" ? "NULL" : getAttr(pick, "BIN"));
					setAttr(storage, "APPLICABLE", "IN");
					setAttr(storage, "MONTHYEAR", today());
					setAttr(storage, "QTY", qty);
					setAttr(storage, "STORAGEVALUE", "0");
					itemDetail.appendChild(storage);
					appendReceiptLotSerial(receiptDoc, storage, pick);
				});
			});
		});
	}

	window.CheckSubmit = function () {
		if (!validateRows()) {
			return false;
		}
		applyRowsToOutData();
		buildIssueData();
		buildReceiptData();
		syncPost("XMLSave.asp?SessionFlag=true&Name=StockTransferData", serializeXml(xmlIsland("OutData")));
		syncPost("XMLSave.asp?SessionFlag=true&Name=mrsIssueData", serializeXml(xmlIsland("IssueData")));
		syncPost("XMLSave.asp?SessionFlag=true&Name=ReceiptLotData", serializeXml(xmlIsland("IntReceipt")));
		if (field("B7")) {
			field("B7").disabled = true;
		}
		form().action = "StkTransferInsert.asp";
		form().submit();
		return false;
	};
}(window, document));
