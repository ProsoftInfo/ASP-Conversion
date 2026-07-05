(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms["formname"] || document.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		var elements;
		var target;
		var index;
		if (!frm || !frm.elements) {
			return null;
		}
		if (frm.elements[name]) {
			return frm.elements[name];
		}
		elements = frm.elements;
		target = String(name).toLowerCase();
		for (index = 0; index < elements.length; index += 1) {
			if (String(elements[index].name || elements[index].id || "").toLowerCase() === target) {
				return elements[index];
			}
		}
		return null;
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
	}

	function valueOf(name, fallback) {
		var item = field(name) || byId(name);
		return item ? item.value : fallback || "";
	}

	function setValue(name, value) {
		var item = field(name) || byId(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function setText(id, value) {
		var item = byId(id);
		if (item) {
			item.textContent = value == null ? "" : String(value);
		}
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.upgradeXmlIslands) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
	}

	function xmlObject(nameOrObject) {
		if (typeof nameOrObject !== "string") {
			return nameOrObject;
		}
		ensureCompat();
		return window[nameOrObject] || document[nameOrObject] || byId(nameOrObject) || null;
	}

	function xmlDocument(nameOrObject) {
		var object = xmlObject(nameOrObject);
		return object && object.XMLDocument || object && object._doc || object && object.nodeType === 9 && object || null;
	}

	function xmlRoot(nameOrObject) {
		var object = xmlObject(nameOrObject);
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function serializeXml(nameOrObject) {
		var doc = xmlDocument(nameOrObject);
		var root = xmlRoot(nameOrObject);
		if (doc) {
			return new XMLSerializer().serializeToString(doc);
		}
		return root ? new XMLSerializer().serializeToString(root) : "";
	}

	function loadXml(name, text) {
		var object = xmlObject(name);
		var doc;
		if (!object || !trim(text)) {
			return false;
		}
		if (typeof object.loadXML === "function") {
			return object.loadXML(text);
		}
		doc = xmlDocument(object);
		if (doc && typeof doc.loadXML === "function") {
			return doc.loadXML(text);
		}
		return false;
	}

	function attr(node, nameOrIndex) {
		var item;
		if (!node || !node.attributes) {
			return "";
		}
		if (typeof nameOrIndex === "number") {
			item = node.attributes.item(nameOrIndex);
			return item ? item.nodeValue : "";
		}
		return node.getAttribute(nameOrIndex) || "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function childElements(node) {
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1;
		});
	}

	function selectNodes(context, expression) {
		var doc;
		var result;
		var nodes = [];
		var index;
		if (!context) {
			return nodes;
		}
		if (typeof context.selectNodes === "function") {
			return Array.prototype.slice.call(context.selectNodes(expression));
		}
		doc = context.nodeType === 9 ? context : context.ownerDocument;
		if (!doc || !doc.evaluate) {
			return nodes;
		}
		try {
			result = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
			for (index = 0; index < result.snapshotLength; index += 1) {
				nodes.push(result.snapshotItem(index));
			}
		} catch (ignore) {}
		return nodes;
	}

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function openDialog(url, args, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
		}
		return window.open(url, "_blank", "height=500,width=500,resizable=no,status=no,scrollbars=yes");
	}

	function continueSelection(url, args, features, callback) {
		openDialog(url, args, features, function (value) {
			var root = xmlRoot(value) || value && value.nodeType === 1 && value || null;
			var action = String(attr(root, "Action")).toUpperCase();
			var query = trim(attr(root, "PassQuery"));
			var base = String(url).replace(/\?.*$/, "");
			if (action && action !== "DONE" && action !== "CLOSE" && query) {
				continueSelection(base + "?" + query, args, features, callback);
				return;
			}
			callback(root || value);
		});
	}

	function popupSize(code, fallback) {
		var value = window.GetWindowSizeForPopup ? window.GetWindowSizeForPopup(code) : fallback;
		var parts = String(value || fallback || "").split(":");
		return {
			program: parts[0] || "PartySelection.asp",
			height: parts[1] || "500",
			width: parts[2] || "500"
		};
	}

	function ShowVouch(transNo, transType, cndnType) {
		var type = trim(transType);
		var url = "";
		if (type === "PJR") {
			url = "PurchaseVouchView_San.asp?TransNo=" + encodeURIComponent(transNo);
		} else if (type === "SJR") {
			url = "SalesVouchView_San.asp?TransNo=" + encodeURIComponent(transNo);
		} else if (type === "CNR" || type === "DNR") {
			if (trim(cndnType) === "OT") {
				url = "CNDNOthVouchView_San.asp?TransNo=" + encodeURIComponent(transNo);
			} else {
				url = "CNDNVouchView_San.asp?TransNo=" + encodeURIComponent(transNo) + "&BankType=" + encodeURIComponent(cndnType);
			}
		} else if (type === "BAP" || type === "BAR") {
			url = "BankVouchView_San.asp?TransNo=" + encodeURIComponent(transNo);
		} else if (type === "CAP" || type === "CAR") {
			url = "CashVouchView_San.asp?TransNo=" + encodeURIComponent(transNo);
		} else if (type === "GJR") {
			url = "GJVouchView_San.asp?TransNo=" + encodeURIComponent(transNo);
		}
		if (url) {
			openDialog(url, "", "dialogHeight:600px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No", function () {});
		}
		return false;
	}

	function selectedPartyFilter() {
		var type = valueOf("SelPartyType");
		var subType = valueOf("hPArtySubType");
		if (type === "CR") {
			return "CR?" + subType + "? ";
		}
		if (type === "DR") {
			return "DR?" + subType + "? ";
		}
		return "";
	}

	function finishPartySelection(root) {
		var names = [];
		var codes = [];
		var subTypes = [];
		var partyTypes = [];
		childElements(root).forEach(function (entry) {
			names.push(attr(entry, "RetField0"));
			codes.push(attr(entry, "RetField1"));
			subTypes.push(attr(entry, "RetField4"));
			partyTypes.push(attr(entry, "RetField3"));
		});
		if (names.length) {
			setValue("txtPartyName", names.join(","));
			setValue("hPartyCode", codes.join(","));
			setValue("hPartySubTypeCheck", subTypes.join(":"));
			setValue("hPartyTypeCheck", partyTypes.join(":"));
		} else {
			setValue("txtPartyName", "");
			setValue("hPartyCode", "0");
			setValue("hPartySubTypeCheck", "0");
			setValue("hPartyTypeCheck", "0");
			return;
		}
		if (valueOf("SelPartySubType") === "A") {
			if (byId("divCRDRSubType")) {
				byId("divCRDRSubType").style.display = "block";
			}
			loadAllPartySubTypes();
		}
	}

	function PartySelection() {
		var size = popupSize("2", "PartySelection.asp:500:500");
		var orgId = valueOf("hOrgID");
		var partyType = selectedPartyFilter();
		var url = "../../Common/" + size.program + "?orgid=" + encodeURIComponent(orgId) + "&Party=" + encodeURIComponent(partyType);
		var features = "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No";
		continueSelection(url, xmlObject("PartyData"), features, finishPartySelection);
		return false;
	}

	function rowCheckbox(name, value) {
		var input = document.createElement("input");
		input.type = "checkbox";
		input.name = name;
		if (value !== undefined) {
			input.value = value;
		}
		return input;
	}

	function addSubTypeRow(table, checkboxName, subtype, partyType, label) {
		var row = table.insertRow(-1);
		var cell = row.insertCell(-1);
		cell.appendChild(rowCheckbox(checkboxName, subtype + ":" + partyType));
		cell.className = "ExcelDisplayCell";
		cell.align = "center";

		cell = row.insertCell(-1);
		cell.textContent = label;
		cell.className = "ExcelDisplayCell";
		cell.align = "Left";

		cell = row.insertCell(-1);
		cell.appendChild(rowCheckbox(partyType === "CR" ? "ChkPartySubType" : "ChkPartySubTypeDR"));
		cell.className = "ExcelDisplayCell";
		cell.align = "center";
	}

	function ClearTable() {
		var crTable = byId("CRTable");
		var drTable = byId("DRTable");
		while (crTable && crTable.rows.length > 2) {
			crTable.deleteRow(2);
		}
		while (drTable && drTable.rows.length > 2) {
			drTable.deleteRow(2);
		}
	}

	function selectedCheckValues() {
		return trim(valueOf("hPartySubTypeCheck")).split(":").filter(function (value) {
			return value !== "";
		});
	}

	function DisplaySubType() {
		var root = xmlRoot("PartyAllTypes");
		var crTable = byId("CRTable");
		var drTable = byId("DRTable");
		var selected = selectedCheckValues();
		var crCount = 0;
		var drCount = 0;
		ClearTable();
		if (!root || !crTable || !drTable) {
			return false;
		}
		setAttr(root, "PartySubType", valueOf("hPartySubTypeCheck"));
		setAttr(root, "PartyType", valueOf("hPartyTypeCheck"));
		childElements(root).forEach(function (node) {
			var subtypeData;
			var partyType;
			var partySubType;
			var label;
			if (String(node.nodeName).toLowerCase() !== "party") {
				return;
			}
			subtypeData = String(attr(node, "SubType")).split("|");
			partyType = subtypeData[0];
			partySubType = subtypeData[1];
			label = trim(node.textContent || node.text || "");
			if (selected.some(function (value) { return Number(value) === Number(partySubType); })) {
				setAttr(node, "Check", "RCheck");
			}
			if (partyType === "CR") {
				crCount += 1;
				addSubTypeRow(crTable, "ChkCR" + crCount, partySubType, partyType, label);
			} else {
				drCount += 1;
				addSubTypeRow(drTable, "ChkDR" + drCount, partySubType, partyType, label);
			}
		});
		setValue("hCRCtr", crCount);
		setValue("hDRCtr", drCount);
		return true;
	}

	function loadAllPartySubTypes() {
		var xhr = syncGet("/Common/PartySubType.asp?OrgCode=" + encodeURIComponent(valueOf("hOrgID")));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("PartyAllTypes", serializeXml(xhr.responseXML));
			DisplaySubType();
		} else if (trim(xhr.responseText)) {
			loadXml("PartyAllTypes", xhr.responseText);
			DisplaySubType();
		}
	}

	function checkedField(name) {
		var item = field(name);
		return !!(item && item.checked);
	}

	function markSelectedSubtype(partyType, partySubType) {
		var root = xmlRoot("PartyAllTypes");
		var data = partyType + "|" + partySubType;
		selectNodes(root, "//Party[@SubType='" + data + "']").forEach(function (node) {
			setAttr(node, "Check", "LCheck");
		});
	}

	function collectSubtype(prefix, countName, partyType) {
		var selected = [];
		var count = Number(valueOf(countName, 0));
		var index;
		var item;
		for (index = 1; index <= count; index += 1) {
			item = field(prefix + index);
			if (item && item.checked) {
				selected.push(String(item.value || "").split(":")[0]);
				markSelectedSubtype(partyType, String(item.value || "").split(":")[0]);
			}
		}
		return selected;
	}

	function submitStatement() {
		form().submit();
	}

	function Validate() {
		var root = xmlRoot("PartyAllTypes");
		var crSubTypes = collectSubtype("ChkCR", "hCRCtr", "CR");
		var drSubTypes = collectSubtype("ChkDR", "hDRCtr", "DR");
		var formValue = valueOf("SelPartyType") + ":" + valueOf("SelPartySubType") + ":" + valueOf("hPArtySubType") + ":" + valueOf("hPartyCode") + ":" + valueOf("SelFromMonth") + ":" + valueOf("SelToMonth");
		formValue += ":" + crSubTypes.join(",");
		formValue += ":" + (crSubTypes.length ? "CR" : "");
		formValue += ":" + drSubTypes.join(",");
		formValue += ":" + (drSubTypes.length ? "DR" : "");
		setValue("hFormVal", formValue);
		if (selectNodes(root, "//Party").length > 0) {
			openDialog("SalTrPartySubTypePopup.asp", xmlObject("PartyAllTypes"), "dialogHeight:300px;dialogWidth:535px;center:Yes;help:No;resizable:No;status:No", submitStatement);
		} else {
			submitStatement();
		}
		return false;
	}

	function SelectPartyHead() {
		var ledgerType = valueOf("SelPartyType");
		var orgId;
		var partyType;
		var url;
		if (valueOf("SelPartySubType") === "A") {
			setText("spPartyHead", "");
			setValue("hPartyCode", "");
			return false;
		}
		if (ledgerType === "S") {
			alert("Select Party type");
			if (field("SelPartySubType")) {
				field("SelPartySubType").selectedIndex = 0;
			}
			if (field("SelPartyType")) {
				field("SelPartyType").focus();
			}
			return false;
		}
		orgId = valueOf("hOrgID");
		partyType = ledgerType + "?" + (valueOf("hPartyCode") === "CR" ? "4" : "3") + "? ";
		url = "../../Common/PartySubTypeSelection.asp?orgId=" + encodeURIComponent(orgId) + "&Party=" + encodeURIComponent(partyType) + "&hSelectMode=R";
		continueSelection(url, xmlObject("PartySubTypeData"), "dialogHeight:500px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (root) {
			var names = [];
			var subTypes = [];
			childElements(root).forEach(function (entry) {
				names.push(attr(entry, "RetField0"));
				subTypes.push(attr(entry, "RetField1"));
			});
			if (subTypes.length) {
				setValue("hPArtySubType", subTypes.join(","));
				setText("spPartyHead", names.join(","));
			} else {
				setText("spPartyHead", "");
				setValue("hPArtySubType", "0");
			}
		});
		return false;
	}

	window.ShowVouch = ShowVouch;
	window.PartySelection = PartySelection;
	window.DisplaySubType = DisplaySubType;
	window.ClearTable = ClearTable;
	window.Validate = Validate;
	window.SelectPartyHead = SelectPartyHead;
}(window, document));
